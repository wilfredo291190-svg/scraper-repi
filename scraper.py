import json
import re
import glob
from datetime import datetime, timezone, timedelta
from pathlib import Path
import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser as HTMLParser

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "application/json, text/html, */*",
    "Accept-Language": "es-PA,es;q=0.9,en;q=0.8",
}

CODE_ALIASES = {
    "RT16N6CDX": "RT3N461NACA",
}

def clean_code(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).upper())

def clean_barcode(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^0-9]", "", s)

def load_catalog():
    excel_files = glob.glob("data/*.xlsx") + glob.glob("*.xlsx")
    if not excel_files:
        raise FileNotFoundError("No se encontro ningun archivo Excel de catalogo en data/ ni en la raiz.")
    
    filepath = excel_files[0]
    print(f"Cargando catalogo desde: {filepath}")
    df = pd.read_excel(filepath, sheet_name="Lista de precios", header=5)
    df = df.dropna(subset=["Código"]).copy()
    
    catalog = []
    for _, row in df.iterrows():
        code_raw = str(row["Código"]).strip()
        code_norm = clean_code(code_raw)
        barcode = clean_barcode(row.get("BarCode", ""))
        
        valid_codes = {code_norm}
        for alias, target in CODE_ALIASES.items():
            if target == code_raw:
                valid_codes.add(clean_code(alias))

        catalog.append({
            "marca": str(row.get("Marca", "HISENSE")).strip().upper(),
            "categoria": str(row.get("Categoría", "")).strip(),
            "codigo": code_raw,
            "codes_norm": list(valid_codes),
            "barcode": barcode,
            "descripcion": str(row.get("Descripción", "")).strip(),
            "precio_regular": float(row["Precio Regular (USD)"]) if pd.notna(row["Precio Regular (USD)"]) else None,
            "map_usd": float(row["MAP (USD)"]) if pd.notna(row["MAP (USD)"]) else None,
            "disponibilidad": str(row.get("Disponibilidad", "")).strip(),
        })
    return catalog

def scrape_panafoto(client: httpx.Client):
    items = []
    try:
        r = client.get("https://www.panafoto.com/catalogsearch/result/?q=Hisense", timeout=25)
        html = r.text
        
        # Buscar llaves de Algolia en el HTML o en scripts referenciados
        app_id_match = re.search(r'"applicationId"\s*:\s*"([A-Z0-9]{10})"', html)
        api_key_match = re.search(r'"apiKey"\s*:\s*"([a-f0-9]{32})"', html)
        index_match = re.search(r'"indexName"\s*:\s*"([a-zA-Z0-9_]+products)"', html)

        if not (app_id_match and api_key_match):
            # Buscar dentro de archivos JS cargados en la pagina
            js_urls = re.findall(r'src="([^"]+\.js[^"]*)"', html)
            for js_u in js_urls:
                if "algolia" in js_u.lower() or "config" in js_u.lower() or "bundle" in js_u.lower():
                    try:
                        r_js = client.get(js_u if js_u.startswith("http") else f"https://www.panafoto.com{js_u}", timeout=15)
                        if not app_id_match:
                            app_id_match = re.search(r'"applicationId"\s*:\s*"([A-Z0-9]{10})"', r_js.text)
                        if not api_key_match:
                            api_key_match = re.search(r'"apiKey"\s*:\s*"([a-f0-9]{32})"', r_js.text)
                        if app_id_match and api_key_match:
                            break
                    except Exception:
                        pass

        if app_id_match and api_key_match:
            app_id = app_id_match.group(1)
            api_key = api_key_match.group(1)
            index_name = index_match.group(1) if index_match else "panafoto_mcprod_cloud_b2c_products"
            algolia_url = f"https://{app_id}-dsn.algolia.net/1/indexes/{index_name}/query"
            algolia_headers = {
                "X-Algolia-Application-Id": app_id,
                "X-Algolia-API-Key": api_key,
                "Content-Type": "application/json",
            }
            res = client.post(algolia_url, headers=algolia_headers, json={"params": "query=Hisense&hitsPerPage=250"}, timeout=20)
            for hit in res.json().get("hits", []):
                price_obj = hit.get("price", {}).get("USD", {})
                price = price_obj.get("default")
                if isinstance(price, dict):
                    price = price.get("default")
                items.append({
                    "tienda": "Panafoto",
                    "sku": str(hit.get("sku", "")),
                    "barcode": clean_barcode(hit.get("ean") or hit.get("barcode") or ""),
                    "titulo": str(hit.get("name", "")),
                    "precio": float(price) if price else None,
                    "url": str(hit.get("url", "")),
                })
            print(f"[Panafoto] Extraidos {len(items)} productos via Algolia.")
            return items

        # Fallback: leer tarjetas HTML si Algolia esta ofuscado
        tree = HTMLParser(html)
        for card in tree.css(".product-item-info, .product-item"):
            title_node = card.css_first(".product-item-link, a.product-item-photo")
            price_node = card.css_first("[data-price-amount]")
            if title_node and price_node:
                items.append({
                    "tienda": "Panafoto",
                    "sku": "",
                    "barcode": "",
                    "titulo": title_node.text(strip=True) or title_node.attributes.get("title", ""),
                    "precio": float(price_node.attributes.get("data-price-amount", 0)),
                    "url": title_node.attributes.get("href", ""),
                })
        print(f"[Panafoto] HTTP {r.status_code} | Extraidos {len(items)} via HTML (Algolia keys no visibles).")
    except Exception as e:
        print(f"[Panafoto] Error: {e}")
    return items

def scrape_shopify_store(client: httpx.Client, store_name: str, base_url: str):
    items = []
    seen_ids = set()
    # Paginar hasta 5 paginas (1,250 productos) para no perder articulos de Multimax/Rodelag
    for page in range(1, 6):
        endpoints = [
            f"{base_url}/collections/hisense/products.json?limit=250&page={page}",
            f"{base_url}/products.json?limit=250&page={page}",
        ]
        found_in_page = False
        for ep in endpoints:
            try:
                r = client.get(ep, timeout=20)
                if r.status_code != 200:
                    continue
                products = r.json().get("products", [])
                if products:
                    found_in_page = True
                for p in products:
                    if "hisense" not in str(p.get("vendor", "")).lower() and "hisense" not in str(p.get("title", "")).lower():
                        continue
                    for v in p.get("variants", []):
                        vid = v.get("id")
                        if vid in seen_ids:
                            continue
                        seen_ids.add(vid)
                        price = float(v["price"]) if v.get("price") else None
                        items.append({
                            "tienda": store_name,
                            "sku": str(v.get("sku") or ""),
                            "barcode": clean_barcode(v.get("barcode") or ""),
                            "titulo": f"{p.get('title', '')} {v.get('title', '')}".strip(),
                            "precio": price,
                            "url": f"{base_url}/products/{p.get('handle', '')}",
                        })
            except Exception:
                pass
        if not found_in_page:
            break
    print(f"[{store_name}] Extraidos {len(items)} variantes via Shopify JSON.")
    return items

def scrape_evision(client: httpx.Client):
    items = scrape_shopify_store(client, "EVision", "https://www.evisionstore.com")
    if items:
        return items
    try:
        # Probar busqueda directa en EVision y diagnosticar estructura
        r = client.get("https://www.evisionstore.com/?s=Hisense&post_type=product", timeout=20)
        tree = HTMLParser(r.text)
        for card in tree.css(".product, .product-item, .type-product"):
            title_node = card.css_first("h2, .woocommerce-loop-product__title, .product-title a, a")
            price_node = card.css_first(".price ins .amount, .price .amount, .price")
            if title_node and price_node:
                raw_price = re.sub(r"[^0-9.]", "", price_node.text(strip=True))
                if raw_price:
                    items.append({
                        "tienda": "EVision",
                        "sku": "",
                        "barcode": "",
                        "titulo": title_node.text(strip=True),
                        "precio": float(raw_price),
                        "url": title_node.attributes.get("href", "https://www.evisionstore.com"),
                    })
        print(f"[EVision] HTTP {r.status_code} | Extraidos {len(items)} productos. (Title page: {tree.css_first('title').text(strip=True) if tree.css_first('title') else 'N/A'})")
    except Exception as e:
        print(f"[EVision] Error: {e}")
    return items

def scrape_novey(client: httpx.Client):
    items = []
    try:
        search_url = "https://www.novey.com.pa/catalogsearch/result/?q=Hisense"
        r = client.get(search_url, timeout=25)
        tree = HTMLParser(r.text)
        # Buscar en bloques JSON-LD o scripts incrustados de Magento/VTEX
        for script in tree.css("script"):
            txt = script.text()
            if "Hisense" in txt and ("price" in txt or "final_price" in txt):
                # Extraer pares SKU/nombre/precio si vienen en JSON embebido
                pass
        for card in tree.css(".product-item-info, .product-item, [data-product-id]"):
            title_node = card.css_first(".product-item-link, a[title]")
            price_node = card.css_first("[data-price-amount], .price")
            if title_node and price_node:
                p_val = price_node.attributes.get("data-price-amount")
                if not p_val:
                    p_val = re.sub(r"[^0-9.]", "", price_node.text(strip=True))
                if p_val:
                    items.append({
                        "tienda": "Novey",
                        "sku": "",
                        "barcode": "",
                        "titulo": title_node.text(strip=True) or title_node.attributes.get("title", ""),
                        "precio": float(p_val),
                        "url": title_node.attributes.get("href", ""),
                    })
        page_title = tree.css_first("title").text(strip=True) if tree.css_first("title") else "N/A"
        print(f"[Novey] HTTP {r.status_code} | Extraidos {len(items)} productos. (Title page: {page_title})")
    except Exception as e:
        print(f"[Novey] Error: {e}")
    return items

def match_products(catalog, scraped_items):
    stores = ["Panafoto", "Novey", "Rodelag", "Multimax", "EVision"]
    results = []
    total_matches = 0

    for sku_row in catalog:
        row_out = {
            "categoria": sku_row["categoria"],
            "codigo": sku_row["codigo"],
            "barcode": sku_row["barcode"],
            "descripcion": sku_row["descripcion"],
            "precio_regular": sku_row["precio_regular"],
            "map_usd": sku_row["map_usd"],
            "disponibilidad": sku_row["disponibilidad"],
            "tiendas": {s: None for s in stores}
        }

        for s in stores:
            store_items = [i for i in scraped_items if i["tienda"] == s and i["precio"] and i["precio"] > 0]
            matched = None

            if sku_row["barcode"]:
                for item in store_items:
                    if item["barcode"] and item["barcode"] == sku_row["barcode"]:
                        matched = item
                        break

            if not matched:
                for item in store_items:
                    haystack = clean_code(item["sku"] + " " + item["titulo"])
                    if any(c in haystack for c in sku_row["codes_norm"] if len(c) >= 4):
                        matched = item
                        break

            if matched:
                total_matches += 1
                row_out["tiendas"][s] = {
                    "precio": matched["precio"],
                    "url": matched["url"],
                    "titulo_tienda": matched["titulo"]
                }

        results.append(row_out)
    print(f"Total de coincidencias (matches) cruzadas contra tu Excel: {total_matches}")
    return results

def main():
    catalog = load_catalog()
    with httpx.Client(headers=HEADERS, follow_redirects=True) as client:
        all_scraped = []
        all_scraped.extend(scrape_panafoto(client))
        all_scraped.extend(scrape_shopify_store(client, "Multimax", "https://www.multimax.net"))
        all_scraped.extend(scrape_shopify_store(client, "Rodelag", "https://rodelag.com"))
        all_scraped.extend(scrape_evision(client))
        all_scraped.extend(scrape_novey(client))

    matched_data = match_products(catalog, all_scraped)
    
    panama_tz = timezone(timedelta(hours=-5))
    now_str = datetime.now(panama_tz).strftime("%Y-%m-%d %H:%M:%S")

    output = {
        "actualizado": now_str,
        "total_skus": len(matched_data),
        "productos": matched_data
    }

    Path("data/precios_dashboard.json").write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    print(f"Proceso completado. Archivo data/precios_dashboard.json actualizado ({now_str}).")

if __name__ == "__main__":
    main()
