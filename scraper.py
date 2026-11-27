import json
import re
import glob
from datetime import datetime, timezone, timedelta
from pathlib import Path
import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser as HTMLParser

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    "Accept": "application/json, text/html, */*",
    "Accept-Language": "es-PA,es;q=0.9,en;q=0.8",
}

# Ranking oficial de clientes en orden de importancia
STORE_ORDER = [
    "Multimax",
    "Novey",
    "Rodelag",
    "Panafoto",
    "EVision",
    "Do It Center",
    "Titan",
    "Spiegel",
    "Selecto",
    "Creditos Mundiales",
    "Unica Panama",
    "Raenco",
]

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

def load_all_catalogs():
    """Lee todos los archivos .xlsx dentro de data/ y consolida los SKUs."""
    excel_files = sorted(glob.glob("data/*.xlsx"))
    if not excel_files:
        excel_files = sorted(glob.glob("*.xlsx"))
    if not excel_files:
        raise FileNotFoundError("No se encontro ningun archivo Excel de catalogo en data/.")

    catalog = []
    seen_codes = set()

    for filepath in excel_files:
        print(f"Cargando catalogo: {filepath}")
        xls = pd.ExcelFile(filepath)
        sheet = "Lista de precios" if "Lista de precios" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(filepath, sheet_name=sheet, header=5)
        if "Código" not in df.columns:
            continue
        df = df.dropna(subset=["Código"]).copy()

        for _, row in df.iterrows():
            code_raw = str(row["Código"]).strip()
            if not code_raw or code_raw.lower() == "nan":
                continue
            code_norm = clean_code(code_raw)
            if code_norm in seen_codes:
                continue
            seen_codes.add(code_norm)

            barcode = clean_barcode(row.get("BarCode", ""))
            valid_codes = {code_norm}
            for alias, target in CODE_ALIASES.items():
                if target == code_raw:
                    valid_codes.add(clean_code(alias))

            reg_price = None
            if "Precio Regular (USD)" in df.columns and pd.notna(row.get("Precio Regular (USD)")):
                try:
                    reg_price = float(row["Precio Regular (USD)"])
                except Exception:
                    reg_price = None

            map_price = None
            if "MAP (USD)" in df.columns and pd.notna(row.get("MAP (USD)")):
                try:
                    map_price = float(row["MAP (USD)"])
                except Exception:
                    map_price = None

            catalog.append({
                "marca": str(row.get("Marca", "")).strip().upper(),
                "categoria": str(row.get("Categoría", "")).strip().upper(),
                "codigo": code_raw,
                "codes_norm": list(valid_codes),
                "barcode": barcode,
                "descripcion": str(row.get("Descripción", "")).strip(),
                "precio_regular": reg_price,
                "map_usd": map_price,
                "disponibilidad": str(row.get("Disponibilidad", "")).strip(),
            })

    brands = sorted({c["marca"] for c in catalog if c["marca"]})
    print(f"Catalogo consolidado: {len(catalog)} SKUs unicos | Marcas: {brands}")
    return catalog, brands

def extract_price_float(raw_str):
    if raw_str is None:
        return None
    cleaned = re.sub(r"[^0-9.]", "", str(raw_str).replace(",", ""))
    if not cleaned:
        return None
    try:
        val = float(cleaned)
        return val if val > 0 else None
    except Exception:
        return None

def scrape_shopify_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list):
    """Extrae variantes de Shopify paginando colecciones por marca y catalogo general."""
    items = []
    seen_ids = set()
    brand_tokens = [b.lower() for b in brands]

    endpoints_templates = [f"{base_url}/products.json?limit=250&page={{page}}"]
    for b in brand_tokens:
        endpoints_templates.insert(0, f"{base_url}/collections/{b}/products.json?limit=250&page={{page}}")

    for tmpl in endpoints_templates:
        for page in range(1, 6):
            ep = tmpl.format(page=page)
            try:
                r = client.get(ep, timeout=18)
                if r.status_code != 200:
                    break
                data = r.json()
                products = data.get("products", [])
                if not products:
                    break
                for p in products:
                    vendor_title = f"{p.get('vendor', '')} {p.get('title', '')}".lower()
                    if not any(b in vendor_title for b in brand_tokens):
                        continue
                    for v in p.get("variants", []):
                        vid = v.get("id")
                        if vid in seen_ids:
                            continue
                        seen_ids.add(vid)
                        price = extract_price_float(v.get("price"))
                        if price:
                            items.append({
                                "tienda": store_name,
                                "sku": str(v.get("sku") or ""),
                                "barcode": clean_barcode(v.get("barcode") or ""),
                                "titulo": f"{p.get('title', '')} {v.get('title', '')}".strip(),
                                "precio": price,
                                "url": f"{base_url}/products/{p.get('handle', '')}",
                            })
            except Exception:
                break
    return items

def scrape_vtex_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list):
    """Consulta API REST de catalogo VTEX por marca."""
    items = []
    seen_skus = set()
    for b in brands:
        for start in range(0, 150, 50):
            url = f"{base_url}/api/catalog_system/pub/products/search/{b.lower()}?_from={start}&_to={start+49}"
            try:
                r = client.get(url, timeout=18)
                if r.status_code not in (200, 206):
                    break
                if not r.headers.get("content-type", "").startswith("application/json"):
                    break
                prods = r.json()
                if not prods:
                    break
                for p in prods:
                    for sku_item in p.get("items", []):
                        item_id = sku_item.get("itemId")
                        if item_id in seen_skus:
                            continue
                        seen_skus.add(item_id)
                        sellers = sku_item.get("sellers", [])
                        price = sellers[0].get("commertialOffer", {}).get("Price") if sellers else None
                        price = extract_price_float(price)
                        if not price:
                            continue
                        ref_ids = " ".join(str(ref.get("Value", "")) for ref in sku_item.get("referenceId", []))
                        items.append({
                            "tienda": store_name,
                            "sku": ref_ids.strip(),
                            "barcode": clean_barcode(sku_item.get("ean", "")),
                            "titulo": f"{p.get('productName', '')} {sku_item.get('name', '')}".strip(),
                            "precio": price,
                            "url": f"{base_url}/{p.get('linkText', '')}/p",
                        })
            except Exception:
                break
    return items

def scrape_woocommerce_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list):
    """Consulta API publica de WooCommerce Store API por marca."""
    items = []
    seen_ids = set()
    for b in brands:
        url = f"{base_url}/wp-json/wc/store/products?search={b}&per_page=100"
        try:
            r = client.get(url, timeout=18)
            if r.status_code != 200:
                continue
            data = r.json()
            if not isinstance(data, list):
                continue
            for p in data:
                pid = p.get("id")
                if pid in seen_ids:
                    continue
                seen_ids.add(pid)
                prices = p.get("prices", {})
                minor = int(prices.get("currency_minor_unit", 2))
                raw_p = prices.get("price")
                price = float(raw_p) / (10 ** minor) if raw_p else None
                if price and price > 0:
                    items.append({
                        "tienda": store_name,
                        "sku": str(p.get("sku", "")),
                        "barcode": "",
                        "titulo": str(p.get("name", "")),
                        "precio": price,
                        "url": str(p.get("permalink", base_url)),
                    })
        except Exception:
            pass
    return items

def scrape_magento_html_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list):
    """Consulta paginas de busqueda Magento 2 / HTML estructurado por marca."""
    items = []
    seen_urls = set()
    for b in brands:
        for page in range(1, 5):
            urls_to_try = [
                f"{base_url}/catalogsearch/result/index/?p={page}&q={b}",
                f"{base_url}/search?q={b}&page={page}",
            ]
            found_in_page = False
            for s_url in urls_to_try:
                try:
                    r = client.get(s_url, timeout=20)
                    if r.status_code != 200:
                        continue
                    tree = HTMLParser(r.text)
                    # 1. JSON-LD ItemList
                    for script in tree.css('script[type="application/ld+json"]'):
                        try:
                            ld = json.loads(script.text())
                            if isinstance(ld, dict) and ld.get("@type") == "ItemList":
                                for el in ld.get("itemListElement", []):
                                    prod = el.get("item", {})
                                    offers = prod.get("offers", {})
                                    price = extract_price_float(offers.get("price") if isinstance(offers, dict) else None)
                                    u = str(prod.get("url", s_url))
                                    if price and u not in seen_urls:
                                        seen_urls.add(u)
                                        found_in_page = True
                                        items.append({
                                            "tienda": store_name,
                                            "sku": str(prod.get("sku", "")),
                                            "barcode": clean_barcode(prod.get("gtin13") or prod.get("gtin") or ""),
                                            "titulo": str(prod.get("name", "")),
                                            "precio": price,
                                            "url": u,
                                        })
                        except Exception:
                            pass
                    # 2. Tarjetas HTML estandar
                    for card in tree.css("li.product-item, div.product-item-info, .product-card, .product, .type-product"):
                        link_node = card.css_first("a.product-item-link, .product-title a, h2 a, a[href]")
                        price_node = card.css_first("[data-price-amount], .price ins .amount, .price .amount, span.price")
                        if link_node and price_node:
                            href = link_node.attributes.get("href", "")
                            if href and not href.startswith("http"):
                                href = base_url.rstrip("/") + "/" + href.lstrip("/")
                            if not href or href in seen_urls:
                                continue
                            p_val = price_node.attributes.get("data-price-amount") or price_node.text(strip=True)
                            price = extract_price_float(p_val)
                            if price:
                                seen_urls.add(href)
                                found_in_page = True
                                items.append({
                                    "tienda": store_name,
                                    "sku": card.attributes.get("data-product-sku", ""),
                                    "barcode": "",
                                    "titulo": link_node.text(strip=True) or link_node.attributes.get("title", ""),
                                    "precio": price,
                                    "url": href,
                                })
                    if found_in_page:
                        break
                except Exception:
                    pass
            if not found_in_page:
                break
    return items

def scrape_panafoto(client: httpx.Client, brands: list):
    """Consulta Algolia en Panafoto (con comillas simples/dobles) o fallback HTML."""
    items = []
    try:
        r = client.get("https://www.panafoto.com/categorias/tv-y-video/tv.html?manufacturer=HISENSE", timeout=22)
        html = r.text
        app_id_m = re.search(r"""['"]applicationId['"]\s*:\s*['"]([A-Z0-9]{10})['"]""", html)
        api_key_m = re.search(r"""['"]apiKey['"]\s*:\s*['"]([a-f0-9]{32})['"]""", html)
        index_m = re.search(r"""['"]indexName['"]\s*:\s*['"]([a-zA-Z0-9_]+products)['"]""", html)

        if app_id_m and api_key_m:
            app_id = app_id_m.group(1)
            api_key = api_key_m.group(1)
            index_name = index_m.group(1) if index_m else "panafoto_mcprod_cloud_b2c_products"
            algolia_url = f"https://{app_id}-dsn.algolia.net/1/indexes/{index_name}/query"
            headers = {
                "X-Algolia-Application-Id": app_id,
                "X-Algolia-API-Key": api_key,
                "Content-Type": "application/json",
            }
            for b in brands:
                res = client.post(algolia_url, headers=headers, json={"params": f"query={b}&hitsPerPage=250"}, timeout=20)
                for hit in res.json().get("hits", []):
                    price_obj = hit.get("price", {}).get("USD", {})
                    price = price_obj.get("default")
                    if isinstance(price, dict):
                        price = price.get("default")
                    price = extract_price_float(price)
                    if price:
                        items.append({
                            "tienda": "Panafoto",
                            "sku": str(hit.get("sku", "")),
                            "barcode": clean_barcode(hit.get("ean") or hit.get("barcode") or ""),
                            "titulo": str(hit.get("name", "")),
                            "precio": price,
                            "url": str(hit.get("url", "https://www.panafoto.com")),
                        })
            if items:
                return items
    except Exception:
        pass
    return scrape_magento_html_multi_brand(client, "Panafoto", "https://www.panafoto.com", brands)

def scrape_generic_store(client: httpx.Client, store_name: str, base_url: str, brands: list):
    """Detector adaptativo en cascada: Shopify -> VTEX -> WooCommerce -> Magento/HTML."""
    items = scrape_shopify_multi_brand(client, store_name, base_url, brands)
    if items:
        return items
    items = scrape_vtex_multi_brand(client, store_name, base_url, brands)
    if items:
        return items
    items = scrape_woocommerce_multi_brand(client, store_name, base_url, brands)
    if items:
        return items
    return scrape_magento_html_multi_brand(client, store_name, base_url, brands)

def match_products(catalog, scraped_items):
    results = []
    total_matches = 0

    by_store = {s: [i for i in scraped_items if i["tienda"] == s and i["precio"]] for s in STORE_ORDER}

    for sku_row in catalog:
        row_out = {
            "marca": sku_row["marca"],
            "categoria": sku_row["categoria"],
            "codigo": sku_row["codigo"],
            "barcode": sku_row["barcode"],
            "descripcion": sku_row["descripcion"],
            "precio_regular": sku_row["precio_regular"],
            "map_usd": sku_row["map_usd"],
            "disponibilidad": sku_row["disponibilidad"],
            "tiendas": {s: None for s in STORE_ORDER},
        }

        for s in STORE_ORDER:
            store_items = by_store[s]
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
                    "titulo_tienda": matched["titulo"],
                }

        results.append(row_out)

    print(f"Total de coincidencias cruzadas: {total_matches}")
    return results

def main():
    catalog, brands = load_all_catalogs()
    stores_config = [
        ("Multimax", "https://www.multimax.net"),
        ("Novey", "https://www.novey.com.pa"),
        ("Rodelag", "https://rodelag.com"),
        ("Panafoto", "https://www.panafoto.com"),
        ("EVision", "https://www.evisionstore.com"),
        ("Do It Center", "https://www.doitcenter.com.pa"),
        ("Titan", "https://titan.com.pa"),
        ("Spiegel", "https://spiegel.com.pa"),
        ("Selecto", "https://selectopa.com"),
        ("Creditos Mundiales", "https://creditosmundiales.com"),
        ("Unica Panama", "https://unicapanama.com"),
        ("Raenco", "https://www.raenco.com"),
    ]

    all_scraped = []
    with httpx.Client(headers=HEADERS, follow_redirects=True) as client:
        for name, url in stores_config:
            if name == "Panafoto":
                res = scrape_panafoto(client, brands)
            else:
                res = scrape_generic_store(client, name, url, brands)
            print(f"[{name}] Extraidos {len(res)} items.")
            all_scraped.extend(res)

    matched_data = match_products(catalog, all_scraped)

    panama_tz = timezone(timedelta(hours=-5))
    now_str = datetime.now(panama_tz).strftime("%Y-%m-%d %H:%M:%S")

    output = {
        "actualizado": now_str,
        "tiendas_orden": STORE_ORDER,
        "total_skus": len(matched_data),
        "productos": matched_data,
    }

    Path("data/precios_dashboard.json").write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Dashboard actualizado correctamente ({now_str}).")

if __name__ == "__main__":
    main()
