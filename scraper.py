import json
import re
import glob
from datetime import datetime, timezone, timedelta
from pathlib import Path
import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser as HTMLParser

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json,*/*;q=0.8",
    "Accept-Language": "es-PA,es;q=0.9,en;q=0.8",
}

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

    brands = sorted({c["marca"] for c in catalog if c["marca"]} | {"HISENSE", "WHIRLPOOL", "KITCHENAID"})
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

def get_ac_and_priority_queries(catalog):
    queries = []
    for item in catalog:
        if "AIRE" in item["categoria"] or "BUILT" in item["categoria"] or "EXTRAC" in item["categoria"]:
            queries.append(item["codigo"])
    return list(dict.fromkeys(queries))

def parse_html_cards(html: str, store_name: str, base_url: str, fallback_url: str, seen_keys: set):
    """Extractor universal de tarjetas HTML, JSON-LD y scripts embebidos."""
    found = []
    tree = HTMLParser(html)

    # 1. JSON-LD ItemList / Product
    for script in tree.css('script[type="application/ld+json"]'):
        try:
            ld = json.loads(script.text())
            objs = ld if isinstance(ld, list) else [ld]
            for obj in objs:
                if not isinstance(obj, dict):
                    continue
                if obj.get("@type") == "ItemList":
                    for el in obj.get("itemListElement", []):
                        prod = el.get("item", {}) if isinstance(el, dict) else {}
                        offers = prod.get("offers", {})
                        if isinstance(offers, list) and offers:
                            offers = offers[0]
                        price = extract_price_float(offers.get("price") if isinstance(offers, dict) else None)
                        u = str(prod.get("url") or fallback_url)
                        sku = str(prod.get("sku") or "")
                        dedup = sku or u
                        if price and dedup not in seen_keys:
                            seen_keys.add(dedup)
                            found.append({
                                "tienda": store_name,
                                "sku": sku,
                                "barcode": clean_barcode(prod.get("gtin13") or prod.get("gtin") or ""),
                                "titulo": str(prod.get("name") or ""),
                                "precio": price,
                                "url": u,
                            })
        except Exception:
            pass

    # 2. Tarjetas de producto HTML (Magento, Hyva, Botble/Laravel, VTEX SSR, Custom)
    card_selectors = (
        "li.product-item, div.product-item-info, .products-grid .product-item, "
        ".product-item, .product-card, .product-wrapper, .product, .type-product, "
        "[data-product-id], [data-product-sku]"
    )
    for card in tree.css(card_selectors):
        link_node = card.css_first(
            "a.product-item-link, .product-item-name a, .product-title a, "
            "h2 a, h3 a, h4 a, .name a, a[title], a[href]"
        )
        price_node = card.css_first(
            "[data-price-amount], span.price-wrapper[data-price-amount], "
            ".special-price .price, .price-final_price .price, "
            ".product-price, .price ins .amount, .price .amount, span.price, .price"
        )
        if not (link_node and price_node):
            continue

        href = link_node.attributes.get("href", "")
        if href and not href.startswith("http"):
            href = base_url.rstrip("/") + "/" + href.lstrip("/")
        if not href or href.endswith("#") or "javascript:" in href:
            href = fallback_url

        title = link_node.text(strip=True) or link_node.attributes.get("title", "")
        if not title:
            img = card.css_first("img[alt]")
            if img:
                title = img.attributes.get("alt", "")

        sku = card.attributes.get("data-product-sku", "") or ""
        if not sku:
            sku_el = card.css_first("[data-product-sku]")
            if sku_el:
                sku = sku_el.attributes.get("data-product-sku", "")
        if not sku:
            m_sku = re.search(r"data-product-sku=[\"']([^\"']+)[\"']", card.html or "")
            if m_sku:
                sku = m_sku.group(1)

        p_val = price_node.attributes.get("data-price-amount") or price_node.text(strip=True)
        price = extract_price_float(p_val)
        dedup = sku or f"{href}|{title}"
        if price and title and dedup not in seen_keys:
            seen_keys.add(dedup)
            found.append({
                "tienda": store_name,
                "sku": sku,
                "barcode": "",
                "titulo": title,
                "precio": price,
                "url": href,
            })

    return found

def scrape_magento_graphql(client: httpx.Client, store_name: str, base_url: str, search_terms: list, seen_keys: set):
    """Consulta el endpoint nativo /graphql de Magento 2 (funciona cuando el frontend es SPA/Hyva/LiveSearch)."""
    items = []
    gql_url = f"{base_url.rstrip('/')}/graphql"
    query = """
    query SearchProducts($search: String!, $pageSize: Int!) {
      products(search: $search, pageSize: $pageSize) {
        items {
          sku
          name
          url_key
          url_suffix
          price_range {
            minimum_price {
              final_price { value }
            }
          }
        }
      }
    }
    """
    for term in search_terms:
        try:
            r = client.post(
                gql_url,
                json={"query": query, "variables": {"search": term, "pageSize": 100}},
                headers={**HEADERS, "Content-Type": "application/json"},
                timeout=20,
            )
            if r.status_code != 200:
                continue
            data = r.json()
            prods = data.get("data", {}).get("products", {}).get("items", [])
            for p in prods:
                sku = str(p.get("sku") or "")
                if not sku or sku in seen_keys:
                    continue
                price = (
                    p.get("price_range", {})
                    .get("minimum_price", {})
                    .get("final_price", {})
                    .get("value")
                )
                price = extract_price_float(price)
                if price:
                    seen_keys.add(sku)
                    ukey = p.get("url_key", "")
                    usuf = p.get("url_suffix") or ".html"
                    url = f"{base_url.rstrip('/')}/{ukey}{usuf}" if ukey else base_url
                    items.append({
                        "tienda": store_name,
                        "sku": sku,
                        "barcode": "",
                        "titulo": str(p.get("name") or ""),
                        "precio": price,
                        "url": url,
                    })
        except Exception:
            pass
    return items

def scrape_panafoto(client: httpx.Client, brands: list, extra_queries: list):
    items = []
    seen_keys = set()
    base_url = "https://www.panafoto.com"
    brand_urls = [
        "https://www.panafoto.com/categorias/marcas/hisense.html?manufacturer=HISENSE",
        "https://www.panafoto.com/categorias/marcas/whirlpool.html?manufacturer=WHIRLPOOL",
        "https://www.panafoto.com/categorias/marcas/kitchen-aid.html?manufacturer=KitchenAid",
    ]
    for b_url in brand_urls:
        for page in range(1, 6):
            sep = "&" if "?" in b_url else "?"
            url = f"{b_url}{sep}p={page}&product_list_limit=60"
            try:
                r = client.get(url, timeout=22)
                if r.status_code != 200:
                    break
                batch = parse_html_cards(r.text, "Panafoto", base_url, url, seen_keys)
                items.extend(batch)
                if not batch:
                    break
            except Exception:
                break

    # Complementar con GraphQL de Magento 2 (clave para aires acondicionados y catálogos dinámicos)
    items.extend(scrape_magento_graphql(client, "Panafoto", base_url, list(brands) + extra_queries[:25], seen_keys))
    return items

def scrape_novey(client: httpx.Client, brands: list, extra_queries: list):
    items = []
    seen_keys = set()
    base_url = "https://www.novey.com.pa"
    search_urls = [
        "https://www.novey.com.pa/catalogsearch/result/?q=HISENSE",
        "https://www.novey.com.pa/catalogsearch/result/?q=WHIRLPOOL",
        "https://www.novey.com.pa/catalogsearch/result/?q=KITCHEN%20AID",
        "https://www.novey.com.pa/catalogsearch/result/?q=ACROS",
    ]
    for s_url in search_urls:
        for page in range(1, 5):
            url = f"{s_url}&p={page}&product_list_limit=60"
            try:
                r = client.get(url, timeout=22)
                if r.status_code != 200:
                    break
                batch = parse_html_cards(r.text, "Novey", base_url, url, seen_keys)
                items.extend(batch)
                if not batch:
                    break
            except Exception:
                break

    # Complementar con GraphQL de Magento 2
    items.extend(scrape_magento_graphql(client, "Novey", base_url, list(brands) + extra_queries[:20], seen_keys))
    return items

def scrape_doitcenter(client: httpx.Client):
    items = []
    seen_keys = set()
    base_url = "https://www.doitcenter.com.pa"
    brand_urls = [
        "https://www.doitcenter.com.pa/categorias/marcas/hisense",
        "https://www.doitcenter.com.pa/categorias/marcas/whirlpool",
        "https://www.doitcenter.com.pa/categorias/marcas/acros",
    ]
    for b_url in brand_urls:
        for page in range(1, 6):
            url = f"{b_url}?page={page}"
            try:
                r = client.get(url, timeout=22)
                if r.status_code != 200:
                    break
                batch = parse_html_cards(r.text, "Do It Center", base_url, url, seen_keys)
                items.extend(batch)
                if not batch:
                    break
            except Exception:
                break
    return items

def scrape_evision(client: httpx.Client):
    items = []
    seen_keys = set()
    base_url = "https://www.evisionstore.com"
    urls = [
        "https://www.evisionstore.com/products?brand=hisense&num=60",
        "https://www.evisionstore.com/products?brand=kitchenaid&num=60",
        "https://www.evisionstore.com/products?layout=product-full-width&num=60&q=whirlpool&view=list",
        "https://www.evisionstore.com/products?num=60&q=acros",
    ]
    for u in urls:
        for page in range(1, 5):
            page_url = f"{u}&page={page}"
            try:
                r = client.get(page_url, timeout=22)
                if r.status_code != 200:
                    break
                batch = parse_html_cards(r.text, "EVision", base_url, page_url, seen_keys)
                items.extend(batch)
                if not batch:
                    break
            except Exception:
                break
    return items

def scrape_shopify_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list, extra_queries: list):
    items = []
    seen_ids = set()
    brand_tokens = [b.lower() for b in brands]

    endpoints_templates = [f"{base_url}/products.json?limit=250&page={{page}}"]
    for b in brand_tokens:
        endpoints_templates.insert(0, f"{base_url}/collections/{b}/products.json?limit=250&page={{page}}")

    for tmpl in endpoints_templates:
        for page in range(1, 8):
            ep = tmpl.format(page=page)
            try:
                r = client.get(ep, timeout=18)
                if r.status_code != 200:
                    break
                products = r.json().get("products", [])
                if not products:
                    break
                for p in products:
                    vendor_title = f"{p.get('vendor', '')} {p.get('title', '')} {p.get('tags', '')}".lower()
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

    for q in extra_queries[:20]:
        try:
            r = client.get(f"{base_url}/search/suggest.json?q={q}&resources[type]=product", timeout=12)
            if r.status_code == 200:
                prods = r.json().get("resources", {}).get("results", {}).get("products", [])
                for p in prods:
                    pid = p.get("id")
                    if pid in seen_ids:
                        continue
                    price = extract_price_float(p.get("price"))
                    if price:
                        seen_ids.add(pid)
                        items.append({
                            "tienda": store_name,
                            "sku": q,
                            "barcode": "",
                            "titulo": str(p.get("title", "")),
                            "precio": price,
                            "url": f"{base_url}{p.get('url', '')}",
                        })
        except Exception:
            pass
    return items

def scrape_vtex_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list, extra_queries: list):
    items = []
    seen_skus = set()
    for term in list(brands) + extra_queries[:20]:
        max_offset = 150 if term in brands else 50
        for start in range(0, max_offset, 50):
            url = f"{base_url}/api/catalog_system/pub/products/search/{term.lower()}?_from={start}&_to={start+49}"
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

def scrape_woocommerce_multi_brand(client: httpx.Client, store_name: str, base_url: str, brands: list, extra_queries: list):
    items = []
    seen_ids = set()
    for term in list(brands) + extra_queries[:20]:
        url = f"{base_url}/wp-json/wc/store/products?search={term}&per_page=100"
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

def scrape_generic_store(client: httpx.Client, store_name: str, base_url: str, brands: list, extra_queries: list):
    items = scrape_shopify_multi_brand(client, store_name, base_url, brands, extra_queries)
    if items:
        return items
    items = scrape_vtex_multi_brand(client, store_name, base_url, brands, extra_queries)
    if items:
        return items
    items = scrape_woocommerce_multi_brand(client, store_name, base_url, brands, extra_queries)
    if items:
        return items
    seen = set()
    out = []
    for term in list(brands) + extra_queries[:15]:
        for p in range(1, 5):
            u = f"{base_url}/catalogsearch/result/index/?p={p}&q={term}"
            try:
                r = client.get(u, timeout=18)
                if r.status_code != 200:
                    break
                batch = parse_html_cards(r.text, store_name, base_url, u, seen)
                out.extend(batch)
                if not batch:
                    break
            except Exception:
                break
    return out

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
                    haystack = clean_code(item["sku"] + " " + item["titulo"] + " " + item["url"])
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
    extra_queries = get_ac_and_priority_queries(catalog)
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
                res = scrape_panafoto(client, brands, extra_queries)
            elif name == "Novey":
                res = scrape_novey(client, brands, extra_queries)
            elif name == "Do It Center":
                res = scrape_doitcenter(client)
            elif name == "EVision":
                res = scrape_evision(client)
            else:
                res = scrape_generic_store(client, name, url, brands, extra_queries)
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
