import json
import re
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from pathlib import Path

import httpx
import pandas as pd
from selectolax.lexbor import LexborHTMLParser as HTMLParser

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-PA,es;q=0.9,en;q=0.8",
}
JSON_HEADERS = {**HEADERS, "Accept": "application/json"}

STORE_ORDER = [
    "Multimax", "Novey", "Rodelag", "Panafoto", "EVision", "Do It Center",
    "Titan", "Spiegel", "Selecto", "Creditos Mundiales", "Unica Panama", "Raenco",
]

STORES = {
    "Multimax": "https://www.multimax.net",
    "Novey": "https://www.novey.com.pa",
    "Rodelag": "https://rodelag.com",
    "Panafoto": "https://www.panafoto.com",
    "EVision": "https://www.evisionstore.com",
    "Do It Center": "https://www.doitcenter.com.pa",
    "Titan": "https://titan.com.pa",
    "Spiegel": "https://spiegel.com.pa",
    "Selecto": "https://selectopa.com",
    "Creditos Mundiales": "https://creditosmundiales.com",
    "Unica Panama": "https://unicapanama.com",
    "Raenco": "https://www.raenco.com",
}

CODE_ALIASES = {"RT16N6CDX": "RT3N461NACA"}
MAX_WORKERS = 12
DEBUG_LOG = Path("data/scraper_log.txt")

# ----------------------------------------------------------------------------
# LOGGING
# ----------------------------------------------------------------------------
_log_lines = []

def log(store, msg):
    line = f"[{datetime.now().strftime('%H:%M:%S')}] [{store}] {msg}"
    print(line, flush=True)
    _log_lines.append(line)

def save_log():
    DEBUG_LOG.parent.mkdir(parents=True, exist_ok=True)
    DEBUG_LOG.write_text("\n".join(_log_lines), encoding="utf-8")

# ----------------------------------------------------------------------------
# UTILIDADES
# ----------------------------------------------------------------------------
def clean_code(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).upper())

def clean_barcode(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^0-9]", "", s)

def extract_price_float(raw):
    if raw is None:
        return None
    if isinstance(raw, dict):
        for k in ("USD", "value", "final_price", "price", "amount"):
            if k in raw:
                return extract_price_float(raw[k])
        return None
    if isinstance(raw, (int, float)):
        return float(raw) if raw > 0 else None
    cleaned = re.sub(r"[^0-9.]", "", str(raw).replace(",", ""))
    try:
        v = float(cleaned) if cleaned else 0
        return v if v > 0 else None
    except ValueError:
        return None

def fetch(client, store, url, as_json=False, **kw):
    """GET con log obligatorio. Devuelve la respuesta o None."""
    try:
        r = client.get(url, headers=JSON_HEADERS if as_json else HEADERS, **kw)
        log(store, f"GET {url[:110]} -> {r.status_code} | {r.headers.get('content-type','')[:40]} | {len(r.text)}b")
        return r if r.status_code == 200 else None
    except Exception as e:
        log(store, f"ERROR GET {url[:110]}: {type(e).__name__}: {e}")
        return None

def fetch_json(client, store, url, **kw):
    r = fetch(client, store, url, as_json=True, **kw)
    if r is None:
        return None
    try:
        return r.json()
    except Exception as e:
        log(store, f"JSON inválido en {url[:80]}: {e}")
        return None

# ----------------------------------------------------------------------------
# CATÁLOGO
# ----------------------------------------------------------------------------
def load_all_catalogs():
    files = sorted(glob.glob("data/*.xlsx")) or sorted(glob.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError("No hay Excel de catálogo en data/ ni en la raíz.")

    catalog, seen = [], set()
    for fp in files:
        xls = pd.ExcelFile(fp)
        sheet = "Lista de precios" if "Lista de precios" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(fp, sheet_name=sheet, header=5)
        if "Código" not in df.columns:
            continue
        df = df.dropna(subset=["Código"])
        for _, row in df.iterrows():
            code_raw = str(row["Código"]).strip()
            code_norm = clean_code(code_raw)
            if not code_norm or code_norm in seen:
                continue
            seen.add(code_norm)
            valid = {code_norm}
            for alias, target in CODE_ALIASES.items():
                if target == code_raw:
                    valid.add(clean_code(alias))
            catalog.append({
                "marca": str(row.get("Marca", "")).strip().upper(),
                "categoria": str(row.get("Categoría", "")).strip().upper(),
                "codigo": code_raw,
                "codes_norm": list(valid),
                "barcode": clean_barcode(row.get("BarCode", "")),
                "descripcion": str(row.get("Descripción", "")).strip(),
                "precio_regular": extract_price_float(row.get("Precio Regular (USD)")),
                "map_usd": extract_price_float(row.get("MAP (USD)")),
                "disponibilidad": str(row.get("Disponibilidad", "")).strip(),
            })

    brands = sorted({c["marca"] for c in catalog if c["marca"]})
    queries = list(dict.fromkeys(
        c["codigo"] for c in catalog
        if any(k in c["categoria"] for k in ("AIRE", "BUILT", "EXTRAC"))
    ))
    print(f"Catálogo: {len(catalog)} SKUs | Marcas: {brands} | Queries prioritarias: {len(queries)}")
    return catalog, brands, queries

# ----------------------------------------------------------------------------
# DETECCIÓN DE PLATAFORMA
# ----------------------------------------------------------------------------
PLATFORM_SIGNATURES = [
    ("shopify", ["cdn.shopify.com", "Shopify.theme"]),
    ("vtex", ["vteximg.com.br", "vtex.com", "/_v/", "vtexassets"]),
    ("magento", ["Magento_", "mage/cookies", "/static/frontend/", "requirejs/require"]),
    ("woocommerce", ["wp-content/plugins/woocommerce", "woocommerce"]),
    ("algolia", ["algolia", "instantsearch"]),
]

def detect_platform(client, store, base_url):
    r = fetch(client, store, base_url)
    if r is None:
        return "unknown", ""
    html = r.text
    scores = {}
    for name, sigs in PLATFORM_SIGNATURES:
        hits = sum(1 for s in sigs if s.lower() in html.lower())
        if hits:
            scores[name] = hits
    platform = max(scores, key=scores.get) if scores else "unknown"
    log(store, f"Plataforma detectada: {platform} {scores}")
    return platform, html

# ----------------------------------------------------------------------------
# PARSEO JSON-LD (Product / ItemList / @graph)
# ----------------------------------------------------------------------------
def _jsonld_products(obj, out):
    if isinstance(obj, list):
        for o in obj:
            _jsonld_products(o, out)
        return
    if not isinstance(obj, dict):
        return
    if "@graph" in obj:
        _jsonld_products(obj["@graph"], out)
    t = obj.get("@type")
    if t == "Product" or (isinstance(t, list) and "Product" in t):
        out.append(obj)
    if t == "ItemList":
        for el in obj.get("itemListElement", []):
            if isinstance(el, dict):
                _jsonld_products(el.get("item", el), out)

def parse_jsonld(html, store, fallback_url):
    items = []
    tree = HTMLParser(html)
    prods = []
    for script in tree.css('script[type="application/ld+json"]'):
        try:
            _jsonld_products(json.loads(script.text()), prods)
        except Exception:
            continue
    for p in prods:
        offers = p.get("offers", {})
        if isinstance(offers, list) and offers:
            offers = offers[0]
        price = extract_price_float(offers.get("price") if isinstance(offers, dict) else None)
        if not price:
            continue
        items.append({
            "tienda": store,
            "sku": str(p.get("sku") or ""),
            "barcode": clean_barcode(p.get("gtin13") or p.get("gtin") or p.get("gtin12") or ""),
            "titulo": str(p.get("name") or ""),
            "precio": price,
            "url": str(p.get("url") or fallback_url),
        })
    return items

def parse_html_cards(html, store, base_url, fallback_url):
    """Respaldo: tarjetas de producto genéricas en HTML."""
    items = []
    tree = HTMLParser(html)
    card_sel = (".product-item, .product-card, .product-wrapper, .product, .type-product, "
                "[data-product-id], [data-product-sku], .ps-product, li.product")
    for card in tree.css(card_sel):
        link = card.css_first("a[href]")
        price_node = card.css_first("[data-price-amount], .price, span.price, .product-price")
        if not (link and price_node):
            continue
        href = link.attributes.get("href", "")
        if href and not href.startswith("http"):
            href = base_url.rstrip("/") + "/" + href.lstrip("/")
        title = link.text(strip=True) or link.attributes.get("title", "")
        price = extract_price_float(price_node.attributes.get("data-price-amount") or price_node.text(strip=True))
        if price and title:
            items.append({
                "tienda": store,
                "sku": card.attributes.get("data-product-sku", "") or "",
                "barcode": "",
                "titulo": title,
                "precio": price,
                "url": href or fallback_url,
            })
    return items

# ----------------------------------------------------------------------------
# ESTRATEGIAS
# ----------------------------------------------------------------------------
def strat_magento_graphql(client, store, base, terms):
    query = """
    query SearchProducts($search: String!, $pageSize: Int!) {
      products(search: $search, pageSize: $pageSize) {
        items { sku name url_key url_suffix
          price_range { minimum_price { final_price { value } } } }
      }
    }"""
    items = []
    for term in terms:
        try:
            r = client.post(f"{base}/graphql", json={"query": query,
                            "variables": {"search": term, "pageSize": 100}},
                            headers={**JSON_HEADERS, "Content-Type": "application/json"}, timeout=20)
            if r.status_code != 200:
                continue
            for p in r.json().get("data", {}).get("products", {}).get("items", []):
                price = extract_price_float(p.get("price_range", {}).get("minimum_price", {})
                                            .get("final_price", {}).get("value"))
                if price:
                    uk = p.get("url_key", "")
                    items.append({"tienda": store, "sku": str(p.get("sku", "")), "barcode": "",
                                  "titulo": str(p.get("name", "")), "precio": price,
                                  "url": f"{base}/{uk}{p.get('url_suffix') or '.html'}" if uk else base})
        except Exception as e:
            log(store, f"magento_graphql term={term}: {type(e).__name__}: {e}")
    return items

def strat_vtex(client, store, base, terms):
    items, seen = [], set()
    for term in terms:
        for start in range(0, 150, 50):
            url = f"{base}/api/catalog_system/pub/products/search?ft={term}&_from={start}&_to={start+49}"
            data = fetch_json(client, store, url)
            if not data:
                break
            for p in data:
                for sku in p.get("items", []):
                    iid = sku.get("itemId")
                    if iid in seen:
                        continue
                    seen.add(iid)
                    sellers = sku.get("sellers", [])
                    price = extract_price_float(sellers[0].get("commertialOffer", {}).get("Price")) if sellers else None
                    if not price:
                        continue
                    refs = " ".join(str(x.get("Value", "")) for x in sku.get("referenceId", []))
                    items.append({"tienda": store, "sku": refs.strip(),
                                  "barcode": clean_barcode(sku.get("ean", "")),
                                  "titulo": f"{p.get('productName','')} {sku.get('name','')}".strip(),
                                  "precio": price,
                                  "url": f"{base}/{p.get('linkText','')}/p"})
            if len(data) < 50:
                break
    return items

def strat_shopify(client, store, base, brand_tokens, terms):
    items, seen = [], set()
    for page in range(1, 8):
        data = fetch_json(client, store, f"{base}/products.json?limit=250&page={page}")
        if not data or not data.get("products"):
            break
        for p in data["products"]:
            blob = f"{p.get('vendor','')} {p.get('title','')} {p.get('tags','')}".lower()
            if not any(b in blob for b in brand_tokens):
                continue
            for v in p.get("variants", []):
                if v.get("id") in seen:
                    continue
                seen.add(v.get("id"))
                price = extract_price_float(v.get("price"))
                if price:
                    items.append({"tienda": store, "sku": str(v.get("sku") or ""),
                                  "barcode": clean_barcode(v.get("barcode") or ""),
                                  "titulo": f"{p.get('title','')} {v.get('title','')}".strip(),
                                  "precio": price, "url": f"{base}/products/{p.get('handle','')}"})
    return items

def strat_woocommerce(client, store, base, terms):
    items, seen = [], set()
    for term in terms:
        data = fetch_json(client, store, f"{base}/wp-json/wc/store/products?search={term}&per_page=100")
        if not isinstance(data, list):
            continue
        for p in data:
            if p.get("id") in seen:
                continue
            seen.add(p.get("id"))
            pr = p.get("prices", {})
            minor = int(pr.get("currency_minor_unit", 2))
            raw = pr.get("price")
            price = float(raw) / (10 ** minor) if raw else None
            if price:
                items.append({"tienda": store, "sku": str(p.get("sku", "")), "barcode": "",
                              "titulo": str(p.get("name", "")), "precio": price,
                              "url": str(p.get("permalink", base))})
    return items

def strat_algolia(client, store, base, html, terms):
    """Detecta appId/apiKey/indexName embebidos en el HTML y consulta Algolia directo."""
    app = re.search(r'applicationId["\']?\s*[:=]\s*["\']([A-Za-z0-9]{10})["\']', html)
    key = re.search(r'apiKey["\']?\s*[:=]\s*["\']([a-f0-9]{32})["\']', html)
    idx = re.search(r'indexName["\']?\s*[:=]\s*["\']([^"\']+)["\']', html)
    if not (app and key):
        log(store, "Algolia: no se encontraron appId/apiKey en el HTML")
        return []
    index = idx.group(1) if idx else None
    if not index:
        log(store, "Algolia: appId/apiKey encontrados, pero falta indexName; revisa el JS de la página")
        return []
    url = f"https://{app.group(1).lower()}-dsn.algolia.net/1/indexes/*/queries"
    headers = {"X-Algolia-Application-Id": app.group(1), "X-Algolia-API-Key": key.group(1),
               "Content-Type": "application/json"}
    items, seen = [], set()
    for term in terms:
        body = {"requests": [{"indexName": index, "params": f"query={term}&hitsPerPage=1000"}]}
        try:
            r = client.post(url, json=body, headers=headers, timeout=20)
            log(store, f"Algolia '{term}' -> {r.status_code}")
            if r.status_code != 200:
                continue
            for h in r.json()["results"][0].get("hits", []):
                sku = str(h.get("sku") or h.get("objectID") or "")
                if sku in seen:
                    continue
                price = extract_price_float(h.get("price") or h.get("price_range") or h.get("final_price"))
                if price:
                    seen.add(sku)
                    items.append({"tienda": store, "sku": sku,
                                  "barcode": clean_barcode(h.get("gtin") or h.get("ean") or ""),
                                  "titulo": str(h.get("name") or h.get("title") or ""),
                                  "precio": price,
                                  "url": str(h.get("url") or h.get("url_key") or base)})
        except Exception as e:
            log(store, f"Algolia error: {type(e).__name__}: {e}")
    return items

def strat_sitemap_jsonld(client, store, base, brand_tokens, max_pages=600):
    """Universal: recorre sitemap, filtra URLs de producto por marca y lee JSON-LD de cada ficha."""
    urls, to_visit, visited = [], [f"{base}/sitemap.xml"], set()
    while to_visit and len(visited) < 40:
        sm = to_visit.pop()
        if sm in visited:
            continue
        visited.add(sm)
        r = fetch(client, store, sm)
        if not r:
            continue
        locs = re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r.text)
        for loc in locs:
            if loc.endswith(".xml") or loc.endswith(".xml.gz"):
                to_visit.append(loc)
            else:
                urls.append(loc)
    log(store, f"Sitemap: {len(urls)} URLs totales")
    cand = [u for u in urls if any(b in u.lower() for b in brand_tokens)]
    cand = cand[:max_pages]
    log(store, f"Sitemap: {len(cand)} URLs con marca en la ruta")

    def read(u):
        try:
            rr = client.get(u, headers=HEADERS, timeout=20)
            return parse_jsonld(rr.text, store, u) if rr.status_code == 200 else []
        except Exception:
            return []

    items = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        for res in ex.map(read, cand):
            items.extend(res)
    return items

def strat_html_listing(client, store, base, terms):
    items = []
    for term in terms:
        for url in (f"{base}/search?q={term}", f"{base}/catalogsearch/result/?q={term}",
                    f"{base}/buscar?q={term}"):
            r = fetch(client, store, url)
            if r:
                found = parse_jsonld(r.text, store, url) or parse_html_cards(r.text, store, base, url)
                items.extend(found)
                if found:
                    break
    return items

def strat_panafoto_brands(client, store, base, brands):
    """Estrategia específica original de Panafoto (funcionaba: 247 items)."""
    items, seen = [], set()
    brand_urls = [
        f"{base}/categorias/marcas/hisense.html?manufacturer=HISENSE",
        f"{base}/categorias/marcas/whirlpool.html?manufacturer=WHIRLPOOL",
        f"{base}/categorias/marcas/kitchen-aid.html?manufacturer=KitchenAid",
    ]
    for bu in brand_urls:
        for page in range(1, 6):
            r = fetch(client, store, f"{bu}&p={page}&product_list_limit=60")
            if not r:
                break
            batch = parse_html_cards(r.text, store, base, bu)
            new = [b for b in batch if (b["sku"] or b["url"]) not in seen]
            for b in new:
                seen.add(b["sku"] or b["url"])
            items.extend(new)
            if not batch:
                break
    return items

# ----------------------------------------------------------------------------
# ORQUESTADOR POR TIENDA
# ----------------------------------------------------------------------------
def scrape_store(name, base, brands, queries):
    client = httpx.Client(follow_redirects=True, timeout=25, headers=HEADERS)
    try:
        brand_tokens = [b.lower() for b in brands]
        terms = list(brands) + queries[:15]
        platform, home_html = detect_platform(client, name, base)

        # Orden de estrategias según la plataforma detectada; luego respaldo universal.
        plan = {
            "shopify": [lambda: strat_shopify(client, name, base, brand_tokens, terms)],
            "vtex": [lambda: strat_vtex(client, name, base, terms)],
            "magento": [lambda: strat_magento_graphql(client, name, base, terms)],
            "woocommerce": [lambda: strat_woocommerce(client, name, base, terms)],
            "algolia": [lambda: strat_algolia(client, name, base, home_html, terms)],
        }.get(platform, [])

        fallbacks = [
            lambda: strat_algolia(client, name, base, home_html, terms),
            lambda: strat_shopify(client, name, base, brand_tokens, terms),
            lambda: strat_vtex(client, name, base, terms),
            lambda: strat_magento_graphql(client, name, base, terms),
            lambda: strat_woocommerce(client, name, base, terms),
            lambda: strat_sitemap_jsonld(client, name, base, brand_tokens),
            lambda: strat_html_listing(client, name, base, terms[:10]),
        ]
        if name == "Panafoto":
            plan = [lambda: strat_panafoto_brands(client, name, base, brands)] + plan

        for strat in plan + fallbacks:
            try:
                items = strat()
            except Exception as e:
                log(name, f"Estrategia falló: {type(e).__name__}: {e}")
                continue
            if items:
                log(name, f"OK: {len(items)} items")
                return items
            log(name, "Estrategia sin resultados, probando siguiente")

        log(name, "SIN RESULTADOS en todas las estrategias")
        return []
    finally:
        client.close()

# ----------------------------------------------------------------------------
# CRUCE CON CATÁLOGO
# ----------------------------------------------------------------------------
def match_products(catalog, scraped):
    by_store = {s: [i for i in scraped if i["tienda"] == s] for s in STORE_ORDER}
    results, total = [], 0
    for row in catalog:
        out = {k: row[k] for k in ("marca", "categoria", "codigo", "barcode", "descripcion",
                                   "precio_regular", "map_usd", "disponibilidad")}
        out["tiendas"] = {s: None for s in STORE_ORDER}
        for s in STORE_ORDER:
            matched = None
            if row["barcode"]:
                matched = next((i for i in by_store[s] if i["barcode"] == row["barcode"]), None)
            if not matched:
                for i in by_store[s]:
                    hay = clean_code(i["sku"] + " " + i["titulo"] + " " + i["url"])
                    if any(c in hay for c in row["codes_norm"] if len(c) >= 4):
                        matched = i
                        break
            if matched:
                total += 1
                out["tiendas"][s] = {"precio": matched["precio"], "url": matched["url"],
                                     "titulo_tienda": matched["titulo"]}
        results.append(out)
    print(f"Total de coincidencias cruzadas: {total}")
    return results

# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main():
    catalog, brands, queries = load_all_catalogs()
    all_scraped, summary = [], {}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(scrape_store, n, u, brands, queries): n for n, u in STORES.items()}
        for f in as_completed(futures):
            name = futures[f]
            try:
                res = f.result()
            except Exception as e:
                log(name, f"FALLO TOTAL: {type(e).__name__}: {e}")
                res = []
            summary[name] = len(res)
            all_scraped.extend(res)

    print("\n=== RESUMEN ===")
    for s in STORE_ORDER:
        print(f"{s:<22} {summary.get(s, 0):>5} items")

    matched = match_products(catalog, all_scraped)
    now = datetime.now(timezone(timedelta(hours=-5))).strftime("%Y-%m-%d %H:%M:%S")
    output = {"actualizado": now, "tiendas_orden": STORE_ORDER,
              "total_skus": len(matched), "productos": matched}
    Path("data/precios_dashboard.json").write_text(
        json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    save_log()
    print(f"\nDashboard actualizado ({now}). Log completo en {DEBUG_LOG}")

if __name__ == "__main__":
    main()
