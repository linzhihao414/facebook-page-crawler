"""
Facebook Page Crawler
Search keywords -> collect public page name, URL, likes, contact info -> export Excel
Edit keywords.txt to change search terms.
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from playwright.async_api import async_playwright

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS = 100
OUTPUT_DIR = Path(__file__).parent / "导出结果"
HEADLESS = False
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag manufacturer"]


async def crawl_facebook(page, keyword):
    """Search Facebook pages by keyword"""
    url = f"https://www.facebook.com/search/pages/?q={keyword}"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(8)

    # Check if login wall
    if "login" in page.url or "checkpoint" in page.url:
        print("  Please login to Facebook in the browser window...", flush=True)
        print("  Waiting for you to complete login...", flush=True)
        # Wait for user to login
        for _ in range(120):
            await asyncio.sleep(2)
            if "search" in page.url and "login" not in page.url:
                break
        await asyncio.sleep(3)

    # Scroll to load results
    print(f"  Loading results...", flush=True)
    for _ in range(15):
        await page.evaluate("window.scrollBy(0, 1000)")
        await asyncio.sleep(1)

    # Get page links
    links = await page.evaluate("""
        () => {
            const as = document.querySelectorAll('a[href*="/pages/"]');
            return Array.from(as).map(a => a.href.split('?')[0]).filter(h => h.includes('/pages/'));
        }
    """)
    links = list(dict.fromkeys(links))[:MAX_RESULTS]
    print(f"  Found {len(links)} pages", flush=True)

    results = []
    for i, link in enumerate(links, 1):
        try:
            await page.goto(link, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(3)

            data = await page.evaluate("""() => {
                const get = (sel) => {
                    const el = document.querySelector(sel);
                    return el ? el.innerText.trim() : '';
                };
                return {
                    name: document.querySelector('h1')?.innerText?.trim() || '',
                    about: get('div[role="main"] div[class*="r"]'),
                };
            }""")

            results.append({
                "keyword": keyword,
                "name": data.get("name", ""),
                "url": page.url,
                "about": data.get("about", "")[:300],
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

            if i % 10 == 0:
                print(f"    [{i}/{len(links)}] {data.get('name','')[:40]}", flush=True)

            await asyncio.sleep(1)
        except Exception as e:
            print(f"    Skip: {e}", flush=True)

    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Facebook Pages"

    headers = ["Keyword", "Page Name", "URL", "About", "Collected At"]
    ws.append(headers)
    for r in all_data:
        ws.append([r["keyword"], r["name"], r["url"], r["about"], r["collect_time"]])

    for i, w in enumerate([15, 35, 50, 50, 18], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    fname = OUTPUT_DIR / f"Facebook_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)


async def main():
    keywords = load_keywords()
    print("=" * 50, flush=True)
    print("  Facebook Page Crawler", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        ctx = await browser.new_context(locale="en-US", viewport={"width": 1280, "height": 800})
        page = await ctx.new_page()

        # First visit facebook.com for login
        print("\nOpening Facebook... Please login if needed.", flush=True)
        await page.goto("https://www.facebook.com", wait_until="domcontentloaded")
        await asyncio.sleep(5)

        total = 0
        for kw in keywords:
            print(f"\n=== {kw} ===", flush=True)
            try:
                results = await crawl_facebook(page, kw)
                save_to_excel(results, kw)
                total += len(results)
            except Exception as e:
                print(f"  Error: {e}", flush=True)

        print(f"\nDone! Total {total} pages.", flush=True)
        print(f"Output: {OUTPUT_DIR.resolve()}", flush=True)
        await browser.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
