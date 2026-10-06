"""
Facebook Page Crawler - Lead Generation Edition
Search keywords -> collect business pages with contact info -> export Excel
Like the XHS tool: extract phone/email/website/WhatsApp and score leads.
Edit keywords.txt to change search terms.
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS = 50
OUTPUT_DIR = Path(__file__).parent / "导出结果"
HEADLESS = False
# ================================

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{3,4}[-.\s]?\d{3,5}")
WHATSAPP_RE = re.compile(r"(?:whatsapp|wa\.me|whats\s*app)\s*[:]?\s*(\+?\d[\d\s-]{6,})", re.I)
WEBSITE_RE = re.compile(r"https?://(?:www\.)?[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:/\S*)?")


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag manufacturer"]


def extract_contacts(text):
    emails = sorted(set(EMAIL_RE.findall(text or "")))
    phones = sorted(set(PHONE_RE.findall(text or "")))
    whatsapps = sorted(set(WHATSAPP_RE.findall(text or "")))
    websites = sorted(set(WEBSITE_RE.findall(text or "")))
    return emails, phones, whatsapps, websites


def score_lead(has_email, has_phone, has_website, has_whatsapp, about_text):
    score = 0
    reasons = []
    if has_email:
        score += 30
        reasons.append("email")
    if has_phone:
        score += 30
        reasons.append("phone")
    if has_whatsapp:
        score += 25
        reasons.append("whatsapp")
    if has_website:
        score += 15
        reasons.append("website")
    # Keyword match in about = business intent
    biz_words = ["manufacturer", "factory", "supplier", "wholesale", "OEM", "ODM",
                 "leather", "bag", "goods", "custom", "producer", "mill"]
    found = [w for w in biz_words if w.lower() in (about_text or "").lower()]
    score += min(20, len(found) * 4)
    if found:
        reasons.append("biz:" + ",".join(found[:3]))
    return min(100, score), reasons


async def crawl_facebook(page, keyword):
    url = f"https://www.facebook.com/search/pages/?q={keyword}"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(8)

    if "login" in page.url or "checkpoint" in page.url:
        print("  Please login to Facebook in the browser window...", flush=True)
        for _ in range(120):
            await asyncio.sleep(2)
            if "search" in page.url and "login" not in page.url:
                break
        await asyncio.sleep(3)

    print("  Loading results...", flush=True)
    for _ in range(20):
        await page.evaluate("window.scrollBy(0, 1200)")
        await asyncio.sleep(1)

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
            await asyncio.sleep(4)

            # Get all visible text from main area
            data = await page.evaluate("""() => {
                const main = document.querySelector('div[role="main"]') || document.body;
                return {
                    name: document.querySelector('h1')?.innerText?.trim() || '',
                    text: main.innerText || '',
                };
            }""")

            text = data.get("text", "")
            emails, phones, was, websites = extract_contacts(text)
            score, reasons = score_lead(
                emails, phones, was, websites, text
            )

            results.append({
                "keyword": keyword,
                "name": data.get("name", ""),
                "url": page.url,
                "emails": ", ".join(emails),
                "phones": ", ".join(phones),
                "whatsapp": ", ".join(was),
                "websites": ", ".join(websites),
                "about": text[:500],
                "score": score,
                "reasons": ", ".join(reasons),
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

            if i % 5 == 0:
                print(f"    [{i}/{len(links)}] {data.get('name','')[:40]} (score={score})", flush=True)

            await asyncio.sleep(1)
        except Exception as e:
            print(f"    Skip: {e}", flush=True)

    # Sort by score desc
    results.sort(key=lambda x: x["score"], reverse=True)
    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Facebook Leads"

    headers = ["Score", "Page Name", "URL", "Email", "Phone", "WhatsApp",
               "Website", "About snippet", "Reasons", "Keyword", "Collected At"]
    ws.append(headers)

    for r in all_data:
        ws.append([
            r["score"], r["name"], r["url"],
            r["emails"], r["phones"], r["whatsapp"], r["websites"],
            r["about"][:200], r["reasons"], r["keyword"], r["collect_time"]
        ])

    # Style header
    fill = PatternFill("solid", fgColor="4472C4")
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    widths = [8, 35, 50, 30, 20, 20, 35, 50, 25, 20, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    fname = OUTPUT_DIR / f"Facebook_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)
    return fname


async def main():
    keywords = load_keywords()
    print("=" * 50, flush=True)
    print("  Facebook Page Crawler - Lead Edition", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        ctx = await browser.new_context(locale="en-US", viewport={"width": 1280, "height": 800})
        page = await ctx.new_page()

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
