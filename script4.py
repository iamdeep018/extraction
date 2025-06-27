import os
import asyncio
import openpyxl
import csv
import aiohttp
from playwright.async_api import async_playwright
from dotenv import load_dotenv

# Paths
excel_path = "input/D2D.xlsx"
output_csv_path = "output/KB_Hyperlinks.csv"
screenshot_dir = "screenshots"
esm_links_csv_path = "output/ESM_Viewer_Links.csv"

# Check if a URL is working
async def check_link(session, url):
    try:
        if not url.startswith("http"):
            return "Broken (Invalid URL)"
        async with session.head(url, timeout=10, allow_redirects=True) as response:
            if 200 <= response.status < 400:
                return "Working"
            return f"Broken ({response.status})"
    except Exception:
        # Retry with GET
        try:
            async with session.get(url, timeout=10, allow_redirects=True) as response:
                if 200 <= response.status < 400:
                    return "Working"
                return f"Broken ({response.status})"
        except Exception as e:
            return f"Broken ({str(e).split()[0]})"

# Extract cookie
async def extract_jsession_id(context):
    cookies = await context.cookies()
    for cookie in cookies:
        if cookie["name"] == "JSESSIONID":
            return cookie["value"]
    return None

# Read KBs
def read_kb_numbers_from_excel(path):
    kb_numbers = []
    try:
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            kb_number = row[0]
            if kb_number:
                kb_numbers.append(str(kb_number).strip())
    except Exception as e:
        print(f" Failed to read Excel: {e}")
    return kb_numbers

#new scrapper
async def scrape_kb_articles(kb_numbers):
    from dotenv import load_dotenv
    load_dotenv()

    os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)
    os.makedirs(screenshot_dir, exist_ok=True)

    # Write header with new columns
    with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["KB Number", "Link Text", "URL", "viewer_link", "Inside_viewer"])

    user_data_dir = r"C:\\Users\\DSaha6\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2"
    async with aiohttp.ClientSession() as session:
        async with async_playwright() as p:
            context = await p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                args=['--start-maximized']
            )
            page = await context.new_page()

            for kb_number in kb_numbers:
                full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
                print(f"\n🔎 Fetching KB: {kb_number} → {full_url}")

                try:
                    response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
                    status = response.status if response else None
                    current_url = page.url
                except Exception as e:
                    print(f"❌ Error loading KB {kb_number}: {e}")
                    await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
                    continue

                if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
                    print(f"🔒 Redirected to login for KB {kb_number}")
                    await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
                    continue

                if status == 200:
                    print(f"✅ Loaded KB {kb_number}")
                    try:
                        await page.wait_for_selector('article.kb-article-content', timeout=50000)
                        article = await page.query_selector('article.kb-article-content')
                        links = await article.query_selector_all('a')

                        with open(output_csv_path, mode='a', newline='', encoding='utf-8') as file:
                            writer = csv.writer(file)
                            for link in links:
                                href = await link.get_attribute("href") or ""
                                text = await link.inner_text() or ""
                                href = href.strip()
                                text = text.strip()

                                if not href or href.startswith("mailto:") or href.startswith("#"):
                                    continue

                                is_viewer_link = "Yes" if "$viewer.do" in href else "No"
                                writer.writerow([kb_number, text, href, is_viewer_link, "No"])
                                print(f"- {text}: {href}")

                                # 🎯 NEW: If it's a viewer link, go and extract internal viewer links
                                if is_viewer_link == "Yes":
                                    print(f"🔍 Exploring viewer: {href}")
                                    try:
                                        await page.goto(href, wait_until="domcontentloaded", timeout=60000)
                                        await page.wait_for_selector('div#viewer.pdfViewer', timeout=10000)
                                        viewer = await page.query_selector('div#viewer.pdfViewer')
                                        if not viewer:
                                            print("❌ Viewer div not found.")
                                            await page.screenshot(path=f"{screenshot_dir}/{kb_number}_viewer_not_found.png")
                                            continue

                                        viewer_links = await viewer.query_selector_all('a')
                                        for vlink in viewer_links:
                                            vhref = (await vlink.get_attribute("href") or "").strip()
                                            vtext = (await vlink.inner_text() or "").strip()
                                            if vhref:
                                                writer.writerow([kb_number, vtext, vhref, "No", "Yes"])
                                                print(f"🔗 Viewer: {vtext} → {vhref}")

                                    except Exception as e:
                                        print(f"❌ Error scraping viewer link: {e}")
                                        await page.screenshot(path=f"{screenshot_dir}/{kb_number}_viewer_error.png")

                    except Exception as e:
                        print(f"❌ Error extracting links in KB {kb_number}: {e}")
                        await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
                else:
                    print(f"⚠️ HTTP {status} error for KB {kb_number}")

            await context.close()

    print(f"\n✅ Completed scraping KB articles and inline viewer links.")


async def verify_links_in_csv(csv_path):
    print("\n Verifying hyperlinks in the CSV file...")

    esm_links = []  # Collect rows with esm viewer links

    # Read the existing CSV
    rows = []
    with open(csv_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        headers = next(reader)
        for row in reader:
            rows.append(row)

    # Function to map HTTP status to label
    def interpret_status(code):
        if 200 <= code < 300:
            return "OK"
        elif 300 <= code < 400:
            return "OK - Redirected"
        elif code in (401, 403, 405, 407, 423):
            return "OK - Access Required"
        else:
            return f"Broken ({code})"

    # Verify links
    async with aiohttp.ClientSession() as session:
        updated_rows = []
        for row in rows:
            url = row[2].strip()
            kb_number = row[0]

            # Skip invalid or irrelevant URLs
            if not url or url.startswith("mailto:") or url.startswith("#"):
                print(f"Skipped: {url}")
                continue

            # Collect ESM viewer links
            if url.startswith("https://esm.slb.com/$viewer.do"):
                esm_links.append([kb_number, url])

            try:
                # Try HEAD request
                async with session.head(url, timeout=10, allow_redirects=True) as response:
                    status = interpret_status(response.status)
            except Exception:
                # Retry with GET
                try:
                    async with session.get(url, timeout=10, allow_redirects=True) as response:
                        status = interpret_status(response.status)
                except Exception as e:
                    status = f"Broken ({str(e).split()[0]})"

            updated_rows.append(row + [status])
            print(f" {url} → {status}")

    # Write back to the original CSV with status
    new_headers = headers + ["Status"]
    with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(new_headers)
        writer.writerows(updated_rows)

    # Write ESM viewer links to a separate CSV
    if esm_links:
        with open(esm_links_csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["KB Number", "Base URL"])
            writer.writerows(esm_links)
        print(f" Extracted {len(esm_links)} ESM viewer links → {esm_links_csv_path}")


    print(f"\n Hyperlink verification complete. Updated: {csv_path}")

# Main
async def main():
    async with async_playwright() as p:
     
        kb_numbers = read_kb_numbers_from_excel(excel_path)
        if not kb_numbers:
            print("No KB numbers found in Excel.")
            return

        await scrape_kb_articles(kb_numbers)
        await verify_links_in_csv(output_csv_path)
        # await extract_links_from_esm_viewer_pages()
        # await verify_esm_viewer_extracted_links()

if __name__ == "__main__":
    asyncio.run(main())