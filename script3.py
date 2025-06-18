 
import os
import asyncio
import openpyxl
import csv
import aiohttp
from playwright.async_api import async_playwright
from dotenv import load_dotenv
 
# Paths
excel_path = "input/ESM KB.xlsx"
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
    USERNAME2 = os.getenv("USERNAME2")
    PASSWORD2 = os.getenv("PASSWORD2")
 
    os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)
    os.makedirs(screenshot_dir, exist_ok=True)
 
    # Write header to output CSV
    with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["KB Number", "Link Text", "URL"])
 
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
 
        # Login manually
        print(" Opening ESM Home Page")
        await page.goto("https://slbsandbox.service-now.com", wait_until="domcontentloaded")
 
        print(" Waiting for Microsoft login page...")
        await page.wait_for_url("https://login.microsoftonline.com/**", timeout=60000)
 
        print(" Entering username...")
        await page.fill('input[name="loginfmt"]', USERNAME2)
        await page.click('#idSIButton9')
 
        print(" Entering password...")
        await page.fill('input[name="passwd"]', PASSWORD2)
        await page.click('#idSIButton9')
 
        try:
            await page.wait_for_selector('input[type="submit"][value="Yes"]', timeout=10000)
            await page.click('input[type="submit"][value="Yes"]')
        except:
            print(" 'Stay signed in?' prompt not shown.")
 
        # Now authenticated, proceed to open KB articles
        for kb_number in kb_numbers:
            full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
            print(f"\n Fetching KB: {kb_number} → {full_url}")
 
            try:
                response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
                status = response.status if response else None
                current_url = page.url
            except Exception as e:
                print(f" Error loading KB {kb_number}: {e}")
                await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
                continue
 
            if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
                print(f" Redirected to login for KB {kb_number}")
                await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
                continue
 
            if status == 200:
                print(f" Successfully authenticated and loaded KB {kb_number}")
                try:
                    await page.wait_for_selector('article.kb-article-content', timeout=50000)
                    article = await page.query_selector('article.kb-article-content')
                    links = await article.query_selector_all('a')
 
                    with open(output_csv_path, mode='a', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        for link in links:
                            href = await link.get_attribute("href") or ""
                            text = await link.inner_text() or ""
                            if href.strip() and not href.startswith("mailto:") and not href.startswith("#"):
                                writer.writerow([kb_number, text.strip(), href.strip()])
                                print(f"- {text.strip()}: {href.strip()}")
 
                except Exception as e:
                    print(f" Error extracting links from KB {kb_number}: {e}")
                    await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
            else:
                print(f" HTTP {status} error for KB {kb_number}")
 
        await browser.close()
        print(f"\n Completed scraping KB articles.")
 
 
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
 
async def extract_links_from_esm_viewer_pages():
    esm_links_csv_path = "output/ESM_Viewer_Links.csv"
    extracted_links_csv_path = "output/ESM_Viewer_Extracted_Links.csv"
 
    if not os.path.exists(esm_links_csv_path):
        print(f" ESM links file not found: {esm_links_csv_path}")
        return
 
    # Load credentials
    load_dotenv()
    USERNAME2 = os.getenv("USERNAME2")
    PASSWORD2 = os.getenv("PASSWORD2")
 
    if not USERNAME2 or not PASSWORD2:
        print(" Missing USERNAME2 or PASSWORD2 in .env")
        return
 
    # Read KB numbers and URLs
    esm_links = []
    with open(esm_links_csv_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header
        esm_links = [row for row in reader if len(row) == 2]
 
    if not esm_links:
        print(" No valid entries found in ESM viewer links file.")
        return
 
    # Prepare output file
    os.makedirs("output", exist_ok=True)
    with open(extracted_links_csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["KB Number", "Link Text", "URL"])
 
    # Start Playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
 
        # Login once at the start
        print(" Opening ESM Home Page for Login")
        await page.goto("https://slbsandbox.service-now.com", wait_until="domcontentloaded")
 
        print(" Waiting for Microsoft login page...")
        await page.wait_for_url("https://login.microsoftonline.com/**", timeout=60000)
 
        # Username input
        print(" Entering username...")
        await page.fill('input[name="loginfmt"]', USERNAME2)
        await page.click('#idSIButton9')
 
        # Password input
        print(" Entering password...")
        await page.fill('input[name="passwd"]', PASSWORD2)
        await page.click('#idSIButton9')
 
        # "Stay signed in?" prompt
        try:
            await page.wait_for_selector('input[type="submit"][value="Yes"]', timeout=10000)
            await page.click('input[type="submit"][value="Yes"]')
        except Exception:
            print("No 'Stay signed in?' prompt detected.")
 
        # Proceed to extract links from ESM viewer pages
        for kb_number, url in esm_links:
            print(f"\n Opening ESM Viewer URL for KB {kb_number}: {url}")
            try:
                response = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                status = response.status if response else None
                if status != 200:
                    print(f" HTTP {status} loading KB {kb_number}")
                    continue
 
                # Wait for the viewer div
                await page.wait_for_selector('div#viewer.pdfViewer', timeout=10000)
                viewer = await page.query_selector('div#viewer.pdfViewer')
                links = await viewer.query_selector_all('a')
 
                with open(extracted_links_csv_path, mode='a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    for link in links:
                        href = await link.get_attribute("href") or ""
                        text = await link.inner_text() or ""
                        href, text = href.strip(), text.strip()
                        if href:
                            writer.writerow([kb_number, text, href])
                            print(f"🔗 Found: {text} → {href}")
 
            except Exception as e:
                print(f" Error processing KB {kb_number}: {e}")
 
        await browser.close()
 
    print(f"\n✅ Extracted links saved to: {extracted_links_csv_path}")
 
async def verify_esm_viewer_extracted_links():
    input_csv = "output/ESM_Viewer_Extracted_Links.csv"
 
    if not os.path.exists(input_csv):
        print(f" File not found: {input_csv}")
        return
 
    # Read links
    with open(input_csv, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        headers = next(reader)
        rows = [row for row in reader]
 
    # HTTP code to label
    def interpret_status(code):
        if 200 <= code < 300:
            return "OK"
        elif 300 <= code < 400:
            return "OK - Redirected"
        elif code in (401, 403, 405, 407, 423):
            return "OK - Access Required"
        else:
            return f"Broken ({code})"
 
    # Check each URL
    async with aiohttp.ClientSession() as session:
        updated_rows = []
        for row in rows:
            url = row[2].strip()
            if not url or url.startswith("mailto:") or url.startswith("#"):
                print(f" Skipped: {url}")
                continue
 
            try:
                async with session.head(url, timeout=10, allow_redirects=True) as response:
                    status = interpret_status(response.status)
            except Exception:
                try:
                    async with session.get(url, timeout=10, allow_redirects=True) as response:
                        status = interpret_status(response.status)
                except Exception as e:
                    status = f"Broken ({str(e).split()[0]})"
 
            updated_rows.append(row + [status])
            print(f" {url} → {status}")
 
    # Write with status
    new_csv = input_csv  # overwrite
    with open(new_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(headers + ["Status"])
        writer.writerows(updated_rows)
 
    print(f"\n ESM viewer links verified and updated → {new_csv}")
 
# Main
async def main():
    async with async_playwright() as p:
        # load_dotenv()
        # USERNAME2 = os.getenv("USERNAME2")
        # PASSWORD2 = os.getenv("PASSWORD2")
        # browser = await p.chromium.launch(headless=False)
        # context = await browser.new_context()
        # page = await context.new_page()
 
        # print("Opening ESM Home Page")
        # await page.goto("https://slbsandbox.service-now.com", wait_until="domcontentloaded")
        # #managing login
        # print(" Waiting for Microsoft login page...")
        # await page.wait_for_url("https://login.microsoftonline.com/**", timeout=60000)
 
        # # Fill in Microsoft email and proceed
        # print(" Entering username...")
        # # await page.get_by_role("textbox", name="someone@example.com").fill(USERNAME)
        # await page.fill('input[name="loginfmt"]', USERNAME2)
        # await page.click('#idSIButton9')
 
        # # Wait and fill in password
        # print(" Entering password...")
        # await page.fill('input[name="passwd"]', PASSWORD2)
        # await page.click('#idSIButton9')
 
        # #stay signed in page
        # await page.wait_for_selector('text="Stay signed in?"', timeout=100000)
        # await page.click('input[type="submit"][value="Yes"]')
 
        # await page.wait_for_selector('div.polaris-header.can-animate.polaris-enabled', timeout=20000)
 
 
        # jsession_id = await extract_jsession_id(context)
        # if not jsession_id:
        #     print("Could not extract jsession_id.")
        #     await browser.close()
        #     return
        # else:
        #     print("Cookies Extracted")  
        # await browser.close()
        kb_numbers = read_kb_numbers_from_excel(excel_path)
        if not kb_numbers:
            print("No KB numbers found in Excel.")
            return
 
        await scrape_kb_articles(kb_numbers)
        await verify_links_in_csv(output_csv_path)
        await extract_links_from_esm_viewer_pages()
        await verify_esm_viewer_extracted_links()
 
 
if __name__ == "__main__":
    asyncio.run(main())