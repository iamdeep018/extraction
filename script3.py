



# import os
# import asyncio
# import openpyxl
# import csv
# import aiohttp
# from dotenv import load_dotenv
# from playwright.async_api import async_playwright

# # Paths
# excel_path = "input/ESM KB.xlsx"
# output_csv_path = "output/KB_Hyperlinks.csv"
# screenshot_dir = "screenshots"
# viewer_extracted_csv_path = "output/ESM_Viewer_Extracted_Links.csv"
# broken_links_log_path = "output/Broken_Links_Log.csv"

# # Read KBs
# def read_kb_numbers_from_excel(path):
#     kb_numbers = []
#     try:
#         wb = openpyxl.load_workbook(path)
#         sheet = wb.active
#         for row in sheet.iter_rows(min_row=2, values_only=True):
#             kb_number = row[0]
#             if kb_number:
#                 kb_numbers.append(str(kb_number).strip())
#     except Exception as e:
#         print(f" Failed to read Excel: {e}")
#     return kb_numbers

# # Scraping Function
# async def scrape_kb_articles(kb_numbers):
#     load_dotenv()

#     # Use Chrome persistent context
#     user_data_dir = r"C:\Users\VV7\AppData\Local\Google\Chrome\User Data\Profile 2"
    

#     os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)
#     os.makedirs(screenshot_dir, exist_ok=True)

#     with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
#         writer = csv.writer(file)
#         writer.writerow(["KB Number", "Link Text", "URL"])

#     with open(viewer_extracted_csv_path, mode='w', newline='', encoding='utf-8') as file:
#         writer = csv.writer(file)
#         writer.writerow(["KB Number", "Link Text", "URL"])

#     async with async_playwright() as p:
#         browser = await p.chromium.launch_persistent_context(
#             user_data_dir=user_data_dir,
#             headless=False,
#             args=[ '--start-maximized']
#         )
#         page = await browser.new_page()

#         for kb_number in kb_numbers:
#             full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
#             print(f"\n Fetching KB: {kb_number} → {full_url}")

#             try:
#                 response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
#                 status = response.status if response else None
#                 current_url = page.url
#             except Exception as e:
#                 print(f" Error loading KB {kb_number}: {e}")
#                 await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
#                 continue

#             if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
#                 print(f" Redirected to login for KB {kb_number}")
#                 await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
#                 continue

#             if status == 200:
#                 print(f" Successfully authenticated and loaded KB {kb_number}")
#                 try:
#                     await page.wait_for_selector('article.kb-article-content', timeout=50000)
#                     article = await page.query_selector('article.kb-article-content')
#                     links = await article.query_selector_all('a')

#                     with open(output_csv_path, mode='a', newline='', encoding='utf-8') as file:
#                         writer = csv.writer(file)
#                         esm_viewer_links = []

#                         for link in links:
#                             href = await link.get_attribute("href") or ""
#                             text = await link.inner_text() or ""
#                             if href.strip() and not href.startswith("mailto:") and not href.startswith("#"):
#                                 writer.writerow([kb_number, text.strip(), href.strip()])
#                                 print(f"- {text.strip()}: {href.strip()}")
#                                 if href.startswith("https://esm.slb.com/$viewer.do"):
#                                     esm_viewer_links.append(href)

#                     for viewer_url in esm_viewer_links:
#                         print(f"  Visiting ESM Viewer: {viewer_url}")
#                         try:
#                             await page.goto(viewer_url, wait_until="domcontentloaded", timeout=60000)
#                             await page.wait_for_selector('div#viewer.pdfViewer', timeout=10000)
#                             viewer = await page.query_selector('div#viewer.pdfViewer')
#                             if viewer:
#                                 viewer_links = await viewer.query_selector_all('a')
#                                 with open(viewer_extracted_csv_path, mode='a', newline='', encoding='utf-8') as file:
#                                     writer = csv.writer(file)
#                                     for vlink in viewer_links:
#                                         vhref = await vlink.get_attribute("href") or ""
#                                         vtext = await vlink.inner_text() or ""
#                                         if vhref.strip():
#                                             writer.writerow([kb_number, vtext.strip(), vhref.strip()])
#                                             print(f"   → {vtext.strip()} | {vhref.strip()}")
#                         except Exception as e:
#                             print(f"  Error loading ESM Viewer: {e}")
#                 except Exception as e:
#                     print(f" Error extracting links from KB {kb_number}: {e}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
#             else:
#                 print(f" HTTP {status} error for KB {kb_number}")

#         await browser.close()
#         print("\n Completed scraping KB articles and embedded viewer links.")

# # Verify KB hyperlinks
# async def verify_links_in_csv(csv_path):
#     print("\n Verifying hyperlinks in the CSV file...")
#     broken_links = []
#     rows = []

#     with open(csv_path, mode='r', newline='', encoding='utf-8') as file:
#         reader = csv.reader(file)
#         headers = next(reader)
#         for row in reader:
#             rows.append(row)

#     def interpret_status(code):
#         if 200 <= code < 300:
#             return "OK"
#         elif 300 <= code < 400:
#             return "OK - Redirected"
#         elif code in (401, 403, 405, 407, 423):
#             return "OK - Access Required"
#         else:
#             return f"Broken ({code})"

#     async with aiohttp.ClientSession() as session:
#         updated_rows = []
#         for row in rows:
#             url = row[2].strip()
#             kb_number, text = row[0], row[1]
#             if not url or url.startswith("mailto:") or url.startswith("#"):
#                 continue
#             try:
#                 async with session.head(url, timeout=10, allow_redirects=True) as response:
#                     status = interpret_status(response.status)
#             except:
#                 try:
#                     async with session.get(url, timeout=10, allow_redirects=True) as response:
#                         status = interpret_status(response.status)
#                 except Exception as e:
#                     status = f"Broken ({str(e).split()[0]})"

#             updated_rows.append(row + [status])
#             print(f" {url} → {status}")
#             if status.startswith("Broken"):
#                 broken_links.append([kb_number, text, url, status])

#     with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
#         writer = csv.writer(file)
#         writer.writerow(headers + ["Status"])
#         writer.writerows(updated_rows)

#     if broken_links:
#         with open(broken_links_log_path, mode='w', newline='', encoding='utf-8') as file:
#             writer = csv.writer(file)
#             writer.writerow(["KB Number", "Link Text", "URL", "Status"])
#             writer.writerows(broken_links)
#         print(f" Broken links saved → {broken_links_log_path}")

#     print(f"\n Hyperlink verification complete. Updated: {csv_path}")

# # Main
# async def main():
#     kb_numbers = read_kb_numbers_from_excel(excel_path)
#     if not kb_numbers:
#         print("No KB numbers found in Excel.")
#         return

#     await scrape_kb_articles(kb_numbers)
#     await verify_links_in_csv(output_csv_path)
#     await verify_links_in_csv(viewer_extracted_csv_path)

# if __name__ == "__main__":
#     asyncio.run(main())









# import os
# import asyncio
# import openpyxl
# import csv
# import aiohttp
# from dotenv import load_dotenv
# from playwright.async_api import async_playwright

# # Paths
# input_folder = "input"
# screenshot_dir = "screenshots"
# os.makedirs(screenshot_dir, exist_ok=True)

# # Extract KB numbers
# def read_kb_numbers_from_excel(file_path):
#     kb_numbers = []
#     try:
#         wb = openpyxl.load_workbook(file_path)
#         sheet = wb.active
#         headers = [cell.value for cell in sheet[1]]
#         if "Number" in headers:
#             number_index = headers.index("Number")
#             for row in sheet.iter_rows(min_row=2, values_only=True):
#                 kb_number = row[number_index]
#                 if kb_number:
#                     kb_numbers.append(str(kb_number).strip())
#         else:
#             print(f"Skipping {file_path}: 'Number' column not found.")
#     except Exception as e:
#         print(f"Failed to read {file_path}: {e}")
#     return kb_numbers

# # Verify link status
# async def verify_link_status(session, url):
#     try:
#         async with session.head(url, timeout=10, allow_redirects=True) as response:
#             return response.status
#     except:
#         try:
#             async with session.get(url, timeout=10, allow_redirects=True) as response:
#                 return response.status
#         except:
#             return None

# def interpret_status(code):
#     if code is None:
#         return "Broken (Timeout)"
#     elif 200 <= code < 300:
#         return "OK"
#     elif 300 <= code < 400:
#         return "OK - Redirected"
#     elif code in (401, 403, 405, 407, 423):
#         return "OK - Access Required"
#     else:
#         return f"Broken ({code})"

# # Scraping function
# async def scrape_kb_articles(kb_numbers, output_csv_path):
#     user_data_dir = r"C:\\Users\\VV7\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2"
#     os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

#     async with aiohttp.ClientSession() as session:
#         async with async_playwright() as p:
#             browser = await p.chromium.launch_persistent_context(
#                 user_data_dir=user_data_dir,
#                 headless=False,
#                 args=['--start-maximized']
#             )
#             page = await browser.new_page()

#             with open(output_csv_path, 'w', newline='', encoding='utf-8') as f:
#                 writer = csv.writer(f)
#                 writer.writerow(["KB Number", "Link Text", "URL", "is_Attachment", "inside_Attachment", "Status"])

#             for kb_number in kb_numbers:
#                 full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
#                 print(f"\nFetching KB: {kb_number} → {full_url}")
#                 try:
#                     response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
#                     status = response.status if response else None
#                     current_url = page.url
#                 except Exception as e:
#                     print(f"Error loading KB {kb_number}: {e}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
#                     continue

#                 if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
#                     print(f"Redirected to login for KB {kb_number}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
#                     continue

#                 if status == 200:
#                     try:
#                         await page.wait_for_selector('article.kb-article-content', timeout=50000)
#                         article = await page.query_selector('article.kb-article-content')
#                         links = await article.query_selector_all('a')
#                         esm_viewer_links = []

#                         for link in links:
#                             href = await link.get_attribute("href") or ""
#                             text = await link.inner_text() or ""
#                             if href.strip() and not href.startswith("mailto:") and not href.startswith("#"):
#                                 is_attachment = "Yes" if href.startswith("https://esm.slb.com/$viewer.do") else "No"
#                                 link_status = interpret_status(await verify_link_status(session, href.strip()))
#                                 with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                     csv.writer(f).writerow([kb_number, text.strip(), href.strip(), is_attachment, "No", link_status])
#                                 if is_attachment == "Yes":
#                                     esm_viewer_links.append(href)

#                         for viewer_url in esm_viewer_links:
#                             try:
#                                 await page.goto(viewer_url, wait_until="domcontentloaded", timeout=60000)
#                                 await page.wait_for_selector('iframe', timeout=10000)
#                                 frames = page.frames
#                                 for frame in frames:
#                                     try:
#                                         viewer_links = await frame.query_selector_all('a')
#                                         for vlink in viewer_links:
#                                             vhref = await vlink.get_attribute("href") or ""
#                                             vtext = await vlink.inner_text() or ""
#                                             if vhref.strip():
#                                                 vlink_status = interpret_status(await verify_link_status(session, vhref.strip()))
#                                                 with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                                     csv.writer(f).writerow([kb_number, vtext.strip(), vhref.strip(), "Yes", "Yes", vlink_status])
#                                     except:
#                                         continue
#                             except Exception as e:
#                                 print(f"  Error loading ESM Viewer: {e}")
#                     except Exception as e:
#                         await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
#                 else:
#                     print(f"HTTP {status} error for KB {kb_number}")

#             await browser.close()

# # Main loop
# async def main():
#     load_dotenv()
#     for filename in os.listdir(input_folder):
#         if filename.endswith(".xlsx"):
#             excel_path = os.path.join(input_folder, filename)
#             base_name = os.path.splitext(filename)[0].replace(" ", "_")
#             output_csv = f"output/{base_name}_All_Links.csv"

#             print(f"\nProcessing file: {filename}")
#             kb_numbers = read_kb_numbers_from_excel(excel_path)
#             if not kb_numbers:
#                 print(f"No KBs found in {filename}")
#                 continue

#             await scrape_kb_articles(kb_numbers, output_csv)

# if __name__ == "__main__":
#     asyncio.run(main())

# import os
# import asyncio
# import openpyxl
# import csv
# import aiohttp
# from dotenv import load_dotenv
# from playwright.async_api import async_playwright

# # Paths
# input_folder = "input"
# screenshot_dir = "screenshots"
# os.makedirs(screenshot_dir, exist_ok=True)

# # Extract KB numbers
# def read_kb_numbers_from_excel(file_path):
#     kb_numbers = []
#     try:
#         wb = openpyxl.load_workbook(file_path)
#         sheet = wb.active
#         headers = [cell.value for cell in sheet[1]]
#         if "Number" in headers:
#             number_index = headers.index("Number")
#             for row in sheet.iter_rows(min_row=2, values_only=True):
#                 kb_number = row[number_index]
#                 if kb_number:
#                     kb_numbers.append(str(kb_number).strip())
#         else:
#             print(f"Skipping {file_path}: 'Number' column not found.")
#     except Exception as e:
#         print(f"Failed to read {file_path}: {e}")
#     return kb_numbers

# # Verify link status
# async def verify_link_status(session, url):
#     try:
#         async with session.head(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
#             return response.status
#     except aiohttp.ClientError:
#         return "Broken (Client Error)"
#     except asyncio.TimeoutError:
#         return "Broken (Timeout)"
#     except:
#         try:
#             async with session.get(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
#                 return response.status
#         except aiohttp.ClientError:
#             return "Broken (Client Error)"
#         except asyncio.TimeoutError:
#             return "Broken (Timeout)"
#         except:
#             return "Broken (Unknown Error)"

# # Function to map HTTP status to label
# def interpret_status(code):
#     if isinstance(code, str):
#         return code
#     elif 200 <= code < 300:
#         return "OK"
#     elif 300 <= code < 400:
#         return "OK - Redirected"
#     elif code in (401, 403, 405, 407, 423):
#         return "OK - Access Required"
#     else:
#         return f"Broken ({code})"

# # Check exclusion
# def is_excluded_link(href):
#     excluded_keywords = [".png", ".jpg", ".jpeg", ".svg", ".gif", "javascript:", ".css", ".js"]
#     return any(keyword in href.lower() for keyword in excluded_keywords)

# # Scraping function
# async def scrape_kb_articles(kb_numbers, output_csv_path):
#     user_data_dir = r"C:\\Users\\VV7\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2"
#     os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

#     async with aiohttp.ClientSession() as session:
#         async with async_playwright() as p:
#             browser = await p.chromium.launch_persistent_context(
#                 user_data_dir=user_data_dir,
#                 headless=False,
#                 args=['--start-maximized']
#             )
#             page = await browser.new_page()

#             with open(output_csv_path, 'w', newline='', encoding='utf-8') as f:
#                 writer = csv.writer(f)
#                 writer.writerow(["KB Number", "Link Text", "URL", "is_Attachment", "inside_Attachment", "Status"])

#             for kb_number in kb_numbers:
#                 full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
#                 print(f"\nFetching KB: {kb_number} → {full_url}")
#                 try:
#                     response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
#                     status = response.status if response else None
#                     current_url = page.url
#                 except Exception as e:
#                     print(f"Error loading KB {kb_number}: {e}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
#                     continue

#                 if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
#                     print(f"Redirected to login for KB {kb_number}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
#                     continue

#                 if status == 200:
#                     try:
#                         await page.wait_for_selector('article.kb-article-content', timeout=50000)
#                         article = await page.query_selector('article.kb-article-content')
#                         links = await article.query_selector_all('a')
#                         esm_viewer_links = []

#                         for link in links:
#                             href = await link.get_attribute("href") or ""
#                             text = await link.inner_text() or ""
#                             if href.strip() and not href.startswith("mailto:") and not href.startswith("#") and not is_excluded_link(href):
#                                 is_attachment = "Yes" if href.startswith("https://esm.slb.com/$viewer.do") else "No"
#                                 raw_status = await verify_link_status(session, href.strip())
#                                 link_status = interpret_status(raw_status)
#                                 with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                     csv.writer(f).writerow([kb_number, text.strip(), href.strip(), is_attachment, "No", link_status])
#                                 if is_attachment == "Yes":
#                                     esm_viewer_links.append(href)

#                         for viewer_url in esm_viewer_links:
#                             try:
#                                 await page.goto(viewer_url, wait_until="domcontentloaded", timeout=60000)
#                                 await page.wait_for_selector('iframe', timeout=10000)
#                                 frames = page.frames
#                                 for frame in frames:
#                                     try:
#                                         viewer_links = await frame.query_selector_all('a')
#                                         for vlink in viewer_links:
#                                             vhref = await vlink.get_attribute("href") or ""
#                                             vtext = await vlink.inner_text() or ""
#                                             if vhref.strip() and not vhref.startswith("mailto:") and not vhref.startswith("#") and not is_excluded_link(vhref):
#                                                 raw_status = await verify_link_status(session, vhref.strip())
#                                                 vlink_status = interpret_status(raw_status)
#                                                 with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                                     csv.writer(f).writerow([kb_number, vtext.strip(), vhref.strip(), "Yes", "Yes", vlink_status])
#                                     except:
#                                         continue
#                             except Exception as e:
#                                 print(f"  Error loading ESM Viewer: {e}")
#                     except Exception as e:
#                         await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
#                 else:
#                     print(f"HTTP {status} error for KB {kb_number}")

#             await browser.close()

# # Main loop
# async def main():
#     load_dotenv()
#     for filename in os.listdir(input_folder):
#         if filename.endswith(".xlsx"):
#             excel_path = os.path.join(input_folder, filename)
#             base_name = os.path.splitext(filename)[0].replace(" ", "_")
#             output_csv = f"output/{base_name}_All_Links.csv"

#             print(f"\nProcessing file: {filename}")
#             kb_numbers = read_kb_numbers_from_excel(excel_path)
#             if not kb_numbers:
#                 print(f"No KBs found in {filename}")
#                 continue

#             await scrape_kb_articles(kb_numbers, output_csv)

# if __name__ == "__main__":
#     asyncio.run(main())



# import os
# import asyncio
# import openpyxl
# import csv
# import aiohttp
# from dotenv import load_dotenv
# from playwright.async_api import async_playwright

# # Paths
# input_folder = "input"
# screenshot_dir = "screenshots"
# os.makedirs(screenshot_dir, exist_ok=True)

# # Extract KB numbers
# def read_kb_numbers_from_excel(file_path):
#     kb_numbers = []
#     try:
#         wb = openpyxl.load_workbook(file_path)
#         sheet = wb.active
#         headers = [cell.value for cell in sheet[1]]
#         if "Number" in headers:
#             number_index = headers.index("Number")
#             for row in sheet.iter_rows(min_row=2, values_only=True):
#                 kb_number = row[number_index]
#                 if kb_number:
#                     kb_numbers.append(str(kb_number).strip())
#         else:
#             print(f"Skipping {file_path}: 'Number' column not found.")
#     except Exception as e:
#         print(f"Failed to read {file_path}: {e}")
#     return kb_numbers

# # Verify link status
# async def verify_link_status(session, url):
#     try:
#         async with session.head(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
#             return response.status
#     except aiohttp.ClientError:
#         return "Broken (Client Error)"
#     except asyncio.TimeoutError:
#         return "Broken (Timeout)"
#     except:
#         try:
#             async with session.get(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
#                 return response.status
#         except aiohttp.ClientError:
#             return "Broken (Client Error)"
#         except asyncio.TimeoutError:
#             return "Broken (Timeout)"
#         except:
#             return "Broken (Unknown Error)"

# # Function to map HTTP status to label
# def interpret_status(code):
#     if isinstance(code, str):
#         return code
#     elif 200 <= code < 300:
#         return "OK"
#     elif 300 <= code < 400:
#         return "OK - Redirected"
#     elif code in (401, 403, 405, 407, 423):
#         return "OK - Access Required"
#     else:
#         return f"Broken ({code})"

# # Check exclusion
# def is_excluded_link(href):
#     excluded_keywords = [".png", ".jpg", ".jpeg", ".svg", ".gif", "javascript:", ".css", ".js"]
#     return any(keyword in href.lower() for keyword in excluded_keywords)

# # Scraping function
# async def scrape_kb_articles(kb_numbers, output_csv_path):
#     user_data_dir = r"C:\\Users\\VV7\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2"
#     os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

#     async with aiohttp.ClientSession() as session:
#         async with async_playwright() as p:
#             browser = await p.chromium.launch_persistent_context(
#                 user_data_dir=user_data_dir,
#                 headless=False,
#                 args=['--start-maximized']
#             )
#             page = await browser.new_page()

#             with open(output_csv_path, 'w', newline='', encoding='utf-8') as f:
#                 writer = csv.writer(f)
#                 writer.writerow(["KB Number", "Link Text", "URL", "is_Attachment", "inside_Attachment", "Status"])

#             for kb_number in kb_numbers:
#                 full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
#                 print(f"\nFetching KB: {kb_number} → {full_url}")
#                 try:
#                     response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
#                     status = response.status if response else None
#                     current_url = page.url
#                 except Exception as e:
#                     print(f"Error loading KB {kb_number}: {e}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
#                     continue

#                 if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
#                     print(f"Redirected to login for KB {kb_number}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
#                     continue

#                 if status == 200:
#                     try:
#                         await page.wait_for_selector('article.kb-article-content', timeout=50000)
#                         article = await page.query_selector('article.kb-article-content')
#                         links = await article.query_selector_all('a')
#                         esm_viewer_links = []

#                         for link in links:
#                             href = await link.get_attribute("href") or ""
#                             text = await link.inner_text() or ""
#                             href = href.strip()
#                             if not href or href.startswith("mailto:") or href.startswith("#") or is_excluded_link(href):
#                                 print(f"Skipped: {href}")
#                                 continue
#                             is_attachment = "Yes" if "$viewer.do" in href else "No"
#                             raw_status = await verify_link_status(session, href)
#                             link_status = interpret_status(raw_status)
#                             with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                 csv.writer(f).writerow([kb_number, text.strip(), href, is_attachment, "No", link_status])
#                             if is_attachment == "Yes":
#                                 esm_viewer_links.append(href)

#                         for viewer_url in esm_viewer_links:
#                             try:
#                                 await page.goto(viewer_url, wait_until="domcontentloaded", timeout=60000)
#                                 await page.wait_for_selector('iframe', timeout=10000)
#                                 frames = page.frames
#                                 for frame in frames:
#                                     try:
#                                         viewer_links = await frame.query_selector_all('a')
#                                         for vlink in viewer_links:
#                                             vhref = await vlink.get_attribute("href") or ""
#                                             vtext = await vlink.inner_text() or ""
#                                             vhref = vhref.strip()
#                                             if not vhref or vhref.startswith("mailto:") or vhref.startswith("#") or is_excluded_link(vhref):
#                                                 continue
#                                             raw_status = await verify_link_status(session, vhref)
#                                             vlink_status = interpret_status(raw_status)
#                                             with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                                 csv.writer(f).writerow([kb_number, vtext.strip(), vhref, "Yes", "Yes", vlink_status])
#                                     except:
#                                         continue
#                             except Exception as e:
#                                 print(f"  Error loading ESM Viewer: {e}")
#                     except Exception as e:
#                         await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
#                 else:
#                     print(f"HTTP {status} error for KB {kb_number}")

#             await browser.close()

# # Main loop
# async def main():
#     load_dotenv()
#     for filename in os.listdir(input_folder):
#         if filename.endswith(".xlsx"):
#             excel_path = os.path.join(input_folder, filename)
#             base_name = os.path.splitext(filename)[0].replace(" ", "_")
#             output_csv = f"output/{base_name}_All_Links.csv"

#             print(f"\nProcessing file: {filename}")
#             kb_numbers = read_kb_numbers_from_excel(excel_path)
#             if not kb_numbers:
#                 print(f"No KBs found in {filename}")
#                 continue

#             await scrape_kb_articles(kb_numbers, output_csv)

# if __name__ == "__main__":
#     asyncio.run(main())





# import os
# import asyncio
# import openpyxl
# import csv
# import aiohttp
# from dotenv import load_dotenv
# from playwright.async_api import async_playwright

# # Paths
# input_folder = "input"
# screenshot_dir = "screenshots"
# os.makedirs(screenshot_dir, exist_ok=True)

# # Extract KB numbers
# def read_kb_numbers_from_excel(file_path):
#     kb_numbers = []
#     try:
#         wb = openpyxl.load_workbook(file_path)
#         sheet = wb.active
#         headers = [cell.value for cell in sheet[1]]
#         if "Number" in headers:
#             number_index = headers.index("Number")
#             for row in sheet.iter_rows(min_row=2, values_only=True):
#                 kb_number = row[number_index]
#                 if kb_number:
#                     kb_numbers.append(str(kb_number).strip())
#         else:
#             print(f"Skipping {file_path}: 'Number' column not found.")
#     except Exception as e:
#         print(f"Failed to read {file_path}: {e}")
#     return kb_numbers

# # Verify link status
# async def verify_link_status(session, url):
#     try:
#         async with session.head(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
#             return response.status
#     except aiohttp.ClientError:
#         return "Broken (Client Error)"
#     except asyncio.TimeoutError:
#         return "Broken (Timeout)"
#     except:
#         try:
#             async with session.get(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
#                 return response.status
#         except aiohttp.ClientError:
#             return "Broken (Client Error)"
#         except asyncio.TimeoutError:
#             return "Broken (Timeout)"
#         except:
#             return "Broken (Unknown Error)"

# # Map status code
# def interpret_status(code):
#     if isinstance(code, str):
#         return code
#     elif 200 <= code < 300:
#         return "OK"
#     elif 300 <= code < 400:
#         return "OK - Redirected"
#     elif code in (401, 403, 405, 407, 423):
#         return "OK - Access Required"
#     else:
#         return f"Broken ({code})"

# # Skip certain file types
# def is_excluded_link(href):
#     excluded_keywords = [".png", ".jpg", ".jpeg", ".svg", ".gif", "javascript:", ".css", ".js"]
#     return any(keyword in href.lower() for keyword in excluded_keywords)

# # Extract links from all frames (for inside ESM viewers)
# async def extract_links_from_all_frames(page, session, kb_number, output_csv_path):
#     links_found = 0
#     for frame in page.frames:
#         try:
#             anchors = await frame.query_selector_all('a')
#             for a in anchors:
#                 href = await a.get_attribute("href") or ""
#                 text = await a.inner_text() or ""
#                 href = href.strip()
#                 if not href or href.startswith("mailto:") or href.startswith("#") or is_excluded_link(href):
#                     continue
#                 raw_status = await verify_link_status(session, href)
#                 link_status = interpret_status(raw_status)
#                 with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                     csv.writer(f).writerow([kb_number, text.strip(), href, "Yes", "Yes", link_status])
#                 links_found += 1
#         except Exception as e:
#             print(f"Frame error: {e}")
#             continue
#     return links_found

# # Main scraping logic
# async def scrape_kb_articles(kb_numbers, output_csv_path):
#     user_data_dir = r"C:\\Users\\VV7\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2"
#     os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

#     async with aiohttp.ClientSession() as session:
#         async with async_playwright() as p:
#             browser = await p.chromium.launch_persistent_context(
#                 user_data_dir=user_data_dir,
#                 headless=False,
#                 args=['--start-maximized']
#             )
#             page = await browser.new_page()

#             with open(output_csv_path, 'w', newline='', encoding='utf-8') as f:
#                 writer = csv.writer(f)
#                 writer.writerow(["KB Number", "Link Text", "URL", "is_Attachment", "inside_Attachment", "Status"])

#             for kb_number in kb_numbers:
#                 full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
#                 print(f"\nFetching KB: {kb_number} → {full_url}")
#                 try:
#                     response = await page.goto(full_url, wait_until="domcontentloaded", timeout=60000)
#                     status = response.status if response else None
#                     current_url = page.url
#                 except Exception as e:
#                     print(f"Error loading KB {kb_number}: {e}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_error.png")
#                     continue

#                 if "login.microsoftonline.com" in current_url or "login_with_sso.do" in current_url:
#                     print(f"Redirected to login for KB {kb_number}")
#                     await page.screenshot(path=f"{screenshot_dir}/{kb_number}_login_redirect.png")
#                     continue

#                 if status == 200:
#                     try:
#                         await page.wait_for_selector('article.kb-article-content', timeout=50000)
#                         article = await page.query_selector('article.kb-article-content')
#                         links = await article.query_selector_all('a')
#                         esm_viewer_links = []

#                         for link in links:
#                             href = await link.get_attribute("href") or ""
#                             text = await link.inner_text() or ""
#                             href = href.strip()
#                             if not href or href.startswith("mailto:") or href.startswith("#") or is_excluded_link(href):
#                                 print(f"Skipped: {href}")
#                                 continue
#                             is_attachment = "Yes" if "$viewer.do" in href else "No"
#                             raw_status = await verify_link_status(session, href)
#                             link_status = interpret_status(raw_status)
#                             with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
#                                 csv.writer(f).writerow([kb_number, text.strip(), href, is_attachment, "No", link_status])
#                             if is_attachment == "Yes":
#                                 esm_viewer_links.append(href)

#                         for viewer_url in esm_viewer_links:
#                             try:
#                                 print(f"  Opening viewer: {viewer_url}")
#                                 await page.goto(viewer_url, wait_until="domcontentloaded", timeout=60000)
#                                 await page.wait_for_selector('iframe', timeout=10000)
#                                 count = await extract_links_from_all_frames(page, session, kb_number, output_csv_path)
#                                 if count == 0:
#                                     print(f"  ⚠️ No embedded links found inside viewer.")
#                             except Exception as e:
#                                 print(f"  Error loading ESM Viewer: {e}")
#                     except Exception as e:
#                         await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
#                 else:
#                     print(f"HTTP {status} error for KB {kb_number}")

#             await browser.close()

# # Main loop
# async def main():
#     load_dotenv()
#     for filename in os.listdir(input_folder):
#         if filename.endswith(".xlsx"):
#             excel_path = os.path.join(input_folder, filename)
#             base_name = os.path.splitext(filename)[0].replace(" ", "_")
#             output_csv = f"output/{base_name}_All_Links.csv"

#             print(f"\nProcessing file: {filename}")
#             kb_numbers = read_kb_numbers_from_excel(excel_path)
#             if not kb_numbers:
#                 print(f"No KBs found in {filename}")
#                 continue

#             await scrape_kb_articles(kb_numbers, output_csv)

# if __name__ == "__main__":
#     asyncio.run(main())
import os
import asyncio
import openpyxl
import csv
import aiohttp
import re
from dotenv import load_dotenv
from playwright.async_api import async_playwright

# Setup
input_folder = "input"
screenshot_dir = "screenshots"
os.makedirs(screenshot_dir, exist_ok=True)
URL_PATTERN = re.compile(r'https?://[^\s<>"\']+|www\.[^\s<>"\']+')

# Excel KB reader
def read_kb_numbers_from_excel(file_path):
    kb_numbers = []
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]
        if "Number" in headers:
            number_index = headers.index("Number")
            for row in sheet.iter_rows(min_row=2, values_only=True):
                kb_number = row[number_index]
                if kb_number:
                    kb_numbers.append(str(kb_number).strip())
        else:
            print(f"Skipping {file_path}: 'Number' column not found.")
    except Exception as e:
        print(f"Failed to read {file_path}: {e}")
    return kb_numbers

# URL status check
async def verify_link_status(session, url):
    try:
        async with session.head(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
            return response.status
    except Exception:
        try:
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as response:
                return response.status
        except Exception:
            return "Broken (Unknown Error)"

def interpret_status(code):
    if isinstance(code, str):
        return code
    elif 200 <= code < 300:
        return "OK"
    elif 300 <= code < 400:
        return "OK - Redirected"
    elif code in (401, 403, 405, 407, 423):
        return "OK - Access Required"
    else:
        return f"Broken ({code})"

def is_excluded_link(href):
    excluded_keywords = [".png", ".jpg", ".jpeg", ".svg", ".gif", "javascript:", ".css", ".js"]
    return any(keyword in href.lower() for keyword in excluded_keywords)

# Core scraping
async def scrape_kb_articles(kb_numbers, output_csv_path):
    user_data_dir = r"C:\\Users\\VV7\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2"
    os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

    async with aiohttp.ClientSession() as session:
        async with async_playwright() as p:
            browser = await p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                args=['--start-maximized']
            )
            page = await browser.new_page()

            with open(output_csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["KB Number", "Link Text", "URL", "is_Attachment", "inside_Attachment", "Status"])

            for kb_number in kb_numbers:
                full_url = f"https://slbsandbox.service-now.com/esc?id=kb_article&sysparm_article={kb_number}"
                print(f"\n📘 Fetching KB: {kb_number} → {full_url}")
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
                    try:
                        await page.wait_for_selector('article.kb-article-content', timeout=50000)
                        article = await page.query_selector('article.kb-article-content')
                        links = await article.query_selector_all('a')
                        esm_viewer_links = []

                        for link in links:
                            href = await link.get_attribute("href") or ""
                            text = await link.inner_text() or ""
                            href = href.strip()
                            if not href or href.startswith("mailto:") or href.startswith("#") or is_excluded_link(href):
                                continue
                            is_attachment = "Yes" if "$viewer.do" in href else "No"
                            raw_status = await verify_link_status(session, href)
                            link_status = interpret_status(raw_status)
                            with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
                                csv.writer(f).writerow([kb_number, text.strip(), href, is_attachment, "No", link_status])
                            if is_attachment == "Yes":
                                esm_viewer_links.append(href)

                        # Process viewer PDF content
                        for viewer_url in esm_viewer_links:
                            try:
                                print(f"🔍 Opening viewer: {viewer_url}")
                                await page.goto(viewer_url, wait_until="domcontentloaded", timeout=60000)
                                await page.wait_for_selector('xpath=/html/body/div[1]/div[2]/div[8]/div/div[1]/div[2]', timeout=30000)
                                viewer_body = await page.query_selector('xpath=/html/body/div[1]/div[2]/div[8]/div/div[1]/div[2]')
                                viewer_links = await viewer_body.query_selector_all('a')

                                for a in viewer_links:
                                    href = (await a.get_attribute("href") or "").strip()
                                    text = (await a.inner_text() or "").strip()
                                    if href and not is_excluded_link(href):
                                        status = interpret_status(await verify_link_status(session, href))
                                        with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
                                            csv.writer(f).writerow([kb_number, text, href, "Yes", "Yes", status])

                                spans = await page.query_selector_all("#viewer .textLayer span")
                                for span in spans:
                                    span_text = await span.inner_text()
                                    matches = URL_PATTERN.findall(span_text)
                                    for match in matches:
                                        clean_url = match.strip().replace("\u202f", "").replace(" ", "")
                                        if clean_url.startswith("www."):
                                            clean_url = "https://" + clean_url
                                        raw_status = await verify_link_status(session, clean_url)
                                        status = interpret_status(raw_status)
                                        with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
                                            csv.writer(f).writerow([kb_number, clean_url, clean_url, "Yes", "Yes", status])

                                # New: process custom viewer div section
                                custom_section = await page.query_selector('div.panel-body.m-b-lg.wrapper-lg')
                                if custom_section:
                                    custom_links = await custom_section.query_selector_all('a')
                                for link in custom_links:
                                    href = (await link.get_attribute("href") or "").strip()
                                    text = (await link.inner_text() or "").strip()
                                    if href and not is_excluded_link(href):
                                        status = interpret_status(await verify_link_status(session, href))
                                        with open(output_csv_path, 'a', newline='', encoding='utf-8') as f:
                                            csv.writer(f).writerow([kb_number, text, href, "Yes", "Yes", status])

                            except Exception as e:
                                print(f"❌ Viewer error: {e}")
                                await page.screenshot(path=f"{screenshot_dir}/{kb_number}_viewer_error.png")
                    except Exception as e:
                        print(f"❌ Scraping error in KB {kb_number}: {e}")
                        await page.screenshot(path=f"{screenshot_dir}/{kb_number}_link_error.png")
                else:
                    print(f"⚠️ HTTP {status} error for KB {kb_number}")

            await browser.close()

# Entry
async def main():
    load_dotenv()
    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx"):
            excel_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0].replace(" ", "_")
            output_csv = f"output/{base_name}_All_Links.csv"
            print(f"\n📄 Processing file: {filename}")
            kb_numbers = read_kb_numbers_from_excel(excel_path)
            if not kb_numbers:
                print(f"⚠️ No KBs found in {filename}")
                continue
            await scrape_kb_articles(kb_numbers, output_csv)

if __name__ == "__main__":
    asyncio.run(main())
