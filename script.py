import pandas as pd
from playwright.sync_api import sync_playwright
from string.templatelib import Template
import traceback
import re

options = ["CP", "BP", "Regional"]

def run_automation():
    print("Step 1: Starting Playwright...")
    with sync_playwright() as p:
        # Launch browser (headless=False so you can watch it work)
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()
        df = pd.read_excel("MASTERSHEET_FOR_DAILY_DMS_ENTERIES_NEW.xlsx")
        print(f"Excel loaded! Found {len(df)} rows.")

        try:
            # 1. Login Phase
            page.goto("/users/login")
            page.get_by_role("textbox", name="Username").fill("")
            page.get_by_role("textbox", name="Password").fill("")
            page.get_by_role("button", name="Login").click()

            # Wait for navigation to complete after login
            page.wait_for_load_state("networkidle")

            # 2. Data Entry Loop
            print("excel rows-->")
            print(df.iterrows())
            grouped = df.groupby('Voucher No.')

            for voucher_no, items in grouped:
                print(f"Processing Voucher: {voucher_no} ({len(items)} items)")
                first_row = items.iloc[0]

                page.goto("/payments/sale")

                page.get_by_role("textbox", name="Select Outlet").click()
                page.get_by_role("searchbox", name="Search").nth(2).fill(str(first_row['PARTY CODE']))
                page.get_by_role("option", name=str(first_row['PARTY CODE'])).click()

                for index, row in items.iterrows():
                    print("row-->", row)
                    # for name in options:
                    #     button = page.get_by_role("button", name=name)
                    #     if button.is_visible():
                    #         button.click()
                    #         break

                    page.get_by_role("button", name=re.compile(r"CP|BP|Regional", re.IGNORECASE)).click()
                    page.locator("a").filter(has_text=row['CP/BP/REG']).click()
                    page.get_by_role("searchbox", name="Search:").fill(str(row['A.code']))                    
                    page.get_by_role("heading", name="Step 1: Account Details").click()
                    page.wait_for_load_state("networkidle")
                    
                    sku_id = page.locator("#skunitstable").locator("tbody").locator("tr").first.locator("td").first.inner_text()
                    print('sku_id', sku_id)
                    page.locator("#price_" + sku_id).fill(str(row['PER UNIT SALE PRICE']))
                    page.locator("#cases_" + sku_id).fill(str(row['Quantity (Case)']))

                page.get_by_role("link", name="Next").first.click()
                page.get_by_role("link", name="Next").click()
                page.close()

            print("Finished All rows!")

        except Exception as e:
            print(f"CRITICAL ERROR: {e}")
            # 3. Save the Trace file ONLY if it fails
            context.tracing.stop(path="error_trace.zip")
            print("Trace saved to error_trace.zip. View it at trace.playwright.dev")
            traceback.print_exc() # Prints the line number of the error
            
        finally:
            browser.close()

if __name__ == "__main__":
    run_automation()
    print("Script finished successfully.")