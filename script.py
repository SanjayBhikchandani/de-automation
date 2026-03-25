import pandas as pd
from playwright.sync_api import sync_playwright
# from string.templatelib import Template # py command
from string import Template # python command
import traceback
import re
from datetime import datetime

options = ["CP", "BP", "Regional"]

def run_automation():
    print("Step 1: Starting Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()
        df = pd.read_excel("MASTERSHEET_FOR_DAILY_DMS_ENTERIES_NEW.xlsx")
        print(f"Excel loaded! Found {len(df)} rows.")

        if 'Status' not in df.columns:
            df['Status'] = 'Pending'

        if 'Processed Date' not in df.columns:
            df['Processed Date'] = ''

        try:
            page.goto("/users/login")

            # page.fill('input[name="identifier"]', "")
            #
            # if page.is_visible('input[name="credentials.passcode"]'):
            #     page.fill('input[name="credentials.passcode"]', "")
            # else:
            #     page.click('input[type="submit"]')
            #     page.fill('input[name="credentials.passcode"]', "")
            #
            # page.click('input[type="submit"]')

            page.get_by_role("textbox", name="Username").fill("")
            page.get_by_role("textbox", name="Password").fill("")
            page.get_by_role("button", name="Login").click()

            page.wait_for_load_state("networkidle")

            grouped = df.groupby('Voucher No.')

            for voucher_no, items in grouped:
                print(f"Processing Voucher: {voucher_no} ({len(items)} items)")
                first_row = items.iloc[0]

                page.goto("/payments/sale")

                outlet_input = page.get_by_role("textbox", name="Select Outlet")
                outlet_input.wait_for()
                outlet_input.click()
                page.get_by_role("searchbox", name="Search").nth(2).fill(str(first_row['PARTY CODE']))
                page.get_by_role("option", name=str(first_row['PARTY CODE'])).click()
                page.wait_for_timeout(2000)

                for index, row in items.iterrows():
                    print("row-->", row)

                    try:
                        if not (page.get_by_role("button", name=row['CP/BP/REG']).is_visible()):
                            page.get_by_role("button", name=re.compile(r"CP|BP|Regional", re.IGNORECASE)).click()
                            page.locator("a").filter(has_text=row['CP/BP/REG']).click()
                            page.wait_for_timeout(500)

                        search_input = page.get_by_role("searchbox", name="Search:")
                        search_input.wait_for()
                        search_input.fill(str(row['A.code']))
                        page.wait_for_timeout(500)

                        sku_id = page.locator("#skunitstable").locator("tbody").locator("tr").first.locator("td").first.inner_text()
                        print('sku_id', sku_id)
                        price_input = page.locator("#price_" + sku_id)
                        price_input.wait_for()
                        price_input.fill(str(row['PER UNIT SALE PRICE']))
                        cases_input = page.locator("#cases_" + sku_id)
                        cases_input.wait_for()
                        cases_input.fill(str(row['Quantity (Case)']))

                        df.loc[index, 'Status'] = 'Processed'
                        print(f"Row {index} status updated to Processed")

                    except Exception as row_error:
                        df.loc[index, 'Status'] = f'Failed: {str(row_error)}'
                        print(f"Row {index} status updated to Failed: {row_error}")
                        continue

                page.get_by_role("link", name="Next").first.click()
                page.get_by_role("link", name="Next").click()

                for index in items.index:
                    if df.loc[index, 'Status'] == 'Processed':
                        df.loc[index, 'Status'] = 'Done'
                        df.loc[index, 'Processed Date'] = datetime.now().strftime("%Y-%m-%d")
                print(f"Voucher {voucher_no} rows updated to Done")

            page.close()
            print("Finished All rows!")


        except Exception as e:
            print(f"CRITICAL ERROR: {e}")
            context.tracing.stop(path="error_trace.zip")
            print("Trace saved to error_trace.zip. View it at trace.playwright.dev")
            traceback.print_exc()

        finally:
            browser.close()

            try:
                df.to_excel("MASTERSHEET_FOR_DAILY_DMS_ENTERIES_NEW.xlsx", index=False)
                print("Excel file saved with processing status updates")
            except PermissionError:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_filename = f"MASTERSHEET_FOR_DAILY_DMS_ENTERIES_NEW_BACKUP_{timestamp}.xlsx"
                df.to_excel(backup_filename, index=False)
                print(f"⚠️ Original file is locked/in use. Saved to: {backup_filename}")
                print("Please close the Excel file and manually replace the original with this backup.")

if __name__ == "__main__":
    run_automation()
    print("Script finished successfully.")
