import pandas as pd
from playwright.sync_api import sync_playwright
from string.templatelib import Template # py command
# from string import Template # python command
import traceback
import re
import numbers
from datetime import datetime
import os

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

options = ["CP", "BP", "Regional"]

def run_automation():

    if load_dotenv:
        load_dotenv()

    def normalize_identifier(value):
        if pd.isna(value):
            return ''
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, numbers.Integral):
            return str(int(value))
        if isinstance(value, numbers.Real):
            numeric_value = float(value)
            if numeric_value.is_integer():
                return str(int(numeric_value))
            return str(value).strip()
        return str(value).strip()

    dms_username = os.getenv("DMS_USERNAME", "").strip()
    dms_password = os.getenv("DMS_PASSWORD", "").strip()
    login_url = os.getenv("LOGIN_URL", "").strip()
    sales_url = os.getenv("SALES_URL", "").strip()
    if not dms_username or not dms_password:
        raise ValueError("Missing DMS credentials. Set DMS_USERNAME and DMS_PASSWORD environment variables.")

    with sync_playwright() as p:
        # Set headless=False for visual debugging, or use PWDEBUG=1 environment variable
        browser = p.chromium.launch(headless=True)  # Change to True for production
        context = browser.new_context()
        context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()
        df = pd.read_excel(
            "Automation_Sheet.xlsx",
            converters={
                'A.code': normalize_identifier,
                'CP/BP/REG': normalize_identifier,
                'PARTY CODE': normalize_identifier
            }
        )
        required_columns = ['A.code', 'CP/BP/REG', 'PARTY CODE']
        for col in required_columns:
            if col in df.columns:
                df[col] = df[col].apply(normalize_identifier)

        print(f"Excel loaded! Found {len(df)} rows.")

        if 'Status' not in df.columns:
            df['Status'] = 'Pending'

        if 'Processed Date' not in df.columns:
            df['Processed Date'] = ''

        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required column(s): {', '.join(missing_columns)}")

        def is_empty(value):
            return pd.isna(value) or str(value).strip() == ''

        skipped_count = 0
        for index, row in df.iterrows():
            missing_fields = [col for col in required_columns if is_empty(row[col])]
            if missing_fields:
                df.loc[index, 'Status'] = f"Skipped: Missing required field(s): {', '.join(missing_fields)}"
                skipped_count += 1

        rows_to_process = df[~df['Status'].astype(str).str.startswith('Skipped:')]
        print(f"Rows eligible for processing: {len(rows_to_process)} | Skipped rows: {skipped_count}")

        try:
            print("Starting Playwright...")
            page.goto(login_url)

            # page.fill('input[name="identifier"]', "")
            #
            # if page.is_visible('input[name="credentials.passcode"]'):
            #     page.fill('input[name="credentials.passcode"]', "")
            # else:
            #     page.click('input[type="submit"]')
            #     page.fill('input[name="credentials.passcode"]', "")
            #
            # page.click('input[type="submit"]')

            page.get_by_role("textbox", name="Username").fill(dms_username)
            page.get_by_role("textbox", name="Password").fill(dms_password)
            page.get_by_role("button", name="Login").click()

            page.wait_for_load_state("networkidle")

            grouped = rows_to_process.groupby('Voucher No.')

            for voucher_no, items in grouped:
                print(f"Processing Voucher: {voucher_no} ({len(items)} items)")
                first_row = items.iloc[0]

                page.goto(sales_url)

                outlet_input = page.get_by_role("textbox", name="Select Outlet")
                outlet_input.wait_for()
                outlet_input.click()
                page.get_by_role("searchbox", name="Search").nth(2).fill(first_row['PARTY CODE'])
                page.get_by_role("option", name=first_row['PARTY CODE']).click()
                # page.wait_for_timeout(2000)
                page.locator('#skunitstable_processing').wait_for(state='hidden', timeout=20000)

                for index, row in items.iterrows():
                    print("row-->", row)

                    try:
                        if not (page.get_by_role("button", name=row['CP/BP/REG']).is_visible()):
                            page.get_by_role("button", name=re.compile(r"CP|BP|Regional", re.IGNORECASE)).click()
                            page.locator("a").filter(has_text=row['CP/BP/REG']).click()
                            page.wait_for_timeout(500)
                            page.locator('#skunitstable_processing').wait_for(state='hidden', timeout=10000)

                        search_input = page.get_by_role("searchbox", name="Search:")
                        search_input.wait_for()
                        search_input.fill(row['A.code'])
                        page.wait_for_timeout(500)
                        page.locator('#skunitstable_processing').wait_for(state='hidden', timeout=10000)

                        sku_id = page.locator("#skunitstable").locator("tbody").locator("tr").first.locator("td").first.inner_text().strip()
                        
                        print('sku_id', sku_id)
                        if not sku_id.isdigit():
                            raise ValueError(f"Invalid SKU ID found: '{sku_id}' for row {index}")
                        
                        cases_available = page.locator("#casesavai_" + sku_id).inner_text()
                        print('cases_available', cases_available)
                        if int(row['Quantity (Case)']) > int(cases_available):
                            raise ValueError(f"Can't sell more than the inventory: '{sku_id}' for row {index}")

                        price_input = page.locator("#price_" + sku_id)
                        price_input.wait_for()
                        price_input.fill(str(row['PER UNIT SALE PRICE']))
                        cases_input = page.locator("#cases_" + sku_id)
                        cases_input.wait_for()
                        cases_input.fill(str(row['Quantity (Case)']))
                        page.wait_for_timeout(500)
                        min_price_alert = page.get_by_text("Minimum Per Unit Sale Price")

                        if min_price_alert.is_visible():
                            page.get_by_role("button", name="Ok").click()

                        df.loc[index, 'Status'] = 'Processed'
                        print(f"Row {index} status updated to Processed")

                    except Exception as row_error:
                        df.loc[index, 'Status'] = f'Failed: {str(row_error)}'
                        print(f"Row {index} status updated to Failed: {row_error}")
                        continue

                page.get_by_role("link", name="Next").first.click()

                if page.get_by_text("No quantity entered").is_visible():
                    for index in items.index:
                        df.loc[index, 'Status'] = 'Failed: No quantity entered dialog shown'
                        df.loc[index, 'Processed Date'] = datetime.now().strftime("%Y-%m-%d")
                    continue

                next_btn = page.get_by_role("link", name="Next")
                next_btn.wait_for()
                next_btn.click()

                finish_btn = page.get_by_role("link", name="Finish")
                finish_btn.wait_for()
                finish_btn.click()

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
            df.to_excel("PROCESSED_RECORDS.xlsx", index=False)
            print("Excel file saved with processing status updates")

if __name__ == "__main__":
    run_automation()
    print("Script finished successfully.")
