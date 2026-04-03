from playwright.sync_api import sync_playwright
import sys

def run():
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.set_content('<h1>Test PDF generation</h1>')
            page.pdf(path='/Users/pl-tq-261/repo/polteq-uren-sturen/test_playwright.pdf')
            browser.close()
            print("Successfully created test PDF")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run()
