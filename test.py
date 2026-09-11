from playwright.sync_api import sync_playwright
from scrape_entities import get_authenticated_context

with sync_playwright() as p:
    browser, context, xsrf = get_authenticated_context(p)
    try:
        page = context.new_page()
        page.route("**/deskpro-messenger/**", lambda route: route.abort())

        page.goto("https://egpkenya.go.ke/public-app", wait_until="domcontentloaded", timeout=60000)
        page.wait_for_selector("table tbody tr", timeout=30000)

        toggle = page.locator("#first-toggle")
        if toggle.get_attribute("aria-expanded") != "true":
            toggle.click()
            page.wait_for_function(
                "document.querySelector('#first-toggle')?.getAttribute('aria-expanded') === 'true'",
                timeout=5000,
            )

        search_input = page.locator("input[formcontrolname='appNumber']")
        search_input.first.fill("AWWDA/704/APP/2026-27/2")
        page.wait_for_timeout(300)

        search_button = page.locator("button[type='submit']:has-text('Search')")
        print(f"Submit buttons found: {search_button.count()}")
        search_button.first.click()
        page.wait_for_timeout(3000)

        rows = page.locator("table tbody tr")
        row_count = rows.count()
        print(f"Rows after search: {row_count}")
        for i in range(row_count):
            print(f"  Row {i}: {rows.nth(i).inner_text().replace(chr(10), ' | ')}")

    finally:
        input("\nPress Enter to close browser...")
        browser.close()