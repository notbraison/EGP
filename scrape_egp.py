import uuid
import pandas as pd
import requests


def scrape_egp_data(page_size=100):
    url = "https://egpkenya.go.ke/api/app/public-app-detail"
    session = requests.Session()

    # 1. Establish session cookies by visiting the main public app page
    base_headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    }

    print("Connecting to e-GP portal to initialize session...")
    session.get("https://egpkenya.go.ke/public-app", headers=base_headers)

    # Extract XSRF token set by server cookies
    xsrf_token = session.cookies.get("XSRF-TOKEN", "")

    api_headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://egpkenya.go.ke",
        "Referer": "https://egpkenya.go.ke/public-app",
        "X-XSRF-TOKEN": xsrf_token,
    }

    all_records = []
    page = 1

    while True:
        # Generate fresh idempotency key per request
        api_headers["Idempotency-Key"] = str(uuid.uuid4())

        payload = {
            "appNumber": "",
            "finYear": 0,
            "procuringEntity": None,
            "pageSize": page_size,
            "page": page,
        }

        print(f"Fetching page {page} with pageSize={page_size}...")
        response = session.post(url, json=payload, headers=api_headers)

        if response.status_code != 200:
            print(
                f"Request failed on page {page} with status code"
                f" {response.status_code}"
            )
            break

        data = response.json()

        # Extract record list depending on standard API keys
        items = (
            data.get("data")
            or data.get("content")
            or data.get("items")
            or data.get("result")
            or []
        )

        if not items:
            print("No records found or end of data reached.")
            break

        all_records.extend(items)
        print(f"Retrieved {len(items)} items (Total so far: {len(all_records)})")

        # Stop if we received fewer records than requested page size
        if len(items) < page_size:
            break

        page += 1

    if not all_records:
        print("No data extracted.")
        return None

    df = pd.DataFrame(all_records)
    return df


if __name__ == "__main__":
    df = scrape_egp_data(page_size=500)

    if df is not None and not df.empty:
        print("\nExtracted Columns:", df.columns.tolist())
        print("\nFirst 5 Rows:")
        print(df.head())

        # Output raw scraped data
        output_file = "scraped_sheet2_raw.xlsx"
        df.to_excel(output_file, index=False)
        print(f"\nSaved scraped records to '{output_file}' successfully!")