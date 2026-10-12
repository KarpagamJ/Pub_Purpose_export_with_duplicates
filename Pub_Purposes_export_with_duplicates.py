import requests
import os
import pandas as pd

def fetch_all_purposes():
    base_url = "https://host_name.onetrust.com/api/consentmanager/v1/purposes"
    token = "paste_your_token_here"
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {token}"
    }

    all_content = []
    page = 0
    page_size = 20

    while True:
        url = f"{base_url}?page={page}&size={page_size}"
        print(f" Fetching page {page + 1}...")

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()

            content = data.get('content')
            total_pages = data.get('totalPages') or data.get('total_pages')

            if not content:
                print(f" Empty or missing 'content' on page {page + 1}")
                break

            all_content.extend(content)

            if total_pages is not None and page + 1 >= total_pages:
                break

            page += 1

        except requests.exceptions.RequestException as e:
            print(f" Failed to fetch page {page + 1}: {e}")
            break

    return all_content

def normalize_label(label: str) -> str:
    if not isinstance(label, str):
        return ""
    return " ".join(label.strip().split())

if __name__ == "__main__":
    purposes = fetch_all_purposes()

    if purposes:
        df = pd.DataFrame(purposes)

        print(" Fetched", len(df), "total purposes")
        print(" Available columns:", df.columns.tolist())

        # Try to find the label column (capitalized and lowercase)
        label_fields = ['Label', 'label', 'name', 'title', 'displayName', 'purposeName']
        label_field = None
        for field in label_fields:
            if field in df.columns:
                label_field = field
                break

        if not label_field:
            print(" Could not find a label field in data.")
            exit()

        # Normalize label for duplicate detection
        df['normalized_label'] = df[label_field].apply(normalize_label)

        # Detect label-based duplicates (different spacing, same label intent)
        duplicate_labels = df[df.duplicated(subset='normalized_label', keep=False)].copy()

        # Output file name
        output_file = "onetrust_purposes_with_duplicates.xlsx"

        # Save to Excel with two sheets
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='All Purposes', index=False)
                duplicate_labels.to_excel(writer, sheet_name='Label Duplicates', index=False)

            print(f"\nExcel file saved: {os.path.abspath(output_file)}")
            print(f" Total records: {len(df)}")
            print(f" Potential label-based duplicates: {len(duplicate_labels)}")

        except Exception as e:
            print(f" Failed to save Excel file: {e}")
    else:
        print(" No purposes fetched from the API.")