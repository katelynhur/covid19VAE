import requests
from bs4 import BeautifulSoup
import os
import pandas as pd

# Base URL for query submission
QUERY_URL = "https://violinet.org/vaxafe/query.php"

# Output directories
DOWNLOAD_FOLDER = "vaxafe_male-vs-female_age/downloaded/"
DEBUG_FOLDER = "vaxafe_male-vs-female_age/debug/"

# Ensure output folders exist
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
os.makedirs(DEBUG_FOLDER, exist_ok=True)

# Vaccine names
vaccines = [
    "COVID19 (COVID19 (PFIZER-BIONTECH))", 
    "COVID19 (COVID19 (MODERNA))",
    "COVID19 (COVID19 (JANSSEN))",
    "COVID19 (COVID19 (PFIZER-BIONTECH BIVALENT))",
    "COVID19 (COVID19 (MODERNA BIVALENT))",  
    "COVID19 (COVID19 (NOVAVAX))"
]

# Age ranges
age_ranges = [(0, 9), (10, 19), (20, 29), (30, 39), (40, 49), 
              (50, 59), (60, 69), (70, 79), (80, 120)]

# Sex options with mapping to query values
sex_options = {"Male": "M", "Female": "F", "Any": ""}

# Iterate over all combinations
for vaccine in vaccines:
    for age_from, age_to in age_ranges:
        for sex_label, sex_value in sex_options.items():
            print(f"Submitting: {vaccine}, Age {age_from}-{age_to}, Sex {sex_label}")

            # POST parameters
            payload = {
                "vaccine[]": "",
                "vaccine_name[]": vaccine,
                "vax_manu[]": "",
                "state[]": "any",
                "age_from": age_from,
                "age_to": age_to,
                "sex": sex_value,
                "vax_date_from": "",
                "vax_date_to": "",
                "recvdate_from": "",
                "recvdate_to": "",
                "Submit2": "Query"
            }

            # Submit query
            session = requests.Session()
            response = session.post(QUERY_URL, data=payload)

            # Save the query response for debugging
            debug_filename = f"{vaccine}_{age_from}-{age_to}_{sex_label}.html".replace(" ", "_").replace("/", "-")
            debug_filepath = os.path.join(DEBUG_FOLDER, debug_filename)

            with open(debug_filepath, "w", encoding="utf-8") as debug_file:
                debug_file.write(response.text)

            # Parse the response for adverse event data
            soup = BeautifulSoup(response.text, "html.parser")
            rows = soup.find_all("tr")

            data = []
            first_row_skipped = False  # Flag to skip the first row if it's a header

            for row in rows:
                cells = row.find_all("td")
                if len(cells) >= 3:  # Ensure there are at least three columns
                    ae_number = cells[0].text.strip()  # Extract number
                    ae_name = cells[1].text.strip()
                    counts_text = cells[2].text.strip()

                    # Skip the first row if it's a header
                    if not first_row_skipped:
                        first_row_skipped = True
                        continue

                    # Extract AE count and total cases
                    if "/" in counts_text:
                        ae_count, total_cases = counts_text.split("/")
                        ae_count, total_cases = ae_count.strip(), total_cases.strip()

                        data.append([ae_number, ae_name, ae_count, total_cases])

            # Save extracted data to an Excel file
            if data:
                df = pd.DataFrame(data, columns=["No.", "Adverse Event", "AE Cases", "Total Cases"])
                excel_filename = f"{vaccine}_{age_from}-{age_to}_{sex_label}.xlsx".replace(" ", "_").replace("/", "-")
                excel_filepath = os.path.join(DOWNLOAD_FOLDER, excel_filename)
                df.to_excel(excel_filepath, index=False)

                print(f"Saved extracted data: {excel_filepath}")
            else:
                print(f"No AE data found for {vaccine}, {age_from}-{age_to}, {sex_label}")
