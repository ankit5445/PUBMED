import requests
from bs4 import BeautifulSoup
import pandas as pd

# Load PMCID list
references_df = pd.read_excel("PMCREF.xlsx")
pmcid_list = references_df["PMCID"].dropna().unique().tolist()

all_sections = []
i = 0

for pmcid in pmcid_list:  # You can increase or remove limit as needed
    url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    i += 1
    print(f"{i}: Fetching {pmcid}")

    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Failed to fetch {pmcid}")
        continue

    soup = BeautifulSoup(response.content, "html.parser")

    # Extract article title
    title = soup.find("h1").text.strip() if soup.find("h1") else ""

    # --- Extract metadata after <hgroup> ---
    meta_div = soup.find("div", class_="ameta")
    if meta_div:
        hgroup = meta_div.find("hgroup")
        if hgroup:
            next_element = hgroup.find_next_sibling()
            while next_element:
                section_text = next_element.get_text(separator=" ", strip=True)
                tag_name = next_element.name if next_element.name else "Unknown"
                if section_text:
                    all_sections.append({
                        "PMCID": pmcid,
                        "Title": title,
                        "Section Title": f"Meta - {tag_name}",
                        "Content": section_text
                    })
                next_element = next_element.find_next_sibling()

    # --- Main article body sections ---
    body = soup.find("section", {"class": "main-article-body"})
    if body:
        for section in body.find_all(["section", "h2", "h3"], recursive=True):
            header = section.find(["h2", "h3"])
            if header:
                section_title = header.text.strip()
                section_text = section.get_text(separator=" ", strip=True)
                all_sections.append({
                    "PMCID": pmcid,
                    "Title": title,
                    "Section Title": section_title,
                    "Content": section_text
                })

# Create final DataFrame
sections_df = pd.DataFrame(all_sections)

# Optional: Save to Excel or CSV
# sections_df.to_excel("PMC_Metadata_Sections.xlsx", index=False)

import math
num_rows = len(sections_df)
split_size = math.ceil(num_rows / 4)

sections_df_split = [sections_df[i*split_size:(i+1)*split_size] for i in range(4)]

for idx, split_df in enumerate(sections_df_split):
    split_df.to_excel(f"split_part_{idx+1}.xlsx", index=False)
    

print("Split into 4 Excel files successfully.")