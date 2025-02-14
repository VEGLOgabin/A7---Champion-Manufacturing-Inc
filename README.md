# Champion Manufacturing Web Scraper

## Overview
This project is a web scraper designed to extract product details from the Champion Manufacturing website. It utilizes **Playwright** for browser automation, **Pandas** for data handling, and **Rich** for enhanced terminal output.

## Features
- Searches for products using Manufacturer Number.
- Extracts product details including:
  - Product URL
  - Image URL
  - Description
  - Dimensions (Width, Height, Depth, Weight Capacity)
  - Standard Features
- Saves extracted data into an **Excel file**.

## Installation

Before running the scraper, ensure you have the required dependencies installed:

```bash
pip install asyncio pandas playwright rich openpyxl
```

You also need to install Playwright browsers:

```bash
playwright install
```

## Usage

Run the script with:

```bash
python scraper.py
```

### Arguments
- **excel_path**: Path to the input Excel file containing product data.
- **output_filename**: Path where the scraped data will be saved.
- **baseurl**: Base URL for Champion Manufacturing search.
- **found**: Counter for products found.
- **missing**: Counter for missing products.
- **headless**: Boolean flag for headless browser mode (default: `False`).

## Output
The scraper generates an Excel file containing the extracted data in the **Grainger** sheet. It includes columns for product details such as images, descriptions, dimensions, and standard features.

## Example Code

```python
scraper = ChampionManufacturingScraper(
    excel_path="Champion Manufacturing Content.xlsx",
    output_filename="output/Champion-manufacturing-output.xlsx",
    baseurl="https://championchair.com?s=",
    found=0,
    missing=0,
    headless=False
)
asyncio.run(scraper.run())
```

## Notes
- Ensure the **Excel file** contains a sheet named `Grainger` with a column for `mfr number`.
- The scraper runs asynchronously using **asyncio**.
- Modify the `baseurl` if the website structure changes.

## License
This project is licensed under the **MIT License**.
