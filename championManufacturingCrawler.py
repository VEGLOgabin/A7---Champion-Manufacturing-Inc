import asyncio
import pandas as pd
from playwright.async_api import async_playwright, expect
from rich import print
import re
import os

class ChampionManufacturingScraper:
    """Web scraper for extracting product details from the Champion Manufacturing."""
    def __init__(self, excel_path: str, output_filename: str, baseurl : str, found : int, missing : int, headless: bool = False):
        self.filepath = excel_path
        self.output_filename = output_filename
        self.baseurl = baseurl
        self.headless = headless
        self.found = found
        self.missing = missing
        self.df = pd.read_excel(self.filepath, sheet_name="Grainger")

    async def launch_browser(self):
        """Initialize Playwright and open the browser."""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=self.headless)
        self.context = await self.browser.new_context()
        self.page = await self.context.new_page()

    async def close_browser(self):
        """Close the browser and Playwright instance."""
        await self.browser.close()
        await self.playwright.stop()

    async def search_product(self, mfr_number: str):
        """Search for a product by mfr number and return its first result URL."""
        try:
            await self.page.goto(self.baseurl + str(mfr_number))
            # Check if the "Nothing Found" message is displayed
            nothing_found_locator = self.page.locator('h1.pt-4.pb-3')
            if await nothing_found_locator.is_visible():
                nothing_found_text = await nothing_found_locator.text_content()
                if "Nothing Found" in nothing_found_text:
                    # print(f"[red]No products found for {mfr_number}[/red]")

                    return None

            # Get the first product URL
            search_result = await self.page.locator('//a[@class="post-wrapper--block col-lg-4 mb-4 position-relative"]').all()
            
            if search_result:
                url = await search_result[0].get_attribute('href')
                if "/product/" in url:
                    return url
                else:
                    # print(f"[red]No products found for {mfr_number}[/red]")
                    return None
        except Exception as e:
            print(f"Error occurred: {e}")
            pass
        return None

    async def scrape_product_details(self, url: str):
        """Extract product details from the given URL."""
        print(f"[cyan]Scraping data from:[/cyan] {url}")
        new_page = await self.context.new_page()
        await new_page.goto(url)

        data = {
            "url": url,
            "image": "",
            "description": "",
            "dimensions": {},
            "standard features": []
        }

        # Extract Product Image (jpg)
        try:
            image_locator = new_page.locator(
                '//div[contains(@class, "carousel-inner")]//div[contains(@class, "carousel-item active")]//figure/img'
            )
            await expect(image_locator).to_be_visible(timeout=5000)
            data["image"] = await image_locator.get_attribute("src")
        except Exception as e:
            print(f"Error extracting image: {e}")

        # Extract Product Description
        try:
            description_locator = new_page.locator('p.mt-2.pt-4.pb-4.pb-md-4.desktop')

            await expect(description_locator).to_be_visible(timeout=5000)
            description_raw = await description_locator.text_content()    #.replace('\n',"").replace('\t',"").replace('\xa0',"")
            data["description"] = re.sub(r'\s+', ' ', description_raw).strip()
        except Exception as e:
            print(f"Error extracting description: {e}")

        # Extract Measurements and Dimensions
        try:
            # Locate each row in the dimensions table
            dimension_rows = await new_page.locator('//div[@id="collapsemanualOne"]//table//tr').all()
            for row in dimension_rows:
                cells = await row.locator('td').all_text_contents()
                if len(cells) >= 2:
                    label = cells[0].strip()
                    value = cells[1].strip()

                    if value:  # Check if value is not empty
                        # Handle ranges (e.g., "29″ – 34″")
                        value = value.split("–")[0].strip()

                        if "Overall Width" in label:
                            data["dimensions"]["width"] = value.split('"')[0].strip().split("(")[0].strip()
                        if "Overall Height" in label or "Seat Height Range " in label or "Standard Height" in label:
                            data["dimensions"]["height"] = value.split('"')[0].strip().split("(")[0].strip()
                        if "Weight Capacity" in label:
                            data["dimensions"]["weight"] = value.split('lbs')[0].strip().split("(")[0].strip()
                        if "Overall Depth" in label:
                            data["dimensions"]["depth"] = value.split('"')[0].strip().split("(")[0].strip()
        except Exception as e:
            print(f"Error extracting dimensions: {e}")

        # Extract Standard Features
        try:
            features_elements = await new_page.locator('//div[@class="card-body"]//ul[@class="features-list"]/li').all()
            features_list = []
            for feature in features_elements:
                text = await feature.text_content()
                features_list.append(text.strip())
            data["standard features"] = features_list
        except Exception as e:
            print(f"Error extracting standard features: {e}")
        await new_page.close()
        return data


    async def run(self):
        """Main function to scrape product details and save them to an Excel file."""

        await self.launch_browser()
        await self.page.goto(self.baseurl)

        for index, row in self.df.iterrows():
            mfr_number = row["mfr number"]
            model_name = row['model name']
            url = await self.search_product(str(mfr_number))
            if not url:
                self.missing += 1
            else:
                self.found += 1
            if url:
                product_data = await self.scrape_product_details(url)
                if product_data:
                    print(f"[green]{model_name} | {mfr_number} [/green] - Data extracted successfully.")
                    self.df.at[index, "Product URL"] = product_data.get("url", "")
                    self.df.at[index, "Product Image (jpg)"] = product_data.get("image", "")
                    self.df.at[index, "Product Image"] = product_data.get("image", "")
                    self.df.at[index, "product description"] = product_data.get("description", "")
                    self.df.at[index, "depth"] = product_data["dimensions"].get("depth", "").split('″')[0].strip().split("(")[0].strip()
                    self.df.at[index, "height"] = product_data["dimensions"].get("height", "").split('″')[0].strip().split("(")[0].strip()
                    self.df.at[index, "width"] = product_data["dimensions"].get("width", "").split('″')[0].strip().split("(")[0].strip()
                    self.df.at[index, "weight"] = product_data["dimensions"].get("weight", "").split('″')[0].strip().split("(")[0].strip()
                    self.df.at[index, "green certification? (Y/N)"] = "N"
                    self.df.at[index, "emergency_power Required (Y/N)"] = "N"
                    self.df.at[index, "dedicated_circuit Required (Y/N)"] = "N"
                    self.df.at[index, "water_cold Required (Y/N)"] = "N"
                    self.df.at[index, "water_hot  Required (Y/N)"] = "N"
                    self.df.at[index, "drain Required (Y/N)"] = "N"
                    self.df.at[index, "water_treated (Y/N)"] = "N"
                    self.df.at[index, "steam  Required(Y/N)"] = "N"
                    self.df.at[index, "vent  Required (Y/N)"] = "N"
                    self.df.at[index, "vacuum Required (Y/N)"] = "N"
                    self.df.at[index, "ada compliant (Y/N)"] = "N"
                    self.df.at[index, "antimicrobial coating (Y/N)"] = "N"
            else:
                print(f"[red]{model_name} | {mfr_number} [/red] - Not found")
        print(f"[red]Missing : {self.missing} [/red]")
        print(f"[green]Found : {self.found} [/green]")
        self.df.to_excel(self.output_filename, index=False, sheet_name="Grainger")
        await self.close_browser()


if __name__ == "__main__":
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    scraper = ChampionManufacturingScraper(
        excel_path="Champion Manufacturing Content.xlsx",
        output_filename="output/Champion-manufacturing-output.xlsx",
        baseurl = "https://championchair.com?s=",
        found = 0 ,
        missing = 0,
        headless=False
    )
    asyncio.run(scraper.run())
