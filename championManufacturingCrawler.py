import asyncio
import pandas as pd
from playwright.async_api import async_playwright, expect
from rich import print


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
                    print(f"[red]No products found for {mfr_number}[/red]")

                    return None

            # Get the first product URL
            search_result = await self.page.locator('//a[@class="post-wrapper--block col-lg-4 mb-4 position-relative"]').all()
            
            if search_result:
                url = await search_result[0].get_attribute('href')
                if "/product/" in url:
                    return url
                else:
                    print(f"[red]No products found for {mfr_number}[/red]")
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

        data = {"url": url, "image": "", "price": "", "description": "", "dimensions": {}, "green certification": "N"}

        # Extract Image
        try:
            image_locator = new_page.locator('//div[@id="swiper-pdp-gallery"]/div/div/a/img')
            await expect(image_locator).to_be_visible(timeout=5000)
            data["image"] = await image_locator.get_attribute("src")
        except:
            pass

        # Extract Price
        try:
            price_locator = new_page.locator('//span[@class="text-red"]/span[@class="pdp-price__price"]')
            await expect(price_locator).to_be_visible(timeout=5000)
            data["price"] = await price_locator.text_content()
        except:
            pass

        # Extract Description
        try:
            description_locator = new_page.locator('//div[@class="pdp-main__short-desc"]/p')
            await expect(description_locator).to_be_visible(timeout=5000)
            data["description"] = await description_locator.text_content()
        except:
            pass

        # Extract Dimensions
        try:
            dimensions = await new_page.locator('//details[@id="dimensions"]//div[@class="metafield-rich_text_field"]/ul/li').all()
            for dimension in dimensions:
                info = (await dimension.text_content()).strip()
                if ":" in info:
                    label, value = info.split(":", 1)
                    data["dimensions"][label.strip()] = value.strip()
        except:
            pass

        # Extract Green Certification
        try:
            certifications = await new_page.locator('//div[@class="pdp-icon flex center-vertically"]').all()
            data["green certification"] = "Y" if any("GREEN" in await cert.text_content() for cert in certifications) else "N"
        except:
            pass

        await new_page.close()
        return data

    async def run(self):
        """Main function to scrape product details and save them to an Excel file."""

                
        print("[blue]Sheet Headers:[/blue]")
        for header in self.df.columns:
            print(header)

        await self.launch_browser()
        await self.page.goto(self.baseurl)

        # for index, row in self.df.iterrows():
        #     mfr_number = row["mfr number"]
        #     print(mfr_number)
        #     url = await self.search_product(str(mfr_number))
        #     if not url:
        #         self.missing += 1
        #     else:
        #         self.found += 1
        #     print(url)
        # print("Found : ", self.found)
        # print("Missing : ", self.missing)

            # if url:
            #     product_data = await self.scrape_product_details(url)

        #         if product_data:
        #             print(f"[green]{mfr_number}[/green] - Data extracted successfully.")
        #             self.df.at[index, "Product URL"] = product_data.get("url", "")
        #             self.df.at[index, "unit cost"] = product_data.get("price", "")
        #             self.df.at[index, "Product Image (jpg)"] = product_data.get("image", "")
        #             self.df.at[index, "Product Image"] = product_data.get("image", "")
        #             self.df.at[index, "product description"] = product_data.get("description", "")
        #             self.df.at[index, "green certification? (Y/N)"] = product_data.get("green certification", "")
        #             self.df.at[index, "depth"] = product_data["dimensions"].get("Overall Depth", "")
        #             self.df.at[index, "height"] = product_data["dimensions"].get("Overall Height", "")
        #             self.df.at[index, "width"] = product_data["dimensions"].get("Overall Width", "")
        #     else:
        #         print(f"[red]{mfr_number}[/red] - Not found")

        # self.df.to_excel(self.output_filename, index=False, sheet_name="Haworth")
        # await self.close_browser()


if __name__ == "__main__":
    scraper = ChampionManufacturingScraper(
        excel_path="Champion Manufacturing Content.xlsx",
        output_filename="Champion-manufacturing-output.xlsx",
        baseurl = "https://championchair.com?s=",
        found = 0 ,
        missing = 0,
        headless=False
    )
    asyncio.run(scraper.run())
