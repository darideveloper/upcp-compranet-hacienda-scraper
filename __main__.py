import os
from dotenv import load_dotenv
from time import sleep
from datetime import datetime
from libs.web_scraping import WebScraping
from libs.xlsx import SpreadsheetManager

# Env variables
START_PAGE = int(os.getenv("START_PAGE"))


class Scraper(WebScraping):

    def __init__(self):
        """ Start chrome, load the home page and initialice excel file"""
        
        # Start scraper
        self.home_page = "https://upcp-compranet.hacienda.gob.mx/sitiopublico/#/"
        
        super().__init__(width=1920, height=1080)
        self.set_page(self.home_page)
        
        # Start xlsx
        current_folder = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(current_folder, "data.xlsx")
        self.sheets = SpreadsheetManager(file_name=excel_path)
        
    def __wait_spinner__(self):
        """ Wait until page loads, checking the spinner """
        
        selector_spinner = '.spinner:not([style="display: none;"])'
        self.wait_die(selector_spinner, time_out=120)
        
    def __set_date__(self, month: int, year: int, selector_calendar: str,
                     selector_back: str, selector_day: str):
        """ Set date in the calendar

        Args:
            month (int): month to set
            year (int): year to set
            selector_calendar (str): selector for the calendar button
            selector_back (str): selector for the back button
            selector_day (str): selector for the day button
        """
        
        # Open calendar
        self.click_js(selector_calendar)
        self.refresh_selenium()
        
        # Current date
        current_month = datetime.now().month
        current_year = datetime.now().year
        
        # Set date
        years = current_year - year
        months = 12 * years + (current_month - month)
        self.refresh_selenium()
        for _ in range(months):
            self.click_js(selector_back)
            sleep(0.3)
            
        # Select day
        self.click_js(selector_day)
        
    def apply_filters(self):
        """ Apply search filters to the page"""
        
        self.selectors = {
            "show_filters": '.p-button-label',
            "date_back": '.p-datepicker-header button',
            "date_from_calendar": '[name="fechaDesdeP"] button',
            "date_to_calendar": '[name="fechaHastaP"] button',
            "date_from_day": '.p-datepicker-calendar tr:first-child > td span',
            "date_to_day": '.p-datepicker-calendar tr:last-child > td span',
            "name": 'input[name="nombreProcedimiento"]',
            "dependency_display": '[name="dependencias"] .p-multiselect-label-container',
            "dependency_search": '.p-multiselect-filter.p-inputtext',
            "dependency_checkbox": '.p-multiselect-item',
            "submit": 'button[type="submit"]',
            "tab": '#p-tabpanel-2-label',
        }
        
        # Wait until page loads
        self.__wait_spinner__()
        
        # Display all filters
        self.click_js(self.selectors["show_filters"])
        self.refresh_selenium()
        
        # Set start date
        self.__set_date__(
            month=1,
            year=2023,
            selector_calendar=self.selectors["date_from_calendar"],
            selector_back=self.selectors["date_back"],
            selector_day=self.selectors["date_from_day"]
        )
        
        # Set end date
        self.__set_date__(
            month=12,
            year=2023,
            selector_calendar=self.selectors["date_to_calendar"],
            selector_back=self.selectors["date_back"],
            selector_day=self.selectors["date_to_day"]
        )
        
        # Set search name
        self.send_data(self.selectors["name"], "MEDICAMENTO")
        
        # Set dependency
        self.click_js(self.selectors["dependency_display"])
        self.refresh_selenium()
        self.send_data(self.selectors["dependency_search"], "IMSS")
        self.refresh_selenium()
        self.click_js(self.selectors["dependency_checkbox"])
        
        # Search
        self.click_js(self.selectors["submit"])
        self.__wait_spinner__()
        self.refresh_selenium()
        
        self.click_js(self.selectors["tab"])
        self.__wait_spinner__()
        self.refresh_selenium()
    
    def __go_next_page_main_table__(self) -> bool:
        """ Go to the next page in the main table
        
        Returns:
            bool: True if there is a next page, False otherwise
        """
        
        selectors = {
            "next_btn": '.p-paginator-next',
            "next_btn_disbaled": '.p-paginator-next.p-disabled',
        }
        
        self.click_js(selectors["next_btn"])
        self.__wait_spinner__()
        self.refresh_selenium()
        
        if self.get_elems(selectors["next_btn_disbaled"]):
            return False
        
        return True
    
    def __extract_main_current_page__(self):
        
        selectors = {
            "row": '.p-datatable-unfrozen-view td:nth-child(1)',
            "id": 'tr:nth-child(index) > td:nth-child(2)',
            "caracter": '.p-datatable-unfrozen-view'
                        ' tr:nth-child(index) > td:nth-child(1)',
            "name": '.p-datatable-unfrozen-view tr:nth-child(index) > td:nth-child(2)',
            "entity": 'tr:nth-child(index) > td:nth-child(3)',
            "post_type": 'tr:nth-child(index) > td:nth-child(7)',
        }
        
        # Get rows num
        rows_num = len(self.get_elems(selectors["row"]))
        
        # Extract each required value for each row
        data = []
        for row_index in range(1, rows_num + 1):
            row_data = []
            for selector_name, selector_value in selectors.items():
                
                # Skip no required selectors
                if selector_name in ["row"]:
                    continue
                
                # Extract text
                selector = selector_value.replace("index", f"{row_index}")
                value = self.get_text(selector)
                row_data.append(value)
                    
            data.append(row_data)
            
        return data
    
    def extract_main_table(self):
        """ Get general data from main table"""
        
        # Set main table in excel
        self.sheets.create_set_sheet("main_table")
        
        # Move to start page
        for _ in range(START_PAGE - 1):
            self.__go_next_page_main_table__()
        
        current_row = 3 + (START_PAGE - 1) * 100
        page = START_PAGE
        while True:
            
            print(f"Extracting page {page} from main table...")
            
            # Extract data
            data = self.__extract_main_current_page__()
            
            # Save data in excel
            self.sheets.write_data(data, current_row)
            self.sheets.save()
            current_row += len(data)
            page += 1
            
            # Move to next page
            more_pages = self.__go_next_page_main_table__()
            if not more_pages:
                break
        
        
if __name__ == "__main__":
    scraper = Scraper()
    scraper.apply_filters()
    scraper.extract_main_table()