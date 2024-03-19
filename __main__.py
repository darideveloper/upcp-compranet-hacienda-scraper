import os
from dotenv import load_dotenv
from time import sleep
from datetime import datetime
from libs.web_scraping import WebScraping
from libs.xlsx import SpreadsheetManager

# Env variables
load_dotenv()
START_PAGE = int(os.getenv("START_PAGE"))


class Scraper(WebScraping):

    def __init__(self):
        """ Start chrome, load the home page and initialice excel file"""
        
        # Paths
        current_folder = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(current_folder, "data.xlsx")
        self.downloads_folder = os.path.join(current_folder, "downloads")
        os.makedirs(self.downloads_folder, exist_ok=True)
        
        # Start scraper
        self.home_page = "https://upcp-compranet.hacienda.gob.mx/sitiopublico/#/"
        
        super().__init__(width=1920, height=1080, download_folder=self.downloads_folder)
        self.set_page(self.home_page)
        
        # Start xlsx
        self.sheet_main_name = "main_table"
        self.sheet_details_name = "details_table"
        self.sheets = SpreadsheetManager(file_name=excel_path)
        
    def __wait_spinner__(self):
        """ Wait until page loads, checking the spinner """
        
        selector_spinner = '.spinner:not([style="display: none;"])'
        self.wait_die(selector_spinner, time_out=300)
        self.refresh_selenium()
        
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
        
        if self.get_elems(selectors["next_btn_disbaled"]):
            return False
        
        return True
    
    def __extract_table__(self, selectors: dict) -> list:
        """ Extract data from table
        
        Args:
            selector (dict): selectors for the table
            
        Returns:
            list: data extracted
        """
        
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
    
    def __extract_main_current_page__(self) -> list:
        """ Extract data from the current page in the main table
        
        Returns:
            list: data extracted from the main current page
        """
        
        selectors = {
            "row": '.p-datatable-unfrozen-view td:nth-child(1)',
            "id": 'tr:nth-child(index) > td:nth-child(2)',
            "caracter": '.p-datatable-unfrozen-view'
                        ' tr:nth-child(index) > td:nth-child(1)',
            "name": '.p-datatable-unfrozen-view tr:nth-child(index) > td:nth-child(2)',
            "entity": 'tr:nth-child(index) > td:nth-child(3)',
            "post_type": 'tr:nth-child(index) > td:nth-child(7)',
        }
        
        return self.__extract_table__(selectors)
        
    def __search_id__(self, id: str):
        """ Search a specific id in the main table
        
        Args:
            id (str): id to search
        """
        
        selectors = {
            "search_input": 'input[name="noProcedimiento"]',
            "submit": 'button[type="submit"]',
            "tab": '#p-tabpanel-2-label',
        }
        
        # Move to tab
        if self.get_elems(selectors["tab"]):
            self.click_js(selectors["tab"])
            self.__wait_spinner__()
            
        # Load home page
        self.set_page(self.home_page)
        self.__wait_spinner__()
        
        # Remove input old value
        script = f"""document.querySelector('{selectors["search_input"]}').value = ''"""
        self.driver.execute_script(script)
        self.refresh_selenium()
        
        # Search
        self.send_data(selectors["search_input"], id)
        self.refresh_selenium()
        self.click_js(selectors["submit"])
        self.__wait_spinner__()
    
    def __extract_contracts__(self) -> list:
        """ Extract contracts from details page
        
        Returns:
            list: matrix with contracts data
        """
        
        selectors = {
            "row": '[key="detalleDRC"] + br + [class="p-grid"]'
                   ' tr td:nth-child(1)',
            "num": '[key="detalleDRC"] + br + [class="p-grid"]'
                   ' tr:nth-child(index) td:nth-child(1)',
            "bidder": '[key="detalleDRC"] + br + [class="p-grid"]'
                      ' tr:nth-child(index) td:nth-child(2)',
            "date": '[key="detalleDRC"] + br + [class="p-grid"]'
                    ' tr:nth-child(index) td:nth-child(6)',
            "taxes": '[key="detalleDRC"] + br + [class="p-grid"]'
                     ' tr:nth-child(index) td:nth-child(8)',
        }
        
        data = self.__extract_table__(selectors)
        return data
    
    def __extract_requirements__(self) -> str:
        """ Extract requirements from details page
        
        Returns:
            str: matrix with requirements data
        """
        
        selectors = {
            "row": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                   ' tr td:nth-child(1)',
            "num": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                   ' tr:nth-child(index) td:nth-child(1)',
            "quantity": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                        ' tr:nth-child(index) td:nth-child(7)',
            "part": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                    ' tr:nth-child(index) td:nth-child(2)',
            "key": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                   ' tr:nth-child(index) td:nth-child(3)',
            "description": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                           ' tr:nth-child(index) td:nth-child(4)',
            "details": '[class="p-fluid p-formgrid p-grid"] > div:last-child'
                       ' tr:nth-child(index) td:nth-child(5)',
        }
        
        data = self.__extract_table__(selectors)
        return data
    
    def __download_files_page__(self, id: str):
        """ Download files from current page
        
        Args:
            id (str): id to search
        """
        
        selectors = {
            "row": '[key="anexos"] + br + .p-grid tr td:nth-child(1)',
            "num": '[key="anexos"] + br + .p-grid tr:nth-child(index) td:nth-child(1)',
            "type": '[key="anexos"] + br + .p-grid tr:nth-child(index) td:nth-child(2)',
            "download_btn": '[key="anexos"] + br + .p-grid'
                            ' tr:nth-child(index) td.oculto-impresion i',
            "next_btn": '[key="anexos"] + br + .p-grid'
                        ' .p-paginator-next:not(.p-disabled)'
        }
        
        # Old files in download folder
        old_files = os.listdir(self.downloads_folder)
        old_files = list(filter(lambda file: file.endswith(".pdf"), old_files))
        
        # Loop rows
        rows_num = len(self.get_elems(selectors["row"]))
        for row_index in range(1, rows_num + 1):
            
            # Get num anbd document type
            num = self.get_text(selectors["num"].replace("index", str(row_index)))
            type = self.get_text(selectors["type"].replace("index", str(row_index)))
            
            # Create id folder
            id_folder = os.path.join(self.downloads_folder, id)
            os.makedirs(id_folder, exist_ok=True)
            
            # New file path
            file_ext = "pdf"
            file_name = f"{id} - {num} - {type}.{file_ext}"
            moved_file_path = os.path.join(id_folder, file_name)
            
            # Try to downbload file 3 times
            downloaded = False
            for _ in range(3):
            
                # Download file and wait to finish
                self.click(selectors["download_btn"].replace("index", str(row_index)))
                self.__wait_spinner__()
                sleep(15)

                # Detect new file
                new_files = os.listdir(self.downloads_folder)
                new_files = list(filter(lambda file: file.endswith(".pdf"), new_files))
                new_files = list(set(new_files) - set(old_files))
                if not new_files:
                    print(f"\t\tFile {num} - {type} not downloaded. Retrying...")
                    continue
                
                downloaded = True
                break
            
            if not downloaded:
                print(f"\t\tFile {num} - {type} not downloaded")
                continue
            
            new_file_path = os.path.join(self.downloads_folder, new_files[0])
            
            # Move file to id folder
            os.rename(new_file_path, moved_file_path)
            
            print(f"\t\tFile {num} - {type} downloaded")
        
        # Validate and go to next page
        if self.get_elems(selectors["next_btn"]):
            self.click_js(selectors["next_btn"])
            self.__wait_spinner__()
            return True
    
        return False
        
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
        
        self.click_js(self.selectors["tab"])
        self.__wait_spinner__()
    
    def extract_main_table(self):
        """ Get general data from main table"""
        
        print("Extracting main table...")
        
        # Set main table in excel
        self.sheets.create_set_sheet(self.sheet_main_name)
        
        # Move to start page
        for _ in range(START_PAGE - 1):
            self.__go_next_page_main_table__()
        
        current_row = 3 + (START_PAGE - 1) * 100
        page = START_PAGE
        while True:
            
            print(f"\tExtracting page {page} from main table...")
            
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
            
    def extract_details(self):
        """ Extract details from each id in the excel """
        
        print("Extracting details tables...")
        
        selectors = {
            "id": 'tr:nth-child(1) > td:nth-child(2)',
            "dependency": 'app-sitiopublico-detalle-datos-ente-pc'
                          ' > div label:nth-child(3)',
            "branch": 'app-sitiopublico-detalle-datos-ente-pc'
                      ' > div div:nth-child(2) label:nth-child(3)',
            "unity": 'app-sitiopublico-detalle-datos-ente-pc'
                     ' > div div:nth-child(3) label:nth-child(3)',
            "in_charge": 'app-sitiopublico-detalle-datos-ente-pc'
                     ' > div div:nth-child(4) label:nth-child(3)',
            "email": 'app-sitiopublico-detalle-datos-ente-pc'
                     ' > div div:nth-child(5) label:nth-child(3)',
            "entity": 'app-sitiopublico-detalle-datos-general-pc'
                      ' div:nth-child(4) label:nth-child(3)',
        }
        
        # Read main table
        self.sheets.create_set_sheet(self.sheet_main_name)
        main_data = self.sheets.get_data()[2:]
        
        # Remove duplicated data
        main_data = list(set(map(tuple, main_data)))
                
        # Read data from excel
        self.sheets.create_set_sheet(self.sheet_details_name)
        
        # Detect last rows saved
        last_index_main = 0
        last_index_details = 0
        details_data = self.sheets.get_data()
        details_data = list(filter(lambda row: row[0], details_data))
        
        if details_data:
            last_id = details_data[-1][0]
            
            for index in range(len(main_data)):
                if main_data[index][0] == last_id:
                    last_index_main = index
                    break
                
            for index in range(len(details_data)):
                if details_data[index][0] == last_id:
                    last_index_details = index
                    break
        
        rows_saved = 3 + last_index_details
        max_row = len(main_data)
        for row in main_data[last_index_main:]:
                        
            # Search id
            id = row[0]
            index_row = main_data.index(row) + 1
            print(f"\tExtracting details from {id} ({index_row}/{max_row})...")
            self.__search_id__(id)
            self.__wait_spinner__()
            
            # Open details
            self.click_js(selectors["id"])
            self.__wait_spinner__()
            sleep(8)
            self.refresh_selenium()
            
            # Extract general data
            general_data = []
            selectors_copy = selectors.copy()
            del selectors_copy["id"]
            for _, selector_value in selectors_copy.items():
                value = self.get_text(selector_value)
                general_data.append(value)
            
            # Extract internal tables
            contracts = self.__extract_contracts__()
            
            # Extract internal tables
            requirements = self.__extract_requirements__()

            # Merge contracts and requirements
            data = []
            len_contracts = len(contracts)
            len_requirements = len(requirements)
            for index in range(max(len_contracts, len_requirements)):
                
                # Add data or empty cell
                if index < len_contracts:
                    contract = contracts[index]
                else:
                    contract = [" "] * 4
                    
                if index < len_requirements:
                    requirement = requirements[index]
                else:
                    requirement = [" "] * 6
                    
                data.append(list(row) + general_data + contract + requirement)
                
            # Write data in excel
            self.sheets.write_data(data, rows_saved)
            self.sheets.save()
            rows_saved += len(data)
            
            sleep(30)

    def download_files(self):
        """ Download attached files from each id in the excel """
        
        selectors = {
            "id": 'tr:nth-child(1) > td:nth-child(2)',
        }
        
        # Read main table
        self.sheets.create_set_sheet(self.sheet_main_name)
        sheets_data = self.sheets.get_data()[2:]
        
        # Remove duplicated data
        sheets_data = list(set(map(tuple, sheets_data)))
        
        # Read data from excel
        self.sheets.create_set_sheet(self.sheet_details_name)
        
        max_row = len(sheets_data)
        for row in sheets_data:
            
            id = row[0]
            
            # Skip if download folder already exists
            id_folder = os.path.join(self.downloads_folder, id)
            if os.path.exists(id_folder):
                print(f"\tFiles already downloaded for {id}. Skipping...")
                continue
                        
            # Search id
            index_row = sheets_data.index(row) + 1
            print(f"\tDownloading files from {id} ({index_row}/{max_row})...")
            self.__search_id__(id)
            self.__wait_spinner__()
            
            # Open details
            self.click_js(selectors["id"])
            self.__wait_spinner__()
            sleep(8)
            self.refresh_selenium()
            
            # Download files
            while True:
                more_pages = self.__download_files_page__(id)
                if not more_pages:
                    break
                
            
if __name__ == "__main__":
    
    # Main menu
    print("1. Extract main data\n2. Extract details\n3. Download files")
    option = input("Select an option: ").lower().strip()
    
    # Start scraper
    scraper = Scraper()
    
    if option == "1":
        # Main table
        scraper.apply_filters()
        scraper.extract_main_table()
    elif option == "2":
        # details tables
        scraper.extract_details()
    elif option == "3":
        # download files
        scraper.download_files()
    else:
        print("Invalid option")
       