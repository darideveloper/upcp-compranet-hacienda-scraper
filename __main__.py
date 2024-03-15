from libs.web_scraping import WebScraping


class Scraper(WebScraping):

    def __init__(self):
        """ Start chrome and load the home page"""
        
        self.home_page = "https://upcp-compranet.hacienda.gob.mx/sitiopublico/#/"
        
        super().__init__()
        self.set_page(self.home_page)
        
    def __wait_spinner__(self):
        """ Wait until page loads, checking the spinner """
        
        selector_spinner = '.spinner:not([style="display: none;"])'
        self.wait_die(selector_spinner)
        
    def apply_filters(self):
        """ Apply search filters to the page"""
        
        self.selectors = {
            "show_filters": '.p-button-label',
            "from_date": 'input[name="fechaDesdeP"]',
            "to_date": 'input[name="fechaHastaP"]',
            "name": 'input[name="nombreProcedimiento"]',
            "dependency_display": '[name="dependencias"] .p-multiselect-label-container',
            "dependency_search": '.p-multiselect-filter.p-inputtext',
            "dependency_checkbox": '[role="checkbox"]',
            "submit": 'button[type="submit"]',
        }
        
        # Wait until page loads
        self.__wait_spinner__()
        
        # Display all filters
        self.click_js(self.selectors["show_filters"])
        self.refresh_selenium()
        
        # Set dates
        self.send_data(self.selectors["from_date"], "01/01/2023")
        self.send_data(self.selectors["from_date"], "31/12/2023")
        
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
        self.refresh_selenium()
        self.__wait_spinner__()
        
        
if __name__ == "__main__":
    scraper = Scraper()
    scraper.apply_filters()