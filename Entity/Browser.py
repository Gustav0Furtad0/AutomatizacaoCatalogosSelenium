from selenium import webdriver

class Browser(webdriver.Chrome):
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    def __init__(self, *args, **kwargs):
        super().__init__(options=self.options, *args, **kwargs)
    
    def wait_response(self, message):
        input(message)