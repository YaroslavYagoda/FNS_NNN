from chromebrowser import ChromeBrowser
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager


class FireFox(ChromeBrowser):
    """
    Браузер FireFox - дочерний класс "ChromeBrowser"
    """

    def __init__(self):
        options = webdriver.FirefoxOptions()
        #options.add_argument('--headless')
        self.driver = webdriver.Firefox(options=options, service=FirefoxService(GeckoDriverManager().install()))
