from chromebrowser import ChromeBrowser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService


class YaBrowser(ChromeBrowser):
    """
    Браузер ЯндексБраузер - дочерний класс "ChromeBrowser"
    """

    def __init__(self):
        binary_yandex_driver_file = 'driver/yandexdriver.exe'
        options = webdriver.ChromeOptions()
        options.add_experimental_option('detach', True)
        # options.add_argument('--headless')
        self.driver = webdriver.Chrome(options=options, service=ChromeService(binary_yandex_driver_file))
