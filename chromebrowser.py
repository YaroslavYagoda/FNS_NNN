from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime


class ChromeBrowser:
    """
    Браузер Chrome - родительский класс для других браузеров
    """

    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option('detach', True)
        # options.add_argument('--headless')
        self.driver = webdriver.Chrome(options=options, service=ChromeService(ChromeDriverManager().install()))

    def get_url(self, url: str):
        """
        Переход на страницу c адресом "url",
        размер окна браузера по размеру экрана
        :param url: адрес страницы в текстовом формате
        """
        self.driver.get(url)
        self.driver.maximize_window()

    def quit(self):
        """
        Закрытие браузера
        """
        self.driver.quit()

    def send_keys_by_id(self, valueID: str, data: str):
        """
        Ввод текстовой информации "data" в элемент на странице по его "valueID"
        :param valueID: ID элемента на странице
        :param data: вводимая текстовая информация
        """
        field = self.driver.find_element(By.ID, valueID)
        field.send_keys(data)

    def send_keys_by_xpath(self, xpath: str, data: str):
        """
        Ввод текстовой информации "data" в элемент на странице по его "xpath"
        :param xpath: XPATH элемента на странице
        :param data: вводимая текстовая информация
        """
        field = self.driver.find_element(By.XPATH, xpath)
        field.send_keys(data)

    def click_by_xpath(self, xpath: str):
        """
        Эмуляция нажатия(click) на элемент страницы по его "xpath"
        :param xpath: XPATH элемента на странице
        """
        field = self.driver.find_element(By.XPATH, xpath)
        field.click()

    def is_element_on_page(self, xpath: str):
        """
        Проверка наличия элемента на странице
        :param xpath:
        :return: True or False
        """
        try:
            self.driver.find_element(By.XPATH, xpath)
        except NoSuchElementException:
            return False
        return True
