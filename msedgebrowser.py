from chromebrowser import ChromeBrowser
from selenium import webdriver
from selenium.webdriver.edge.service import Service as MsEdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager


class MsEdge(ChromeBrowser):
    """
    Браузер Microsoft Edge - дочерний класс "ChromeBrowser"
    """

    def __init__(self):
        options = webdriver.EdgeOptions()
        options.add_experimental_option('detach', True)
        #options.add_argument('--headless')
        self.driver = webdriver.Edge(options=options, service=MsEdgeService(EdgeChromiumDriverManager().install()))
