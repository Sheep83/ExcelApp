from DataService import DataService
from ListItem import ListItem


class Application:
    def __init__(self, window, dataservice):
        self.window = window
        self.dataservice = DataService()

    def populateMenu(self, list):
        options = []
        for item in list:
            options.append(str(item.name))
        return options
