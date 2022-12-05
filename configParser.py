import configparser

class Settings():

    def __init__(self) -> None:
       self.config = configparser.ConfigParser()

       self.config.read('settings.ini')

    def get_from_settings(self,parametr):
        return self.config["DEFAULT"][parametr]

    def set_to_settings(self,parametr,value):
        self.config["DEFAULT"][parametr] = value
        with open('settings.ini', 'w') as configfile:
            self.config.write(configfile)