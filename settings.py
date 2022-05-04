from pathlib import Path
from configparser import ConfigParser

path = str(Path.home()) + '\.web\date-verification'
config = ConfigParser()

if Path(path + '\config.ini').is_file():
    config.read(path + '\config.ini')
else:
    try:
        Path(path).mkdir(parents=True)
    except FileExistsError:
        pass
    config['ShippingOptions'] = {'TransitDays': 3, 'IncludeWeekends': False}
    config.write(open(path + '\config.ini', 'w'))

def load():
    return config

def save(new_config):
    new_config.write(open(path + '\config.ini', 'w'))
