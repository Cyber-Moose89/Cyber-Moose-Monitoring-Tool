import configparser

def load_config(file_name='config.ini'):
    config = configparser.ConfigParser()
    config.read(file_name)
    return config

def save_config(config, file_name='config.ini'):
    with open(file_name, 'w') as configfile:
        config.write(configfile)
