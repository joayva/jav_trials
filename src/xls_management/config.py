import yaml

from xls_management import HOMEPATH, ROOTPATH


class ATEConfig():
    config_file = ROOTPATH / "config.yml"
    def __init__(self):
        self.config:dict|None = None
        if(ATEConfig.config_file.exists()):
            self.load_config_file()
        else:
            self.set_default_config_file()

    def load_config_file(self) -> None:
        with open(ATEConfig.config_file, 'r') as file:
            self.config = yaml.safe_load(file)

    def set_default_config_file(self) -> None:
        self.config = {}
        self.config['workbook_path_BsM'] = str(
            HOMEPATH / 'vw/data/ATE-Status_Berichtsversion.xlsx',
        )
        self.config['default_path'] = str(
            HOMEPATH / 'vw/in'
        )
        self.config['blacklist_name'] = 'Blacklist'
        yaml_str = yaml.dump(self.config)
        with open(ATEConfig.config_file, 'w') as file:
            file.writelines(yaml_str)
        self.config['blacklist_attribute'] = 'LAH, die ignoriert werden sollen'
        yaml_str = yaml.dump(self.config)
        with open(ATEConfig.config_file, 'w') as file:
            file.writelines(yaml_str)
    
    def get(self, *args, **kvargs):
        return self.config.get(*args, **kvargs)
    
    def erase(self):
        if(ATEConfig.config_file.exists()):
            ATEConfig.config_file.unlink()
