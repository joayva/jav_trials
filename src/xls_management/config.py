import os
from pathlib import Path

import yaml

from xls_management import ROOTPATH, WORKPATH


class ATEConfig():
    config_file = WORKPATH / ".xls/config.yml"
    def __init__(self):
        self.config:dict = {}
        if(ATEConfig.config_file.exists()):
            self.load_config_file()
        else:
            self.set_default_config_file()

    def load_config_file(self) -> None:
        with open(ATEConfig.config_file, 'r', encoding='utf8') as file:
            config = yaml.safe_load(file)
            if config is None:
                config = {}
            self.config = config
    
    def config_from(self, file_path:str|Path) ->None:
        config_data:dict = {}
        with open(file_path, 'r', encoding='utf8') as file:
            config_data = yaml.safe_load(file)
        for key,value in config_data.items():
            _key:str = f'{key}'
            self.config[_key] = value

    def set_default_config_file(self) -> None:
        self.config = {}
        self.config['workbook_path_BsM'] = str(
            WORKPATH / 'vw/data/ATE-Status_Berichtsversion.xlsx',
        )
        self.config['default_path'] = str(
            WORKPATH / 'vw/in'
        )
        self.config['requirements_path'] = ''
        self.config['verification_criteria_path'] = ''
        self.config['security_orders_path'] = ''
        self.config['test_cases_path'] = ''
        self.config['timing_path'] = ''
        self.config['requirements_mb_path'] = ''
        self.config['blacklist_name'] = 'Blacklist'
        self.config['blacklist_attribute'] = 'LAH, die ignoriert werden sollen'
        self.config['header_style'] = {
            'font': {
                'name': 'Arial',
                'size': 10,
                'bold': True,
            },
            'fill': {
              'fill_type': 'solid',
              'fgColor': 'C3E8FF',
            }
        }    
        self.config['worksheet_widths'] = {
            'ATE_Status': [
                19.31,
                8.97,
                12.31,
                57.19,
                25.97,
                32.08,
                8.97,
                22.64,
                12.75,
                20.64,
                17.19,
                17.42,
                17.86,
                16.53,
                17.19,
                5.97,
                17.64,
                19.19,
                32.19,
                22.75,
                38.75,
                38.75,
                8.97,
                17.64,
                18.31,
                20.86,
                8.97,
                40.53,
                67.19,
                41.86,
                25.31,
                21.97,
                29.53,
                17.97,
                7.08,
            ],
            'TD_Status': [
                8.97,
                40.53,
                67.19,
                41.86,
                25.31,
                21.97,
                29.53,
                15.75,
                24.42,
                32.19,
                9.75,
                17.97,
                32.08,
                14.08,
                22.64,
                38.75,
                38.75,
                8.97,
                57.19,
                17.97,
                7.08,
            ],
        }
        yaml_str = yaml.dump(self.config)
        if not ATEConfig.config_file.parent.exists():
            os.makedirs(ATEConfig.config_file.parent, exist_ok=True)
        with open(ATEConfig.config_file, 'w',encoding='utf8') as file:
            file.writelines(yaml_str)
    
    def get(self, *args, **kvargs):
        return self.config.get(*args, **kvargs)
    
    def erase(self):
        if(ATEConfig.config_file.exists()):
            ATEConfig.config_file.unlink()
