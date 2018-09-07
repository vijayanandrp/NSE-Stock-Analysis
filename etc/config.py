# -*- coding: utf-8 -*-


import configparser
from lib.util import OsUtil, JsonUtil

config = configparser.ConfigParser()
config_file = OsUtil.join_path(OsUtil.current_directory(__file__), 'config.ini')
config.read(config_file)

if not OsUtil.check_if_file_exist(config_file):
    print(config_file)
    raise FileNotFoundError

# [ DIRECTORY CONFIGURATIONS ]
path_root = OsUtil.root_directory(__file__)
path_data = OsUtil.join_path(path_root, config['path']['data'], create_dir=True)
path_etc = OsUtil.join_path(path_root, config['path']['etc'], create_dir=True)
path_log = OsUtil.join_path(path_data, config['path']['logs'], create_dir=True)
path_files = OsUtil.join_path(path_data, config['path']['files'], create_dir=True)
path_output = OsUtil.join_path(path_data, config['path']['output'], create_dir=True)
path_tmp = OsUtil.join_path(path_data, config['path']['tmp'], create_dir=True)

# [ FILE CONFIGURATIONS ]
file_log = config['file']['log']
file_masterdata = config['file']['masterdata']
file_insight = config['file']['insight']
file_52w = config['file']['52w_insight']


# [URL CONFIGURATIONS]
broad_indices = config['url']['broad_indices']
get_quote = config['url']['get_quote']

# SCORE
index_score = config['value']['index_score']
index_score = JsonUtil.read(OsUtil.join_path(path_etc, index_score))
index_score = index_score['index_score']
