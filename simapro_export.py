from shutil import copy2
import os, fnmatch
import subprocess
import pyautogui
import glob
import pyperclip
from time import sleep

"""
steps for simapro excel export:

1 based on file name of each model recreate database using 
template database (auto_simapro_db_export_template)

2 for each model build product system

3 execute process and product mapping commands

4 export each database

"""

class Mapping:

    def __init__(self):
        self.process_mapping_cmds = []
        self.product_mapping_cmds = []

def dir_paths():
    """
    sub simapro model folder path generator, eg Dairy/, Feed/, Insulation/ 
    """
    simapro_dirs = ['Dairy', 'Feed', 'Insulation', 'Leather',  \
    'Packed water', 'Pasta', 'Pet food', 'Retail']
    for dir_path, dirs, _ in os.walk('../'):
        for name in dirs:
            if name not in simapro_dirs:
                continue
            yield os.path.join(os.path.join(dir_path, name))

def copy_builder_to():
    """
    copy the process map builder to all the simapro sub folders for 
    generating sql files 
    """
    map_builder_file = os.path.join(os.getcwd(), 'sp_process_map_builder.py')
    if not os.path.exists(map_builder_file):
        print('{} muss be located in the root folder.' \
        .format('sp_process_map_builder.py'))
        return
    for dir in dir_paths():
        copy2(map_builder_file, dir)
                
def rename_dir():
    """
    ensure that all folder inside simapro sub folder with the same 
    folde name 'complete models'
    """
    for dir_path, dirs, _ in os.walk('../'):
        for name in dirs:
            if not name.startswith('complete'):
                continue
            if not name.endswith('models'):
                src = os.path.join(dir_path, name)
                dst = os.path.join(dir_path, 'complete models')
                os.rename(src, dst)

def exc_builder_file():
    for dir in dir_paths():
        os.chdir(dir)
        cmd = ['python ', 'sp_process_map_builder.py', ' ', '1']
        subprocess.call(''.join(cmd)) 

def rename_database(database: str):
    """
    based on database template recreate a database for each simapro model 
    and prepare for simapro csv import 
    """
    # template database auto_simapro_db_export_template position (156, 113)
    pyautogui.rightClick(156, 113)

    # copy database position (241, 216)
    pyautogui.click(241, 216)

    # rename database
    pyautogui.typewrite(database)

    sleep(2)

    # conformation for rename database (1139, 597)
    pyautogui.click(1139, 597)

    sleep(1)

def copy_database():
    model_dir = 'complete models'
    for dir in dir_paths():
        csv_path = os.path.join(dir, model_dir)
        files = os.listdir(csv_path)
        for file in files:
            if not fnmatch.fnmatch(file, '*.csv'):
                continue
            database = validate(file.split('.CSV')[0])
            prefix = 'auto_gen_' + database
            rename_database(prefix)
            sleep(1)

def open_dev_win():
    # open window menu
    pyautogui.click(151, 40, duration=1)

    # click on develop tools menu
    pyautogui.click(212, 101, duration=1)

def open_windows():
    open_dev_win()
    # open python window
    pyautogui.click(399, 141)
    open_dev_win()
    # open sql window
    pyautogui.click(404, 96)

def exc_py(cmds: list, first: int):
    """
    the exact mouse position should be first tested
    this function based strictly on the postion of mouse
    execute the python command to build product system
    only copy and paste at first run
    """
    # click on the python window tab
    pyautogui.click(523, 94)

    # edit field position (930, 165)
    pyautogui.click(648, 165, duration=1)

    if first == 0:
        # select all statments
        pyautogui.hotkey('ctrl', 'a')

        # copy commands to the edit fields
        pyperclip.copy(''.join(cmds))

        # paste to the edit field
        pyautogui.hotkey('ctrl', 'v')
    
        print('{} lines of command copied. '.format(len(cmds)))

    # execution button position (115, 65)
    pyautogui.click(115, 65, duration=1)

def exc_sql(cmds: list):
    # click on sql window tab
    pyautogui.click(766, 87)
    
    # edit position (765, 261)
    pyautogui.click(765, 261, duration = 1)

    # select all 
    pyautogui.hotkey('ctrl', 'a')

    # copy commands to the edit fields
    pyperclip.copy(''.join(cmds))

    # paste to the edit field
    pyautogui.hotkey('ctrl', 'v')

    print('{} lines of command copied. '.format(len(cmds)))
    # execution button position (115, 65)
    pyautogui.click(115, 65, duration = 2)

def read_py_file() -> list:
    """
    read the product system builder file 
    """
    file = os.path.join(os.getcwd(), 'olca_build_systems.py')
    if not os.path.exists(file):
        print('{} muss be located in the root folder.' \
        .format('olca_build_systems.py'))
        return
    with open(file, 'r', encoding = 'utf-8') as f:
        return f.readlines()

def read_sql_file() -> {}:
    """
    return a map of format model -> sql mapping commands 
    """
    model_2_sql = {}
    for dir in dir_paths():
        model = dir.split('/')[1]
        mapping_cmds = Mapping()
        for file in glob.glob('{}/*.sql'.format(dir)):
            with open(file, encoding = 'utf-8') as f:
                cmds = f.readlines()
                print('read {} lines of commands from file {}. ' \
                .format(str(len(cmds)), file))
                if file.endswith('process_mappings.sql'):
                    mapping_cmds.process_mapping_cmds = cmds
                if file.endswith('product_mappings.sql'):
                    mapping_cmds.product_mapping_cmds = cmds
        model_2_sql[model.lower()] = mapping_cmds
    return model_2_sql

# TODO: may automate csv import later on
# TODO: product sql mapping execution causes program freezing 
# may detect program status and wait until execution is done
def iter_database():
    """
    first build product system and then execute sql file 
    start position (205, 116), vertical offset = 23
    """
    py_cmds = read_py_file()
    model_sql_cmds = read_sql_file() 
    
    x_pos = 205
    y_pos = 116

    # open python and sql window
    open_windows()

    sleep(2)

    for i in range(0, 26):
        # open database
        pyautogui.doubleClick(x_pos, y_pos, duration = 1)

        # sleep 1 sec
        sleep(1)

        exc_py(py_cmds, i)

        if i < 5:
            cmds = model_sql_cmds['dairy']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds) 
        elif i == 5:
            cmds = model_sql_cmds['feed']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)
        elif i < 9 or i == 19:
            cmds = model_sql_cmds['insulation']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)
        elif i < 12:
            cmds = model_sql_cmds['packed water']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)
        elif i < 15:
            cmds = model_sql_cmds['pasta']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)
        elif i < 19:
            cmds = model_sql_cmds['pet food']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)
        elif i < 22:
            cmds = model_sql_cmds['retail']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)
        else:
            cmds = model_sql_cmds['leather']
            exc_sql(cmds.process_mapping_cmds)
            #exc_sql(cmds.product_mapping_cmds)

        sleep(1)

        # close database
        pyautogui.doubleClick(x_pos, y_pos)
        # vertival off set 
        y_pos += 23

def exc_bash_file():
    """
    execute bash file that located in ../pef_remod_xlsx-0.0.2 folder 
    """
    pass

chars = ['(', ')', ',', '.']

def validate(label: str):
    for ch in chars:
        if ch in label:
            label = label.replace(ch, '')
    return label.replace(' ', '_')

def main():
    #copy_builder_to()
    #exc_builder_file()
    #print(pyautogui.position())
    #copy_database()
    # may automate import csv for each model later
    iter_database()

if __name__ == '__main__':
    main()

        