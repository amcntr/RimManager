import win32api
import os
import configparser

# rimworld install directory
rimworld_root = ""
# mod directory in install directory
rimworld_moddir = ""
# config directory
rimworld_config = ""
# rimworld steam workshop directory
steam_workshop = ""
# rimworld appid
appid = 294100
# mod list
# key is mod folder name
# value is path to mod
mod_list = {}


# find the install directory of rimworld
# and the mods folder in said directory
def find_root():
    global rimworld_root, rimworld_moddir
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    for drive in drives:
        for root, dirs, files in os.walk(drive):
            if "RimWorldWin.exe" in files:
                rimworld_root = root
                rimworld_moddir = rimworld_root + "\\Mods"
                return True
    return False


# find the config directory for rimworld
def find_config():
    global rimworld_config
    path = os.getenv('LOCALAPPDATA')
    path += 'Low\\Ludeon Studios\\RimWorld by Ludeon Studios\\Config'
    if not os.path.isdir(path):
        return False
    rimworld_config = path
    return True


# find the workshop directory for rimworld
def find_workshop():
    global steam_workshop, appid
    path = ""
    if "steam" in rimworld_root:
        path = rimworld_root.replace("common\\RimWorld", "")
        path += "workshop\\content\\" + str(appid)
        if not os.path.isdir(path):
            return False
    steam_workshop = path
    return True


# generates a config file that stores
# the directory locations
def generate_config(config):
    if not find_root():
        print("Failed to find rimworld install directory.\n")
        return False
    if not find_config():
        print("Failed to find rimworld config directory.\n")
        return False
    if not find_workshop():
        print("Failed to find rimworld workshop directory.\n")
        return False
    config['paths'] = {'root': rimworld_root,
                       'mods': rimworld_moddir,
                       'workshop': steam_workshop,
                       'config': rimworld_config}
    return True


# get the available mods from the steam workshop folder
# and the root mod folder and puts them in the mod_list
# dictionary
def get_mods():
    global mod_list
    for rmod in os.listdir(rimworld_moddir):
        if os.path.isdir(os.path.join(rimworld_moddir, rmod)):
            mod_list[rmod] = os.path.join(rimworld_moddir, rmod)
    if steam_workshop:
        for smod in os.listdir(steam_workshop):
            if os.path.isdir(os.path.join(steam_workshop, smod)):
                mod_list[smod] = os.path.join(steam_workshop, smod)
    for modname, moddir in mod_list.items():
        print(modname, ':', moddir)


# pull data from the config file and
# populates the mod_list dictionary
def initialize():
    global rimworld_root, rimworld_moddir, rimworld_config, steam_workshop
    config = configparser.ConfigParser()
    config.read('config.ini')
    if not config.sections():
        if not generate_config(config):
            os.remove()
            return False
        with open('config.ini', 'w') as configfile:
            config.write(configfile)
    else:
        rimworld_root = config['paths']['root']
        rimworld_moddir = config['paths']['mods']
        rimworld_config = config['paths']['config']
        steam_workshop = config['paths']['workshop']
    get_mods()
    return True


if __name__ == '__main__':
    if not initialize():
        print("\nInitialization failed.")
    print("\nInitalized.")
