import win32api
import os
import configparser
import tkinter as tk
import lxml.etree as ET


class ModList(tk.Frame):
    def __init__(self, _data, master=None):
        self.data = _data
        super().__init__(master)
        self.create_widgets()
        self.populate()
        self.update()

    def create_widgets(self):
        self.info_frame = tk.Frame(self)
        self.info_frame.grid(row=0,
                             column=0,
                             rowspan=8,
                             columnspan=6)
        self.title_label = tk.Label(self,
                                    text="Installed Mods")
        self.title_label.grid(row=0, column=7)
        self.modlist = tk.Listbox(self,
                                  selectmode=tk.EXTENDED,
                                  exportselection=0,
                                  activestyle='none',
                                  width=40,
                                  height=40)
        self.modlist.grid(row=1,
                          column=7,
                          columnspan=3,
                          rowspan=8)

        self.yscroll = tk.Scrollbar(command=self.modlist.yview,
                                    orient=tk.VERTICAL)
        self.yscroll.grid(row=0, column=5, sticky=tk.N + tk.S)
        self.modlist.configure(yscrollcommand=self.yscroll.set)

        self.buttonAdd = tk.Button(self.info_frame,
                                   text="Add",
                                   command=self.add)
        self.buttonAdd.grid(row=0, column=0)
        self.buttonRefresh = tk.Button(self.info_frame,
                                       text="Refresh",
                                       command=self.refresh)
        self.buttonRefresh.grid(row=0, column=1)

    def populate(self):
        self.modlist.delete(0, tk.END)
        for index, mod in self.data.mods_unactive.items():
            name = self.data.mods_list[mod][0].text
            self.modlist.insert(tk.END, name)

    def update(self):
        if self.data.unactive_mods_change:
            self.populate()
        self.data.unactive_mods_change = False
        self.after(50, self.update)

    def add(self):
        lines = self.modlist.curselection()
        parser = ET.XMLParser(remove_blank_text=True)
        config = ET.parse(self.data.rimworld_config, parser)
        root = config.getroot()
        for line in lines:
            mod = self.data.mods_unactive[line]
            element = ET.Element('li')
            element.text = mod
            root[1].append(element)
            del self.data.mods_unactive[line]
        config.write(self.data.rimworld_config,
                     pretty_print=True,
                     encoding="utf-8",
                     xml_declaration=True)
        self.data.active_mods_change = True
        self.data.unactive_mods_change = True

    def refresh(self):
        self.data.find_mods()
        self.populate()


class ActiveList(tk.Frame):
    def __init__(self, _data, master=None):
        self.data = _data
        super().__init__(master)
        self.create_widgets()
        self.populate()
        self.update()

    def create_widgets(self):
        self.title_label = tk.Label(self,
                                    text="Active Mods")
        self.title_label.grid(row=0, column=0)
        self.activelist = tk.Listbox(self,
                                     selectmode=tk.EXTENDED,
                                     exportselection=0,
                                     activestyle='none',
                                     width=40,
                                     height=40)
        self.activelist.grid(row=1,
                             columnspan=3,
                             column=0)

        self.yscroll = tk.Scrollbar(self,
                                    command=self.activelist.yview,
                                    orient=tk.VERTICAL)
        self.yscroll.grid(row=1, column=4, sticky=tk.N + tk.S)
        self.activelist.configure(yscrollcommand=self.yscroll.set)

        self.buttonGrid = tk.Frame(self)
        self.buttonGrid.grid(row=2, column=0, sticky=tk.W)
        self.buttonup = tk.Button(self.buttonGrid,
                                  text="Up",
                                  command=self.move_up)
        self.buttonup.grid(row=0, column=0)
        self.buttondown = tk.Button(self.buttonGrid,
                                    text="Down",
                                    command=self.move_down)
        self.buttondown.grid(row=0, column=1)
        self.buttonrm = tk.Button(self,
                                  text="Remove",
                                  command=self.remove)
        self.buttonrm.grid(row=2, column=2)

    def populate(self):
        config = ET.parse(self.data.rimworld_config)
        root = config.getroot()
        for mod in root[1]:
            name = self.data.mods_list[mod.text][0].text
            self.activelist.insert(tk.END, name)

    def update(self):
        if self.data.active_mods_change:
            lines = self.activelist.curselection()
            self.activelist.delete(0, tk.END)
            self.populate()
            for line in lines:
                self.activelist.select_set(line)
            self.data.active_mods_change = False
        self.after(50, self.update)

    def move_up(self):
        lines = self.activelist.curselection()
        parser = ET.XMLParser(remove_blank_text=True)
        config = ET.parse(self.data.rimworld_config, parser)
        root = config.getroot()
        newLines = []
        for line in lines:
            if line > 0:
                item = root[1][line]
                root[1].remove(item)
                root[1].insert(line - 1, item)
                newLines.append(line - 1)
                self.activelist.selection_clear(line)
        for line in newLines:
            self.activelist.select_set(line)
        config.write(self.data.rimworld_config,
                     pretty_print=True,
                     encoding="utf-8",
                     xml_declaration=True)
        self.data.active_mods_change = True

    def move_down(self):
        lines = self.activelist.curselection()
        parser = ET.XMLParser(remove_blank_text=True)
        config = ET.parse(self.data.rimworld_config, parser)
        root = config.getroot()
        newLines = []
        for line in reversed(lines):
            if line < (self.activelist.size() - 1):
                item = root[1][line]
                root[1].remove(item)
                root[1].insert(line + 1, item)
                newLines.append(line + 1)
                self.activelist.selection_clear(line)
        for line in newLines:
            self.activelist.select_set(line)
        config.write(self.data.rimworld_config,
                     pretty_print=True,
                     encoding="utf-8",
                     xml_declaration=True)
        self.data.active_mods_change = True

    def remove(self):
        lines = self.activelist.curselection()
        parser = ET.XMLParser(remove_blank_text=True)
        config = ET.parse(self.data.rimworld_config, parser)
        root = config.getroot()
        count = 0
        for line in lines:
            root[1].remove(root[1][line - count])
            count += 1
        config.write(self.data.rimworld_config,
                     pretty_print=True,
                     encoding="utf-8",
                     xml_declaration=True)
        self.data.check_mods()
        self.data.active_mods_change = True
        self.data.unactive_mods_change = True


class Data():
    def __init__(self):
        # rimworld install directory
        self.rimworld_root = ""
        # mod directory in install directory
        self.rimworld_moddir = ""
        # config directory
        self.rimworld_config = ""
        # rimworld steam workshop directory
        self.steam_workshop = ""
        # rimworld appid
        self.appid = 294100
        # mod list
        # key is mod folder name
        # value is path to mod
        self.mods_unactive = {}
        self.mods_list = {}
        self.unactive_mods_change = False
        self.active_mods_change = False
        self.config = configparser.ConfigParser()

    # pull data from the config file and
    # populates the mods_list dictionary
    def load(self):
        self.config.read('config.ini')
        if not self.config.sections():
            if not self.find_root():
                print("Failed to find rimworld install directory.\n")
                return False
            if not self.find_config():
                print("Failed to find rimworld config directory.\n")
                return False
            if not self.find_workshop():
                print("Failed to find rimworld workshop directory.\n")
                return False
            self.config['paths'] = {'root': self.rimworld_root,
                                    'mods': self.rimworld_moddir,
                                    'workshop': self.steam_workshop,
                                    'config': self.rimworld_config}
            with open('config.ini', 'w') as configfile:
                self.config.write(configfile)
        else:
            self.rimworld_root = self.config['paths']['root']
            self.rimworld_moddir = self.config['paths']['mods']
            self.rimworld_config = self.config['paths']['config']
            self.steam_workshop = self.config['paths']['workshop']
        self.find_mods()
        return True

    # find the install directory of rimworld
    # and the mods folder in said directory
    def find_root(self):
        drives = win32api.GetLogicalDriveStrings()
        drives = drives.split('\000')[:-1]
        for drive in drives:
            for root, dirs, files in os.walk(drive):
                if "RimWorldWin.exe" in files:
                    self.rimworld_root = root
                    self.rimworld_moddir = self.rimworld_root + "\\Mods"
                    return True
        return False

    # find the config directory for rimworld
    def find_config(self):
        path = os.getenv('LOCALAPPDATA')
        path += 'Low\\Ludeon Studios\\RimWorld by Ludeon Studios\\Config'
        if not os.path.isdir(path):
            return False
        self.rimworld_config = path + "\\ModsConfig.xml"
        return True

    # find the workshop directory for rimworld
    def find_workshop(self):
        path = ""
        if "steam" in self.rimworld_root:
            path = self.rimworld_root.replace("common\\RimWorld", "")
            path += "workshop\\content\\" + str(self.appid)
            if not os.path.isdir(path):
                return False
        self.steam_workshop = path
        return True

    # get the available mods from the steam workshop folder
    # and the root mod folder and puts them in the mods_list
    # dictionary
    def find_mods(self):
        self.mods_list.clear()
        for rmod in os.listdir(self.rimworld_moddir):
            if os.path.isdir(os.path.join(self.rimworld_moddir, rmod)):
                path = os.path.join(self.rimworld_moddir, rmod)
                file = path + "\\About\\About.xml"
                if os.path.isfile(file):
                    tree = ET.parse(file)
                    self.mods_list[rmod] = tree.getroot()
        if self.steam_workshop:
            for smod in os.listdir(self.steam_workshop):
                if os.path.isdir(os.path.join(self.steam_workshop, smod)):
                    path = os.path.join(self.steam_workshop, smod)
                    file = path + "\\About\\About.xml"
                    if os.path.isfile(file):
                        tree = ET.parse(file)
                        self.mods_list[smod] = tree.getroot()
        self.check_mods()

    def check_mods(self):
        self.mods_unactive.clear()
        active_mods = {}
        config = ET.parse(self.rimworld_config)
        root = config.getroot()
        for mod in root[1]:
            active_mods[mod.text] = 1
        index = 0
        for mod, data in self.mods_list.items():
            try:
                if active_mods[mod] > 0:
                    pass
            except(KeyError):
                self.mods_unactive[index] = mod
                index += 1


if __name__ == '__main__':
    data = Data()
    if not data.load():
        print("\nInitialization failed.")
    window = tk.Tk()
    window.title("RimManager")
    activelist = ActiveList(data,
                            master=window)
    activelist.grid(row=0, column=0)
    modlist = ModList(data,
                      master=window)
    modlist.grid(row=0, column=2)
    window.mainloop()
