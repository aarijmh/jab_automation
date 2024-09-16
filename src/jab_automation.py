from jab_action import JabAction
import csv
import time
from pyjab.jabdriver import JABDriver
from pyjab.common.win32utils import Win32Utils

class JabAutomation:
    def __init__(self, command_file : str):
        self._commands = self.make_list_of_commands(command_file)

    def init_commands(self, command_file : str):
        self._commands = self.make_list_of_commands(command_file)
        
    def make_list_of_commands(self, command_file : str) -> JabAction:
        # Parse the command
        commands = []
        with open(command_file, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                commands.append(JabAction(name = row[0], role = row[1], action = row[2], action_type=row[3], value=row[4], index_in_parent=row[5], element_depth=row[6]))
        return commands

    def run(self, jabdriver : JABDriver):
        for command in self._commands:
            print(f"Executing command: {command}")
            if command.action_type.lower() == "pause":
                time.sleep(float(command.value))
                print("sleep finished")
            elif command.role.lower() == "text":
                #elem = jabdriver.find_element_by_name(command.name)
                hasMulti = False
                elems = []
                if command.element_depth and int(command.element_depth) >= 0:
                    hasMulti = True
                    elems = jabdriver.find_elements_by_object_depth(int(command.element_depth))
                elif command.index_in_parent and int(command.index_in_parent) >= 0:
                    hasMulti = True
                    elems  = jabdriver.find_elements_by_xpath(f"//text[@name=contains('{command.name}')]")
                
                if hasMulti == True:
                    for e in elems:
                        # print(f"{e.index_in_parent}   {e.object_depth}")
                        if e.index_in_parent == int(command.index_in_parent):
                            elem = e
                            break
                else:
                    elem  = jabdriver.find_element_by_xpath(f"//text[@name=contains('{command.name}')]")
                elem.send_text(command.value, True)
            elif command.role.lower() == "password text":
                elem  = jabdriver.find_element_by_xpath(f"//password text[@name=contains('{command.name}')]")
                elem.send_text(command.value, True)
            elif command.role.lower() == "push button":
                elem  = jabdriver.find_element_by_xpath(f"//push button[@name=contains('{command.name}')]")
                elem.click()
            elif command.role.lower() == "message_box":
                hasMulti = False
                elems = []
                if command.element_depth and int(command.element_depth) >= 0:
                    hasMulti = True
                    elems = jabdriver.find_elements_by_object_depth(int(command.element_depth))
                elif command.index_in_parent and int(command.index_in_parent) >= 0:
                    hasMulti = True
                    elems  = jabdriver.find_elements_by_xpath(f"//push button[@name=contains('{command.name}')]")
                if hasMulti == True:
                    for e in elems:
                        # print(f"{e.index_in_parent}   {e.object_depth}")
                        if e.index_in_parent == int(command.index_in_parent):
                            elem = e
                            break
                else: elems  = jabdriver.find_elements_by_xpath(f"//push button[@name=contains('{command.name}')]")

                for ele in elems:
                    if (ele.name == 'OK ALT O') and ele.is_showing() == True:
                        elem = e
                        break
                elem.click()

            elif command.role.lower() == "label":
                elem = jabdriver.find_element_by_xpath(f"//label[@name=contains('{command.name}')]")
                elem.click()
            elif command.role.lower() == "check box":
                elem = jabdriver.find_element_by_xpath(f"//check box[@name=contains('{command.name}')]")
                elem.click()
            elif command.role.lower() == "option pane":
                elem = jabdriver.find_element_by_xpath(f"//option pane[@name=contains('{command.name}')]")
                elem.click()
            elif command.role.lower() == "key_press":
                try :
                    print(dir(jabdriver))
                    print(dir(jabdriver.win32utils))
                    jabdriver.win32utils._press_key(command.value)
                except Exception as ex:
                    print(ex)
                    continue

                # jabdriver.send_key('tab')
                
                # elem = jabdriver.find_element_by_xpath(f"//option pane[@name=contains('{command.name}')]")
                # elem.click()
            # elif command.role.lower() == "combo box":
            #     elem = jabdriver.find_element_by_xpath(f"//combo box[@name=contains('{command.name}')]")
            #     elem.click()
            #     li = elem.find_elements()
            #     for e in li:
            #         print(f"{e.role} {e.name}")
            #         if e.role == 'list':
            #             e.select(option='Yes')
            #             break

                #elem.select(option=command.value, simulate=False)
        print("All executed")
    
    def start(self, windows_title:str):
        if windows_title is not None:
            # Get handles of all windows
            window_handles = w32.enum_windows()
            for k,v in window_handles.items():
                try:
                    if v == windows_title:
                        jabdriver = JABDriver(hwnd=k)
                        self.run(jabdriver)
                        break
                except Exception as e:
                    if v == windows_title:
                        print(e)                    
                
if __name__ == "__main__":
    
    w32 = Win32Utils()
    jabAutomation = JabAutomation("commands.csv")
    # Get handles of all windows
    window_handles = w32.enum_windows()
    title = "SYMBOLS - [LVNDRATNETKHI]"
    for k,v in window_handles.items():
        #print(f"{k}   {v}")
        try:
            if v == title:
                jabdriver = JABDriver(hwnd=k)
                print("xuccdxx1")
                jabAutomation.run(jabdriver=jabdriver)
                break
        except Exception as e:
            if v == "SYMBOLS - [LVNDRATNETKHI]":
                print(e)
            
       
    
    
    #jabdriver = JABDriver("New Window Java")
    #jabautomation = JabAutomation("commands.csv")
    #jabautomation.run(jabdriver)