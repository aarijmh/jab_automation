from jab_action import JabAction
import csv
import time
from pyjab.jabdriver import JABDriver
from pyjab.common.win32utils import Win32Utils

class JabAutomation:
    def __init__(self, command_file : str):
        self._commands = self.make_list_of_commands(command_file)
        
    def make_list_of_commands(self, command_file : str) -> JabAction:
        # Parse the command
        commands = []
        with open(command_file, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                commands.append(JabAction(name = row[0], role = row[1], action = row[2], action_type=row[3], value=row[4]))
        return commands

    def run(self, jabdriver : JABDriver):
        for command in self._commands:
            if command.action_type.lower() == "pause":
                time.sleep(float(command.value))
            elif command.role.lower() == "text":
                elem = jabdriver.find_element_by_name(command.name)
                elem.send_text(command.value, True)
            elif command.role.lower() == "push button":
                elem = jabdriver.find_element_by_name(command.name)
                elem.click()
            elif command.role.lower() == "label":
                elem = jabdriver.find_element_by_xpath(f"//label[@name=contains('{command.name}')]")
                elem.click()
                
if __name__ == "__main__":
    
    w32 = Win32Utils()
    jabAutomation = JabAutomation("commands1.csv")
    # Get handles of all windows
    window_handles = w32.enum_windows()
    
    for k,v in window_handles.items():
        try:
            jabdriver = JABDriver(hwnd=k)
            jabAutomation.run(jabdriver=jabdriver)
            break
        except Exception as e: pass
            
       
    
    
    #jabdriver = JABDriver("New Window Java")
    #jabautomation = JabAutomation("commands.csv")
    #jabautomation.run(jabdriver)