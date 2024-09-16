import json
import time
from pyjab.jabdriver import JABDriver
from pyjab.common.win32utils import Win32Utils

class JabAutomation:
    def __init__(self, command_file: str):
        self._commands = self.load_commands(command_file)
        self._timeout = 20

    def load_commands(self, command_file: str):
        with open(command_file, 'r') as f:
            return json.load(f)

    def run(self, jabdriver: JABDriver):
        for command in self._commands:
            print(f"Executing command: {command}")
            action_type = command['action_type'].lower()
            role = command['role'].lower()

            if action_type == "pause":
                time.sleep(float(command['value']))
                print("Sleep finished")
            elif role == "text":
                self.handle_text_command(jabdriver, command)
            elif role == "password text":
                self.handle_password_text_command(jabdriver, command)
            elif role == "push button":
                self.handle_push_button_command(jabdriver, command)
            elif role == "message_box":
                self.handle_message_box_command(jabdriver, command)
            elif role == "label":
                self.handle_label_command(jabdriver, command)
            elif role == "check box":
                self.handle_check_box_command(jabdriver, command)
            elif role == "option pane":
                self.handle_option_pane_command(jabdriver, command)
            elif role == "key_press":
                self.handle_key_press_command(jabdriver, command)
            elif role == "combo box":
                self.handle_combo_box_command(jabdriver, command)
        print("All executed")

    def handle_text_command(self, jabdriver, command):
        has_multi = False
        elems = []
        if command['element_depth'] and int(command['element_depth']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_object_depth(int(command['element_depth']))
        elif command['index_in_parent'] and int(command['index_in_parent']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_xpath(f"//text[@name=contains('{command['name']}')]")

        if has_multi:
            for e in elems:
                if e.index_in_parent == int(command['index_in_parent']) and e.is_showing() == True:
                    elem = e
                    break
        else:
            elem = jabdriver.find_element_by_xpath(f"//text[@name=contains('{command['name']}')]")
        
        start_time = time.time()
        while time.time() - start_time < self._timeout:
            try :
                elem.send_text(command['value'], True)
                jabdriver.win32utils._press_key('tab')
                break
            except Exception as ex:
                print(f"Retrying setting text, error: {ex}")
                time.sleep(1)

    def handle_password_text_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//password text[@name=contains('{command['name']}')]")
        elem.send_text(command['value'], True)

    def handle_push_button_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//push button[@name=contains('{command['name']}')]")
        elem.click()

    def handle_message_box_command(self, jabdriver, command):
        has_multi = False
        elems = []
        if command['element_depth'] and int(command['element_depth']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_object_depth(int(command['element_depth']))
        elif command['index_in_parent'] and int(command['index_in_parent']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_xpath(f"//push button[@name=contains('{command['name']}')]")
        
        if has_multi:
            for e in elems:
                if e.index_in_parent == int(command['index_in_parent']):
                    elem = e
                    break
                elif e.name == 'OK ALT O' and e.is_showing():
                    elem = e
                    break
        else:
            elems = jabdriver.find_elements_by_xpath(f"//push button[@name=contains('{command['name']}')]")
            for ele in elems:
                if ele.name == 'OK ALT O' and ele.is_showing():
                    elem = ele
                    break
        
        elem.click()

    def handle_label_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//label[@name=contains('{command['name']}')]")
        elem.click()

    def handle_check_box_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//check box[@name=contains('{command['name']}')]")
        elem.click()

    def handle_option_pane_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//option pane[@name=contains('{command['name']}')]")
        elem.click()

    def handle_combo_box_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//option pane[@name=contains('{command['name']}')]")
        elem.click()

    def handle_key_press_command(self, jabdriver, command):
        try:
            jabdriver.win32utils._press_key(command['value'])
        except Exception as ex:
            print(ex)

if __name__ == "__main__":
    w32 = Win32Utils()
    jabAutomation = JabAutomation("commands.json")
    window_handles = w32.enum_windows()
    title = "SYMBOLS - [LVNDRATNETKHI]"
    for k, v in window_handles.items():
        try:
            if v == title:
                jabdriver = JABDriver(hwnd=k, timeout=60)
                jabAutomation.run(jabdriver=jabdriver)
                break
        except Exception as e:
            # if v == "SYMBOLS - [LVNDRATNETKHI]":
            print(e)
