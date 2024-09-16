import json
import time
import re
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
            elif role == "page tab list":
                self.handle_page_tab_list_command(jabdriver, command)
            elif role == "ctr_tab":
                self.handle_key_press_ctr_tab_command(jabdriver, command)
            elif role == "extract_digits":
                self.handle_extract_digits_command(jabdriver, command)
            elif role == "approve_client":
                self.handle_approve_client_command(jabdriver, command)
                
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
            try:
                elem.send_text(command['value'], True)
                jabdriver.win32utils._press_key('tab')
                break
            except Exception as ex:
                break
                # print(f"Retrying setting text, error: {ex}")    
                # if elem.text == command['value']:
                #     break
                # else:
                #     time.sleep(1)

    def handle_password_text_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//password text[@name=contains('{command['name']}')]")
        elem.send_text(command['value'], True)

    def handle_push_button_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//push button[@name=contains('{command['name']}')]")
        if elem.is_showing():
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
                if e.index_in_parent == int(command['index_in_parent']) and e.is_showing() and e.name == command['name'] and e.role == 'push button':
                    elem = e
                    break
        else:
            elems = jabdriver.find_elements_by_xpath(f"//push button[@name=contains('{command['name']}')]")
            for ele in elems:
                if ele.name == 'OK ALT O' and ele.is_showing():
                    elem = ele
                    break
        try:
            elem.click()
        except Exception as e:
            print(f"MessageBox not found: {e}")

    def handle_message_box_2_command(self, jabdriver, command):
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
                if e.index_in_parent == int(command['index_in_parent']) and e.is_showing() and e.name == 'OK ALT O':
                    elem = e
                    break
                elif e.name == 'OK ALT O' and e.is_showing() and e.role == 'push button':
                    elem = e
                    break
        else:
            elems = jabdriver.find_elements_by_xpath(f"//push button[@name=contains('{command['name']}')]")
            for ele in elems:
                if ele.name == 'OK ALT O' and ele.is_showing():
                    elem = ele
                    break
        try:
            elem.click()
        except Exception as e:
            print(f"MessageBox not found: {e}")

    def handle_label_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//label[@name=contains('{command['name']}')]")
        elem.click()

    def handle_check_box_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//check box[@name=contains('{command['name']}')]")
        elem.click()

    def handle_option_pane_command(self, jabdriver, command):
        elem = jabdriver.find_element_by_xpath(f"//option pane[@name=contains('{command['name']}')]")
        elem.click()

    def handle_approve_client_command(self, jabdriver, command):
        with open('tracking_id.txt') as file:
            trackingID = file.read()

        trackingNumbers = jabdriver.find_elements_by_xpath(f"//text[@name=contains('Track No Required')]")
        approveButtons = jabdriver.find_elements_by_xpath(f"//radio button[@name=contains('Accept')]")

        index = 0
        for trackingNmber in trackingNumbers:
            if len(trackingNmber.text) > 0:
                print(trackingNmber.text)
                if trackingID == trackingNmber.text:
                    approveButtons[index].click()
                    break
                index = index + 1

        # print(len(trackingNumbers))
        # print(len(approveButtons))
        # elem.click()

    def handle_combo_box_command(self, jabdriver, command):
        has_multi = False
        elems = []
        if command['element_depth'] and int(command['element_depth']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_object_depth(int(command['element_depth']))
        elif command['index_in_parent'] and int(command['index_in_parent']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_xpath(f"//combo box[@name=contains('{command['name']}')]")
        
        if has_multi:
            for e in elems:
                if e.index_in_parent == int(command['index_in_parent']) and e.is_showing():
                    combo_box = e
                    break
        else:
            combo_box = jabdriver.find_element_by_xpath(f"//combo box[@name=contains('{command['name']}')]")
        
        if combo_box:
            try:
                # combo_box.click()
                combo_box.select(command['value'], True, True)  # Use the select method
                print(f"Selected '{command['value']}' from combo box '{command['name']}'")
            except Exception as e:
                print(f"Failed to select '{command['value']}' from combo box '{command['name']}': {e}")
        else:
            raise ValueError(f"Combo box '{command['name']}' not found")

    def handle_page_tab_list_command(self, jabdriver, command):
        has_multi = False
        elems = []
        if command['element_depth'] and int(command['element_depth']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_object_depth(int(command['element_depth']))
        elif command['index_in_parent'] and int(command['index_in_parent']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_xpath(f"//page tab list[@name=contains('{command['name']}')]")
        
        if has_multi:
            for e in elems:
                if e.index_in_parent == int(command['index_in_parent']) and e.is_showing() and e.role == command["role"]:
                    tablist = e
                    break
        else:
            tablist = jabdriver.find_element_by_xpath(f"//page tab list[@name=contains('{command['name']}')]")
        
        if tablist:
            start_time = time.time()
            while time.time() - start_time < self._timeout:
                try:
                    # Try to locate the specific tab and select it
                    tablist.select(command['value'], True, True)
                    tab = tablist.find_element_by_xpath(f"//page tab[@name=contains('{command['value']}')]")
                    if tab and tab.is_showing():
                        # print(dir(tab))
                        # tab.click()
                        tab.select(command['value'], True, True)
                        print(f"Switched to tab '{command['value']}'")
                        break
                except Exception as ex:
                    print(f"Retrying selecting tab, error: {ex}")
                    time.sleep(1)
            else:
                print(f"Failed to select '{command['value']}' within the timeout period.")
        else:
            raise ValueError(f"Tab list '{command['name']}' not found or not interactable.")
    
    def handle_extract_digits_command(self, jabdriver, command):
        has_multi = False
        elems = []
        if command['element_depth'] and int(command['element_depth']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_object_depth(int(command['element_depth']))
        elif command['index_in_parent'] and int(command['index_in_parent']) >= 0:
            has_multi = True
            elems = jabdriver.find_elements_by_xpath(f"//label[@name=contains('{command['name']}')]")
        
        if has_multi:
            for e in elems:
                if e.index_in_parent == int(command['index_in_parent']) and e.is_showing() and e.role == 'label':
                    mLabel = e
                    break
        
        if mLabel:
            start_time = time.time()
            while time.time() - start_time < self._timeout:
                try:
                    temp = mLabel.name.split('Assigned')[1]
                    digits = ''.join(re.findall(r'\d+', temp))
                    print(f"Extracted digits '{digits}'")
                    with open(command['file_name'], 'w') as file:
                        file.write(digits)
                    break
                except Exception as ex:
                    print(f"Retrying extracting digits, error: {ex}")
                    time.sleep(1)
            else:
                print(f"Failed to extract digits within the timeout period.")
        else:
            raise ValueError(f"Label not found or not interactable.")
        
    def handle_key_press_command(self, jabdriver, command):
        try:
            jabdriver.win32utils._press_key(command['value'])
        except Exception as ex:
            print(ex)

    def handle_key_press_ctr_tab_command(self, jabdriver, command):
        try:
            jabdriver.win32utils._press_and_hold_key('ctrl')
            jabdriver.win32utils._press_key('tab')
            jabdriver.win32utils._press_hold_release_key('ctrl')
        except Exception as ex:
            print(ex)

if __name__ == "__main__":
    w32 = Win32Utils()
    jabAutomation = JabAutomation("client_approval.json")
    window_handles = w32.enum_windows()
    title = "SYMBOLS - [LVNDRATNETKHI]"
    for k, v in window_handles.items():
        try:
            if v == title:
                jabdriver = JABDriver(hwnd=k, timeout=60)
                jabAutomation.run(jabdriver=jabdriver)
                break
        except Exception as e:
            print(e)
