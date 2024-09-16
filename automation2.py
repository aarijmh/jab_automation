import json
import time
from pyjab.jabdriver import JABDriver
from pyjab.common.win32utils import Win32Utils
from concurrent.futures import ThreadPoolExecutor

class JabAutomation:
    def __init__(self, command_file: str):
        self._commands = self.load_commands(command_file)
        self._element_cache = {}

    def load_commands(self, command_file: str):
        with open(command_file, 'r') as f:
            return json.load(f)

    def custom_find_element(self, jabdriver, name, role, index_in_parent=None, element_depth=None):
        # Check cache first
        cache_key = (name, role, index_in_parent, element_depth)
        if cache_key in self._element_cache:
            return self._element_cache[cache_key]

        # Start with object depth if available
        elems = []
        if element_depth is not None:
            elems = jabdriver.find_elements_by_object_depth(element_depth)
        else:
            xpath = self.optimize_xpath(role, name)
            elems = jabdriver.find_elements_by_xpath(xpath)

        # If index_in_parent is specified, filter by it
        if index_in_parent is not None:
            elems = [e for e in elems if e.index_in_parent == index_in_parent]

        # Cache the result if an element was found
        if elems:
            self._element_cache[cache_key] = elems[0]
            return elems[0]
        return None

    def optimize_xpath(self, role, name):
        # Simplify the XPath query to speed up element search
        # return f"//{role}[contains(@name, '{name}')]"
        # f"//push button[@name=contains('{command['name']}')]"
        return f"//{role}[@name=contains('{name}')]"

    def run(self, jabdriver: JABDriver):
        # Use ThreadPoolExecutor for parallel processing
        with ThreadPoolExecutor(max_workers=4) as executor:
            for command in self._commands:
                executor.submit(self.execute_command, jabdriver, command)
        print("All executed")

    def execute_command(self, jabdriver, command):
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

    def handle_text_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'text', command['index_in_parent'], command['element_depth'])
        if elem:
            elem.send_text(command['value'], True)

    def handle_password_text_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'password text', command['index_in_parent'], command['element_depth'])
        if elem:
            elem.send_text(command['value'], True)

    def handle_push_button_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'push button', command['index_in_parent'], command['element_depth'])
        if elem:
            elem.click()

    def handle_message_box_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'push button', command['index_in_parent'], command['element_depth'])
        if elem:
            elem.click()

    def handle_label_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'label', command['index_in_parent'], command['element_depth'])
        if elem:
            elem.click()

    def handle_check_box_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'check box', command['index_in_parent'], command['element_depth'])
        if elem:
            elem.click()

    def handle_option_pane_command(self, jabdriver, command):
        elem = self.custom_find_element(jabdriver, command['name'], 'option pane', command['index_in_parent'], command['element_depth'])
        if elem:
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
                jabdriver = JABDriver(hwnd=k)
                jabAutomation.run(jabdriver=jabdriver)
                break
        except Exception as e:
            if v == "SYMBOLS - [LVNDRATNETKHI]":
                print(e)
