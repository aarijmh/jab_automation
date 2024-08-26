from pyjab.jabdriver import JABDriver
from pyjab.common.win32utils import Win32Utils

# Initialize Win32Utils
w32 = Win32Utils()

# Get handles of all windows
window_handles = w32.enum_windows()

# Print titles of all windows
for k, v in window_handles.items():
    title = k
    try:
        jabdriver = JABDriver(hwnd=k)
        print(f"Success {k}   {v}")
        elements = jabdriver.find_elements()
        if v == "SYMBOLS - [LVNDRATNETKHI]":
            for element in elements:
                print(f"Element: {element.name}, Role: {element.role}, Description: {element.description}, Tag : {element.index_in_parent}")	
                if (element.role == 'text') and (element.name == 'Enter UserName Required'):
                    try :
                        # print('DONE')
                        # print(dir(element))
                        element.send_text("692548", True)
                    except Exception as ex:
                        print(ex)
                if (element.role == 'password text') and (element.name == 'Enter Password'):
                    try :
                        # print('DONE')
                        # print(dir(element))
                        element.send_text("cbs", True)
                    except Exception as ex:
                        print(ex)
                if (element.role == 'push button') and (element.name == 'Connect alt o'):
                    try :
                        print('DONE')
                        print(dir(element))
                        element.click()
                    except Exception as ex:
                        print(ex)
                # if (element.role == 'text') and (element.name == 'Enter UserName Required'):
                #     try :
                #         print('DONE')
                #         element.set_text("ASAD")
                #     except ex:
                #         print(ex)

    except: pass
