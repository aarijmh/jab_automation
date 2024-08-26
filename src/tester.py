from pywinauto.application import Application
from pywinauto import Desktop

# Print titles of all active windows
windows = Desktop(backend="uia").windows()
wi = None
for window in windows:
    print(f"{window.window_text()} -- {window.class_name()}")
    if window.window_text() == "SYMBOLS - [LVNDRATNETKHI]" and window.class_name() == "SunAwtFrame":
        wi = window
        break
wi.print_control_identifiers()
# Connect to the running Java application
app = Application(backend="uia").connect(title_re="Dynamic Tree Demo")

# Access elements in the Java application
main_window = app.window(title_re="Dynamic Tree Demo")
main_window.print_control_identifiers()

# Interact with a specific control, for example, a button
button = main_window.child_window(title="Button Name", control_type="Button")
button.click()