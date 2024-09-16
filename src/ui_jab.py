import tkinter as tk
from tkinter import filedialog, messagebox
from jab_automation import JabAutomation  # Assuming jab_automation.py is in the same directory

# Global variable to store the selected configuration file path
config_file_path = ""
jabAutomation = JabAutomation()
# Function to select a configuration file
def select_config_file():
    global config_file_path
    config_file_path = filedialog.askopenfilename(
        title="Select Configuration File",
        filetypes=(("Config Files", "*.json"), ("All Files", "*.*"))
    )
    if config_file_path:
        config_label.config(text=f"Selected: {config_file_path}")

# Function to start automation
def start_automation():
    if not config_file_path:
        messagebox.showerror("Error", "Please select a configuration file first.")
        return
    try:
        # Call the appropriate function from jab_automation with the config file
        jabAutomation.init_commands(config_file_path)
        jabAutomation.start("SYMBOLS - [LVNDRATNETKHI]")

        messagebox.showinfo("Success", "Automation started successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start automation: {e}")

# Function to stop automation
def stop_automation():
    try:
        # Call the appropriate function from jab_automation
        jabAutomation.stop()
        messagebox.showinfo("Success", "Automation stopped successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to stop automation: {e}")

# Create the main application window
root = tk.Tk()
root.title("Oracle Forms Automation")

# Add UI elements
tk.Label(root, text="Oracle Forms Automation").pack(pady=10)

config_button = tk.Button(root, text="Select Configuration File", command=select_config_file)
config_button.pack(pady=5)

config_label = tk.Label(root, text="No file selected")
config_label.pack(pady=5)

start_button = tk.Button(root, text="Start Automation", command=start_automation)
start_button.pack(pady=5)

stop_button = tk.Button(root, text="Stop Automation", command=stop_automation)
stop_button.pack(pady=5)

# Run the main loop
root.mainloop()