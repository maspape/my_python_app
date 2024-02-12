import json
import tkinter as tk
from tkinter import filedialog

def load_settings():
    try:
        with open('telnet_setting.json', 'r') as file:
            settings = json.load(file)
        return settings
    except FileNotFoundError:
        return {}

def save_settings(settings):
    with open('telnet_setting.json', 'w') as file:
        json.dump(settings, file, indent=4)

def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

def save_settings_and_close(settings, window):
    save_settings(settings)
    window.destroy()

def main():
    settings = load_settings()

    root = tk.Tk()
    root.title("Telnet Settings")

    tk.Label(root, text="Destination Folder:").grid(row=0, column=0, sticky='w')
    dest_entry = tk.Entry(root, width=50)
    dest_entry.grid(row=0, column=1, pady=5)
    dest_entry.insert(0, settings.get('dest', ''))

    dest_button = tk.Button(root, text="Browse", command=lambda: browse_folder(dest_entry))
    dest_button.grid(row=0, column=2)

    tk.Label(root, text="Username:").grid(row=1, column=0, sticky='w')
    username_entry = tk.Entry(root, width=50)
    username_entry.grid(row=1, column=1, pady=5)
    username_entry.insert(0, settings.get('username', ''))

    tk.Label(root, text="Password:").grid(row=2, column=0, sticky='w')
    password_entry = tk.Entry(root, width=50, show='*')
    password_entry.grid(row=2, column=1, pady=5)
    password_entry.insert(0, settings.get('password', ''))

    tk.Button(root, text="Save", command=lambda: save_settings_and_close({
        'dest': dest_entry.get(),
        'username': username_entry.get(),
        'password': password_entry.get()
    }, root)).grid(row=3, column=0, columnspan=3, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
