import os
import json
import xlwings as xw
import win32com.client
import tkinter as tk
from tkinter import filedialog

def load_settings():
    try:
        with open('mail_search_settings.json', 'r') as file:
            settings = json.load(file)
        return settings
    except FileNotFoundError:
        return {}

def save_settings(settings):
    with open('mail_search_settings.json', 'w') as file:
        json.dump(settings, file, indent=4)

def search_emails(folder_path, search_options):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders.Item(1)
    folder = inbox.Folders.ItemFromPath(folder_path)
    
    search_results = []
    for email in folder.Items:
        if email.MessageClass == "IPM.Note":
            subject = email.Subject
            body = email.Body
            for option in search_options:
                if option['search_range'] == 'Subject':
                    if option['search_type'] == 'Keyword':
                        if option['keyword'] in subject:
                            search_results.append((subject, body))
                    elif option['search_type'] == 'Regex':
                        import re
                        if re.search(option['regex'], subject):
                            search_results.append((subject, body))
                elif option['search_range'] == 'Body':
                    if option['search_type'] == 'Keyword':
                        if option['keyword'] in body:
                            search_results.append((subject, body))
                    elif option['search_type'] == 'Regex':
                        import re
                        if re.search(option['regex'], body):
                            search_results.append((subject, body))
    return search_results

def export_to_excel(search_results):
    wb = xw.Book()
    ws = wb.sheets['Sheet1']
    ws.range('A1').value = ['Subject', 'Body']
    for i, result in enumerate(search_results, start=2):
        ws.range(f'A{i}').value = result
    timestamp = xw.utils.datetime.datetime.now().strftime('%y%m%d')
    wb.save(f'SearchResults-{timestamp}.xlsx')
    wb.close()

def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

def add_search_option(search_options, search_range, search_type, keyword):
    search_options.append({
        'search_range': search_range,
        'search_type': search_type,
        'keyword': keyword
    })

def save_settings_and_close(folder_entry, options_list, window):
    save_settings({
        'folder_path': folder_entry.get(),
        'search_options': options_list
    })
    window.destroy()

def main():
    settings = load_settings()

    root = tk.Tk()
    root.title("Mail Search Settings")

    tk.Label(root, text="Folder Path:").grid(row=0, column=0, sticky='w')
    folder_entry = tk.Entry(root, width=50)
    folder_entry.grid(row=0, column=1, pady=5)
    folder_entry.insert(0, settings.get('folder_path', ''))

    folder_button = tk.Button(root, text="Browse", command=lambda: browse_folder(folder_entry))
    folder_button.grid(row=0, column=2)

    tk.Label(root, text="Search Options:").grid(row=1, column=0, sticky='w')
    options_frame = tk.Frame(root)
    options_frame.grid(row=1, column=1, columnspan=2, pady=5, sticky='w')

    options_list = settings.get('search_options', [])
    for i, option in enumerate(options_list):
        tk.Label(options_frame, text=f"Option {i + 1}:").grid(row=i, column=0, sticky='w')
        tk.Label(options_frame, text=f"Search Range: {option['search_range']}").grid(row=i, column=1, sticky='w')
        tk.Label(options_frame, text=f"Search Type: {option['search_type']}").grid(row=i, column=2, sticky='w')
        tk.Label(options_frame, text=f"Keyword: {option['keyword']}").grid(row=i, column=3, sticky='w')

    add_option_button = tk.Button(root, text="Add Search Option", command=lambda: add_search_option(options_list, 'Subject', 'Keyword', ''))
    add_option_button.grid(row=2, column=0, columnspan=3, pady=5)

    save_button = tk.Button(root, text="Save", command=lambda: save_settings_and_close(folder_entry, options_list, root))
    save_button.grid(row=3, column=0, columnspan=3, pady=10)

    root.mainloop()

    search_results = search_emails(settings['folder_path'], settings['search_options'])
    export_to_excel(search_results)

if __name__ == "__main__":
    main()
