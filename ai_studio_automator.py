
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, scrolledtext
import json
import os
import platform
import webbrowser
import time
import uuid
import re
import subprocess

# --- Library Imports ---
PYPDF2_AVAILABLE = False
TKINTERDND2_AVAILABLE = False
PYAUTOGUI_AVAILABLE = False
PYPERCLIP_AVAILABLE = False
PYWIN32_AVAILABLE = False

try:
    from PyPDF2 import PdfReader
    PYPDF2_AVAILABLE = True
except ImportError:
    print("Warning: 'PyPDF2' library not found. PDF features will be limited.")
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKINTERDND2_AVAILABLE = True
except ImportError:
    print("Warning: 'tkinterdnd2' library not found. Drag-and-drop feature will be unavailable.")
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    print("Warning: 'pyautogui' library not found. AI Studio automation will be unavailable.")
try:
    import pyperclip
    PYPERCLIP_AVAILABLE = True
except ImportError:
    print("Warning: 'pyperclip' library not found. Tkinter will be used for clipboard operations.")

if platform.system() == "Windows":
    try:
        import win32clipboard
        import win32con
        PYWIN32_AVAILABLE = True
        print("pywin32 library found. File copying feature on Windows is active.")
    except ImportError:
        print("Warning: 'pywin32' library not found. The 'copy as file' feature for TXT files on Windows will be unavailable.")
else:
    print("Running on a non-Windows system. The 'copy as file' feature is only applicable to Windows.")


DEFAULT_TEMPLATE_FILE = "file_processor_template_v9.7.json" # Version bump
APP_STATE_FILE = "file_processor_state_v9.7.json"      # Version bump
APP_TITLE = "File Processor and AI Studio Integration v9.7 (Prompt Writing via Clipboard)"

AI_STUDIO_URL = "https://aistudio.google.com/prompts/new_chat"
BROWSER_LOAD_DELAY = 5
PASTE_DELAY = 2 
FILE_UPLOAD_PROCESS_DELAY = 10
PROMPT_PASTE_DELAY = 1 # Delay after pasting the instructional prompt
SUBMIT_DELAY = 1
NEXT_TAB_DELAY = 3
NEXT_FILE_PROCESSING_DELAY = 4

CHAPTERS_PLACEHOLDER = "{CHAPTERS}"

ALL_FILES_ID = "__ALL_FILES__"
UNCATEGORIZED_ID = "__UNCATEGORIZED__"

def make_file_iid(path):
    return f"file_{path.replace(' ', '_').replace('/', '_').replace(':', '_').replace('.', '_')}"

def make_block_iid(file_path, block_id):
    return f"block_{make_file_iid(file_path)}_{block_id}"

def make_folder_iid(folder_id_uuid):
    return f"folder_{folder_id_uuid}"

def parse_complex_page_range_string(complex_range_str):
    complex_range_str = complex_range_str.strip()
    if not complex_range_str: return []
    all_pages = set()
    parts = complex_range_str.split(',')
    for part in parts:
        part = part.strip()
        if not part: continue
        if re.fullmatch(r"\d+", part):
            try:
                page = int(part)
                if page > 0: all_pages.add(page)
            except ValueError: pass
            continue
        match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)
        if match:
            try:
                start = int(match.group(1))
                end = int(match.group(2))
                if start > 0 and end > 0:
                    for page_num in range(min(start, end), max(start, end) + 1):
                        all_pages.add(page_num)
            except ValueError: pass
    return sorted(list(all_pages))

class FileProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1600x900")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabel", padding=5, font=('Arial', 10))
        style.configure("TButton", padding=5, font=('Arial', 10))
        style.configure("Header.TLabel", font=('Arial', 12, 'bold'))
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        style.configure("Treeview", rowheight=25)
        style.configure("Folder.Treeview.Heading", font=('Arial', 10, 'bold'))
        style.configure("Folder.Treeview", rowheight=25)

        self.file_items = []
        self.prompts = {}
        self._editing_item_iid = None
        self._edit_widget = None
        self._editing_field_name = None

        self.folders = []
        self.selected_folder_id = ALL_FILES_ID

        self.load_prompts()
        self.load_app_state()

        self.setup_ui()
        self.update_folder_treeview()
        self.update_file_treeview()

        if not PYAUTOGUI_AVAILABLE:
            messagebox.showwarning("Missing Library", "PyAutoGUI is required. AI Studio automation is disabled.")
        
        if platform.system() == "Windows" and not PYWIN32_AVAILABLE:
             messagebox.showwarning("Missing Library (Windows)",
                                   "'pywin32' library not found or could not be loaded.\n"
                                   "TXT files will be sent as text content instead of 'copy as file'.",
                                   parent=self.root)
        


        if TKINTERDND2_AVAILABLE:
            try:
                self.file_panel_frame.drop_target_register(DND_FILES)
                self.file_panel_frame.dnd_bind('<<Drop>>', self.handle_drop)
            except Exception as e: print(f"Could not set up drag-and-drop: {e}")
        else: print("tkinterdnd2 is not installed. Drag-and-drop is disabled.")

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        if self._edit_widget: self._commit_in_tree_edit()
        self.save_app_state()
        self.root.destroy()

    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.main_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.main_tab, text='Main Screen')
        self.create_main_tab_layout(self.main_tab)
        self.settings_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.settings_tab, text='Prompt Settings')
        self.create_settings_tab_content(self.settings_tab)
        self.notebook.pack(expand=True, fill='both')

    def create_main_tab_layout(self, tab):
        main_paned_window = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)
        folder_panel_frame = ttk.Frame(main_paned_window, padding=(0,0,5,0))
        main_paned_window.add(folder_panel_frame, weight=1)
        self.create_folder_panel_content(folder_panel_frame)
        self.file_panel_frame = ttk.Frame(main_paned_window)
        main_paned_window.add(self.file_panel_frame, weight=4)
        self.create_file_panel_content(self.file_panel_frame)

    def create_folder_panel_content(self, parent_frame):
        ttk.Label(parent_frame, text="Folders", style="Header.TLabel").pack(pady=(0,5), anchor="w")
        folder_controls_frame = ttk.Frame(parent_frame)
        folder_controls_frame.pack(fill=tk.X, pady=(0,5))
        add_folder_button = ttk.Button(folder_controls_frame, text="Add Folder", command=self.add_folder_dialog)
        add_folder_button.pack(side=tk.LEFT, padx=(0,5))
        self.rename_folder_button = ttk.Button(folder_controls_frame, text="Rename", command=self.rename_folder_dialog, state=tk.DISABLED)
        self.rename_folder_button.pack(side=tk.LEFT, padx=(0,5))
        self.delete_folder_button = ttk.Button(folder_controls_frame, text="Delete Folder", command=self.delete_selected_folder, state=tk.DISABLED)
        self.delete_folder_button.pack(side=tk.LEFT, padx=(0,5))
        folder_tree_frame = ttk.Frame(parent_frame)
        folder_tree_frame.pack(fill=tk.BOTH, expand=True)
        self.folder_tree = ttk.Treeview(folder_tree_frame, columns=("name",), show="tree headings", style="Folder.Treeview")
        self.folder_tree.heading("#0", text="Folder Name")
        self.folder_tree.column("#0", width=200, stretch=tk.YES)
        self.folder_tree.column("name", width=0, stretch=tk.NO, minwidth=0)
        folder_tree_scrollbar_y = ttk.Scrollbar(folder_tree_frame, orient="vertical", command=self.folder_tree.yview)
        self.folder_tree.configure(yscrollcommand=folder_tree_scrollbar_y.set)
        folder_tree_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_tree_selection_change)

    def create_file_panel_content(self, parent_frame):
        controls_frame = ttk.Frame(parent_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 10))
        add_button = ttk.Button(controls_frame, text="Add File", command=self.add_files_dialog)
        add_button.pack(side=tk.LEFT, padx=(0, 5))
        self.add_chapter_button = ttk.Button(controls_frame, text="Add Chapter", command=self.add_chapter_block_to_selected_file, state=tk.DISABLED)
        self.add_chapter_button.pack(side=tk.LEFT, padx=(0,5))
        self.remove_selected_button = ttk.Button(controls_frame, text="Delete Selected", command=self.remove_selected_tree_item, state=tk.DISABLED)
        self.remove_selected_button.pack(side=tk.LEFT, padx=(0,5))
        clear_button = ttk.Button(controls_frame, text="Clear Displayed Files", command=self.clear_displayed_files)
        clear_button.pack(side=tk.LEFT, padx=(0, 5))
        self.ai_studio_button1 = ttk.Button(controls_frame, text="AI Studio (Prompt 1)", command=lambda: self.perform_ai_studio_search_for_displayed_items('prompt1'))
        self.ai_studio_button1.pack(side=tk.LEFT, padx=(10, 0))
        self.ai_studio_button2 = ttk.Button(controls_frame, text="AI Studio (Prompt 2)", command=lambda: self.perform_ai_studio_search_for_displayed_items('prompt2'))
        self.ai_studio_button2.pack(side=tk.LEFT, padx=(5, 0))
        self.ai_studio_button3 = ttk.Button(controls_frame, text="AI Studio (Prompt 3)", command=lambda: self.perform_ai_studio_search_for_displayed_items('prompt3'))
        self.ai_studio_button3.pack(side=tk.LEFT, padx=(5, 0))
        self.full_book_all_button = ttk.Button(controls_frame, text="Full Book for All Displayed", command=self.process_full_book_for_all_displayed_files)
        self.full_book_all_button.pack(side=tk.LEFT, padx=(5,0))

        if not PYAUTOGUI_AVAILABLE:
            self.ai_studio_button1.config(state=tk.DISABLED)
            self.ai_studio_button2.config(state=tk.DISABLED)
            self.ai_studio_button3.config(state=tk.DISABLED)
            self.full_book_all_button.config(state=tk.DISABLED)

        tree_frame = ttk.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.file_tree = ttk.Treeview(tree_frame, columns=("type", "details", "page_range", "total_chapters"), show="tree headings")
        self.file_tree.heading("#0", text="File / Chapter Block")
        self.file_tree.heading("type", text="Type")
        self.file_tree.heading("details", text="Chapter Text / File Path")
        self.file_tree.heading("page_range", text="Page Range (PDF) / N/A (TXT)")
        self.file_tree.heading("total_chapters", text="Total Chapters (FullBook)")
        self.file_tree.column("#0", width=250, stretch=tk.YES)
        self.file_tree.column("type", width=80, anchor="center")
        self.file_tree.column("details", width=300)
        self.file_tree.column("page_range", width=180, anchor="w")
        self.file_tree.column("total_chapters", width=180, anchor="center")
        tree_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview)
        tree_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)
        tree_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.file_tree.bind("<Double-1>", self.on_tree_item_double_click)
        self.file_tree.bind("<Delete>", lambda e: self.remove_selected_tree_item())
        self.file_tree.bind("<<TreeviewSelect>>", self.on_file_tree_selection_change)
        self.file_tree.bind("<Button-3>", self.show_context_menu)

    def add_folder_dialog(self):
        folder_name = simpledialog.askstring("New Folder", "Folder Name:", parent=self.root)
        if folder_name:
            folder_name = folder_name.strip()
            if not folder_name: messagebox.showwarning("Invalid Name", "Folder name cannot be empty."); return
            if any(f['name'] == folder_name for f in self.folders): messagebox.showwarning("Existing Folder", f"A folder named '{folder_name}' already exists."); return
            new_folder_id = uuid.uuid4().hex
            self.folders.append({'id': new_folder_id, 'name': folder_name})
            self.folders.sort(key=lambda f: f['name'].lower())
            self.update_folder_treeview()
            new_folder_iid = make_folder_iid(new_folder_id)
            if self.folder_tree.exists(new_folder_iid):
                self.folder_tree.selection_set(new_folder_iid); self.folder_tree.focus(new_folder_iid); self.folder_tree.see(new_folder_iid)
            self.save_app_state()

    def rename_folder_dialog(self):
        selected_folder_iids = self.folder_tree.selection()
        if not selected_folder_iids: return
        selected_folder_iid = selected_folder_iids[0]
        if selected_folder_iid == ALL_FILES_ID or selected_folder_iid == UNCATEGORIZED_ID: messagebox.showinfo("Info", "This special view cannot be renamed."); return
        folder_to_rename = next((f for f in self.folders if make_folder_iid(f['id']) == selected_folder_iid), None)
        if not folder_to_rename: return
        new_name = simpledialog.askstring("Rename Folder", "New Folder Name:", initialvalue=folder_to_rename['name'], parent=self.root)
        if new_name:
            new_name = new_name.strip()
            if not new_name: messagebox.showwarning("Invalid Name", "Folder name cannot be empty."); return
            if new_name != folder_to_rename['name'] and any(f['name'] == new_name for f in self.folders): messagebox.showwarning("Existing Folder", f"A folder named '{new_name}' already exists."); return
            folder_to_rename['name'] = new_name
            self.folders.sort(key=lambda f: f['name'].lower())
            self.update_folder_treeview()
            if self.folder_tree.exists(selected_folder_iid): self.folder_tree.selection_set(selected_folder_iid); self.folder_tree.focus(selected_folder_iid)
            self.save_app_state()

    def delete_selected_folder(self):
        selected_folder_iids = self.folder_tree.selection()
        if not selected_folder_iids: return
        selected_folder_iid = selected_folder_iids[0]
        if selected_folder_iid == ALL_FILES_ID or selected_folder_iid == UNCATEGORIZED_ID: messagebox.showinfo("Info", "This special view cannot be deleted."); return
        folder_to_delete = next((f for f in self.folders if make_folder_iid(f['id']) == selected_folder_iid), None)
        if not folder_to_delete: return
        if messagebox.askyesno("Delete Folder", f"Are you sure you want to delete the folder '{folder_to_delete['name']}'?\n(Files inside will be moved to 'Uncategorized')", parent=self.root):
            folder_id_to_remove = folder_to_delete['id']
            self.folders = [f for f in self.folders if f['id'] != folder_id_to_remove]
            for file_item in self.file_items:
                if file_item.get('folder_id') == folder_id_to_remove: file_item['folder_id'] = None 
            if self.selected_folder_id == folder_id_to_remove: self.selected_folder_id = ALL_FILES_ID
            self.update_folder_treeview(); self.update_file_treeview(); self.save_app_state()

    def update_folder_treeview(self):
        current_selection = self.folder_tree.selection()
        for item_id in self.folder_tree.get_children(""): self.folder_tree.delete(item_id)
        self.folder_tree.insert("", tk.END, iid=ALL_FILES_ID, text="All Files", values=("All Files",))
        self.folder_tree.insert("", tk.END, iid=UNCATEGORIZED_ID, text="Uncategorized", values=("Uncategorized",))
        for folder in sorted(self.folders, key=lambda f: f['name'].lower()):
            folder_iid = make_folder_iid(folder['id'])
            self.folder_tree.insert("", tk.END, iid=folder_iid, text=folder['name'], values=(folder['name'],))
        sel_id_to_restore = self.selected_folder_id
        if not self.folder_tree.exists(sel_id_to_restore):
            if current_selection and self.folder_tree.exists(current_selection[0]): sel_id_to_restore = current_selection[0]
            else: sel_id_to_restore = ALL_FILES_ID
        if self.folder_tree.exists(sel_id_to_restore):
            self.folder_tree.selection_set(sel_id_to_restore); self.folder_tree.focus(sel_id_to_restore); self.folder_tree.see(sel_id_to_restore)
        self.on_folder_tree_selection_change() # Ensure state is updated

    def on_folder_tree_selection_change(self, event=None):
        selected_iids = self.folder_tree.selection()
        if selected_iids:
            selected_iid = selected_iids[0]
            if selected_iid == ALL_FILES_ID or selected_iid == UNCATEGORIZED_ID:
                self.selected_folder_id = selected_iid
                self.rename_folder_button.config(state=tk.DISABLED); self.delete_folder_button.config(state=tk.DISABLED)
            else:
                folder_obj = next((f for f in self.folders if make_folder_iid(f['id']) == selected_iid), None)
                if folder_obj: self.selected_folder_id = folder_obj['id']; self.rename_folder_button.config(state=tk.NORMAL); self.delete_folder_button.config(state=tk.NORMAL)
                else: self.selected_folder_id = ALL_FILES_ID; self.rename_folder_button.config(state=tk.DISABLED); self.delete_folder_button.config(state=tk.DISABLED)
            self.update_file_treeview()
        else: self.selected_folder_id = ALL_FILES_ID; self.rename_folder_button.config(state=tk.DISABLED); self.delete_folder_button.config(state=tk.DISABLED); self.update_file_treeview()

    def get_displayed_file_items(self):
        if self.selected_folder_id == ALL_FILES_ID: return list(self.file_items)
        elif self.selected_folder_id == UNCATEGORIZED_ID: return [item for item in self.file_items if item.get('folder_id') is None]
        elif self.selected_folder_id: return [item for item in self.file_items if item.get('folder_id') == self.selected_folder_id]
        return []

    def on_file_tree_selection_change(self, event=None):
        selected_items = self.file_tree.selection()
        if selected_items:
            self.remove_selected_button.config(state=tk.NORMAL)
            first_selected_iid = selected_items[0]
            if self.file_tree.exists(first_selected_iid):
                tags = self.file_tree.item(first_selected_iid, "tags")
                if "file_item" in tags or "chapter_block_item" in tags: self.add_chapter_button.config(state=tk.NORMAL)
                else: self.add_chapter_button.config(state=tk.DISABLED)
        else: self.remove_selected_button.config(state=tk.DISABLED); self.add_chapter_button.config(state=tk.DISABLED)

    def get_file_item_by_path(self, path):
        for item in self.file_items:
            if item['path'] == path: return item
        return None

    def get_chapter_block_from_iid(self, block_iid):
        if not block_iid or not block_iid.startswith("block_"): return None, None
        parent_iid = self.file_tree.parent(block_iid)
        if not parent_iid: return None, None
        file_data = self.get_file_data_from_tree_iid(parent_iid)
        if not file_data: return None, None
        try:
            block_id_from_iid = block_iid.split('_')[-1]
            for block in file_data['chapter_blocks']:
                if block['id'] == block_id_from_iid: return block, file_data
        except Exception as e: print(f"Error parsing block IID: {block_iid}, {e}")
        return None, None

    def on_tree_item_double_click(self, event):
        if self._edit_widget: self._commit_in_tree_edit()
        item_iid = self.file_tree.identify_row(event.y); column_id_str = self.file_tree.identify_column(event.x)
        if not item_iid or not column_id_str: return
        item_tags = self.file_tree.item(item_iid, "tags")
        if "file_item" in item_tags and column_id_str == "#0":
            file_data = self.get_file_data_from_tree_iid(item_iid)
            if file_data and file_data.get('path'):
                try:
                    filepath = os.path.abspath(file_data['path'])
                    if not os.path.exists(filepath): messagebox.showerror("Error", f"File not found: {filepath}"); return
                    if platform.system() == "Windows": os.startfile(filepath)
                    elif platform.system() == "Darwin": subprocess.call(('open', filepath))
                    else: subprocess.call(('xdg-open', filepath))
                    return
                except Exception as e: messagebox.showerror("Error", f"Could not open file: {file_data['filename']}\n{e}"); return
        column_heading = self.file_tree.heading(column_id_str)['text'] if column_id_str != "#0" else "#0"
        field_to_edit = None; current_value_for_edit = ""
        if "chapter_block_item" in item_tags:
            block_to_edit, parent_file_data = self.get_chapter_block_from_iid(item_iid)
            if not block_to_edit: return
            if column_heading == "File / Chapter Block" or column_heading == "Chapter Text / File Path": field_to_edit = "text"
            elif column_heading == "Page Range (PDF) / N/A (TXT)" and parent_file_data.get('type') == 'pdf': field_to_edit = "page_ranges_str"
            if field_to_edit: current_value_for_edit = block_to_edit.get(field_to_edit, "")
        elif "file_item" in item_tags:
            file_to_edit = self.get_file_data_from_tree_iid(item_iid)
            if not file_to_edit: return
            if column_heading == "Total Chapters (FullBook)": field_to_edit = "total_chapters_for_full_book"; current_value_for_edit = str(file_to_edit.get(field_to_edit, 0))
        if not field_to_edit: return
        x, y, width, height = self.file_tree.bbox(item_iid, column=column_id_str)
        self._editing_item_iid = item_iid; self._editing_field_name = field_to_edit
        self._edit_widget = ttk.Entry(self.file_tree); self._edit_widget.place(x=x, y=y, width=width, height=height)
        self._edit_widget.insert(0, current_value_for_edit); self._edit_widget.select_range(0, tk.END); self._edit_widget.focus_set()
        self._edit_widget.bind("<Return>", self._commit_in_tree_edit); self._edit_widget.bind("<FocusOut>", self._commit_in_tree_edit); self._edit_widget.bind("<Escape>", self._cancel_in_tree_edit)

    def _commit_in_tree_edit(self, event=None):
        if not self._edit_widget or not self._editing_item_iid or not self._editing_field_name:
            if self._edit_widget: self._edit_widget.destroy()
            self._edit_widget = None; return
        new_value_str = self._edit_widget.get(); item_iid = self._editing_item_iid; field_name = self._editing_field_name
        tags = self.file_tree.item(item_iid, "tags"); commit_successful = False
        if "file_item" in tags and field_name == "total_chapters_for_full_book":
            file_data = self.get_file_data_from_tree_iid(item_iid)
            if file_data:
                try:
                    new_val_int = int(new_value_str)
                    if new_val_int < 0: messagebox.showwarning("Invalid Value", "Total chapters must be >= 0."); self._edit_widget.focus_set(); return
                    file_data[field_name] = new_val_int; self.file_tree.set(item_iid, "total_chapters", str(new_val_int)); commit_successful = True
                except ValueError: messagebox.showerror("Invalid Value", "Please enter an integer."); self._edit_widget.focus_set(); return
        elif "chapter_block_item" in tags:
            block_data, _ = self.get_chapter_block_from_iid(item_iid)
            if block_data:
                block_data[field_name] = new_value_str
                if field_name == "text": display_text = f"Chapter: {new_value_str}" if new_value_str else "Chapter: [Empty]"; self.file_tree.item(item_iid, text=display_text); self.file_tree.set(item_iid, "details", new_value_str)
                elif field_name == "page_ranges_str": self.file_tree.set(item_iid, "page_range", new_value_str)
                commit_successful = True
        if commit_successful: self.save_app_state()
        self._edit_widget.destroy(); self._edit_widget = None; self._editing_item_iid = None; self._editing_field_name = None

    def _cancel_in_tree_edit(self, event=None):
        if self._edit_widget: self._edit_widget.destroy()
        self._edit_widget = None; self._editing_item_iid = None; self._editing_field_name = None

    def get_file_data_from_tree_iid(self, file_tree_iid):
        for file_data in self.file_items:
            if make_file_iid(file_data['path']) == file_tree_iid: return file_data
        return None

    def add_files_dialog(self):
        filetypes = (("PDF files", "*.pdf"), ("Text files", "*.txt"), ("All files", "*.*"))
        filepaths = filedialog.askopenfilenames(title="Select Files", filetypes=filetypes)
        if filepaths:
            added_count = 0
            for path in filepaths:
                if path.lower().endswith(".pdf") and not PYPDF2_AVAILABLE: messagebox.showerror("Error", f"PyPDF2 is not installed. '{os.path.basename(path)}' could not be added."); continue
                if not (path.lower().endswith(".pdf") or path.lower().endswith(".txt")): messagebox.showwarning("Unsupported File", f"'{os.path.basename(path)}' is not supported."); continue
                if self.add_file_to_list(path): added_count += 1
            if added_count > 0: self.update_file_treeview(); self.save_app_state()

    def handle_drop(self, event):
        raw_paths = event.data; paths = []
        if '{' in raw_paths and '}' in raw_paths: paths = [p.strip('{}') for p in re.findall(r'\{[^{}]*\}|[^{}\s]+', raw_paths)]
        else: paths = raw_paths.split()
        added_count = 0
        for p in paths:
            p_cleaned = p.strip().strip('"').strip("'")
            if not p_cleaned: continue
            if p_cleaned.lower().endswith(".pdf"):
                if not PYPDF2_AVAILABLE: messagebox.showerror("Error", f"PyPDF2 is not installed. '{os.path.basename(p_cleaned)}' could not be added."); continue
                if self.add_file_to_list(p_cleaned): added_count +=1
            elif p_cleaned.lower().endswith(".txt"):
                if self.add_file_to_list(p_cleaned): added_count +=1
            else: print(f"Dragged file is not PDF/TXT, skipping: {p_cleaned}")
        if added_count > 0: self.update_file_treeview(); self.save_app_state()

    def add_file_to_list(self, filepath):
        filename = os.path.basename(filepath)
        if self.get_file_item_by_path(filepath): messagebox.showinfo("Info", f"'{filename}' is already in the list."); return False
        file_type = "pdf" if filepath.lower().endswith(".pdf") else "txt"
        current_folder_id = None
        if self.selected_folder_id and self.selected_folder_id != ALL_FILES_ID and self.selected_folder_id != UNCATEGORIZED_ID: current_folder_id = self.selected_folder_id
        new_file_item = {'path': filepath, 'filename': filename, 'type': file_type, 'chapter_blocks': [{'id': uuid.uuid4().hex, 'text': '', 'page_ranges_str': ""}], 'total_chapters_for_full_book': 0, 'folder_id': current_folder_id}
        self.file_items.append(new_file_item)
        return True

    def add_chapter_block_to_selected_file(self):
        selected_iids = self.file_tree.selection()
        if not selected_iids: messagebox.showinfo("Info", "Select a file or chapter."); return
        target_file_iid = selected_iids[0]
        if "chapter_block_item" in self.file_tree.item(target_file_iid, "tags"): target_file_iid = self.file_tree.parent(target_file_iid)
        if not target_file_iid or "file_item" not in self.file_tree.item(target_file_iid, "tags"): messagebox.showinfo("Info", "Select a file or a chapter belonging to a file."); return
        file_data = self.get_file_data_from_tree_iid(target_file_iid)
        if file_data:
            new_block_id = uuid.uuid4().hex
            file_data['chapter_blocks'].append({'id': new_block_id, 'text': '', 'page_ranges_str': ""})
            open_states = {iid: self.file_tree.item(iid, "open") for iid in self.file_tree.get_children("") if self.file_tree.exists(iid) and self.file_tree.item(iid, "open")}
            self.update_file_treeview(open_states_to_restore=open_states)
            new_block_tree_iid = make_block_iid(file_data['path'], new_block_id)
            if self.file_tree.exists(new_block_tree_iid): self.file_tree.selection_set(new_block_tree_iid); self.file_tree.focus(new_block_tree_iid); self.file_tree.see(new_block_tree_iid)
            self.save_app_state()

    def remove_selected_tree_item(self):
        selected_iids = self.file_tree.selection()
        if not selected_iids: messagebox.showinfo("Info", "Select an item to delete."); return
        if not messagebox.askyesno("Confirm", f"Delete {len(selected_iids)} item(s)?"): return
        items_to_delete = []
        for item_iid in selected_iids:
            if not self.file_tree.exists(item_iid): continue
            tags = self.file_tree.item(item_iid, "tags")
            if "file_item" in tags:
                file_data = self.get_file_data_from_tree_iid(item_iid)
                if file_data: items_to_delete.append({'type': 'file', 'path': file_data['path']})
            elif "chapter_block_item" in tags:
                block_data, parent_data = self.get_chapter_block_from_iid(item_iid)
                if block_data and parent_data: items_to_delete.append({'type': 'block', 'file_path': parent_data['path'], 'block_id': block_data['id']})
        new_list = []
        for file_model in self.file_items:
            if any(d['type'] == 'file' and d['path'] == file_model['path'] for d in items_to_delete): continue
            file_model['chapter_blocks'] = [b for b in file_model['chapter_blocks'] if not any(d['type'] == 'block' and d['file_path'] == file_model['path'] and d['block_id'] == b['id'] for d in items_to_delete)]
            new_list.append(file_model)
        self.file_items = new_list
        open_states = {iid: self.file_tree.item(iid, "open") for iid in self.file_tree.get_children("") if self.file_tree.exists(iid) and self.file_tree.item(iid, "open")}
        self.update_file_treeview(open_states_to_restore=open_states)
        self.save_app_state(); self.on_file_tree_selection_change()

    def update_file_treeview(self, open_states_to_restore=None):
        if self._edit_widget: self._commit_in_tree_edit()
        if open_states_to_restore is None: open_states_to_restore = {iid: self.file_tree.item(iid, "open") for iid in self.file_tree.get_children("") if self.file_tree.exists(iid) and self.file_tree.item(iid, "open")} if hasattr(self, 'file_tree') and self.file_tree.winfo_exists() else {}
        selection = self.file_tree.selection() if hasattr(self, 'file_tree') and self.file_tree.winfo_exists() else []
        yview_pos = self.file_tree.yview() if hasattr(self, 'file_tree') and self.file_tree.winfo_exists() else (0.0, 1.0)
        for item_id in self.file_tree.get_children(""): self.file_tree.delete(item_id)
        for file_item in self.get_displayed_file_items():
            file_iid = make_file_iid(file_item['path'])
            total_ch = file_item.get('total_chapters_for_full_book', 0); file_type_disp = file_item.get('type', 'N/A').upper()
            file_node = self.file_tree.insert("", tk.END, iid=file_iid, text=file_item['filename'], values=(file_type_disp, file_item['path'], "", str(total_ch)), tags=("file_item",))
            if file_iid in open_states_to_restore and open_states_to_restore[file_iid]: self.file_tree.item(file_node, open=True)
            for block in file_item['chapter_blocks']:
                block_iid = make_block_iid(file_item['path'], block['id'])
                block_text_disp = f"Chapter: {block['text']}" if block['text'] else "Chapter: [Empty]"
                page_range_disp = block.get('page_ranges_str', "") if file_item.get('type') == 'pdf' else "N/A"
                self.file_tree.insert(file_node, tk.END, iid=block_iid, text=block_text_disp, values=("Chapter Block", block['text'], page_range_disp, ""), tags=("chapter_block_item",))
        valid_selection = [s for s in selection if self.file_tree.exists(s)]
        if valid_selection:
            self.file_tree.selection_set(valid_selection)
            if valid_selection[0]: self.file_tree.focus(valid_selection[0]); self.file_tree.see(valid_selection[0])
        self.root.update_idletasks()
        if yview_pos and len(yview_pos) == 2: self.file_tree.yview_moveto(yview_pos[0])
        self.on_file_tree_selection_change()

    def clear_displayed_files(self):
        current_folder_id = self.selected_folder_id; display_name = ""; confirm_msg = ""
        if current_folder_id == ALL_FILES_ID: display_name = "All Files"; confirm_msg = "Delete all files in the list?"
        elif current_folder_id == UNCATEGORIZED_ID: display_name = "the Uncategorized section"; confirm_msg = f"Delete all files in {display_name}?"
        else:
            folder_obj = next((f for f in self.folders if f['id'] == current_folder_id), None)
            if folder_obj: display_name = f"the '{folder_obj['name']}' folder"; confirm_msg = f"Delete all files in {display_name}?"
            else: messagebox.showerror("Error", "Could not get selected folder information."); return
        files_to_clear = self.get_displayed_file_items()
        if not files_to_clear: messagebox.showinfo("Info", f"{display_name} is already empty."); return
        if messagebox.askyesno("Confirm", confirm_msg, parent=self.root):
            paths_to_remove = {item['path'] for item in files_to_clear}
            self.file_items = [item for item in self.file_items if item['path'] not in paths_to_remove]
            self.update_file_treeview(open_states_to_restore={}); self.save_app_state(); self.on_file_tree_selection_change()

    def extract_text_from_file(self, filepath, list_of_pages_to_extract, file_type):
        if file_type == 'txt':
            try:
                encodings_to_try = ['utf-8', 'latin-1', 'windows-1252']
                for enc in encodings_to_try:
                    try:
                        with open(filepath, 'r', encoding=enc) as f: return f.read()
                    except UnicodeDecodeError:
                        if enc == encodings_to_try[-1]: raise
                return None 
            except Exception as e: messagebox.showerror("TXT Read Error",f"Error reading '{os.path.basename(filepath)}': {e}"); return None
        elif file_type == 'pdf':
            if not PYPDF2_AVAILABLE: return None
            try:
                reader = PdfReader(filepath)
                if not list_of_pages_to_extract:
                    full_text = "".join(page.extract_text() + "\n" for page in reader.pages if page.extract_text())
                    return full_text.strip()
                else:
                    num_pages_total = len(reader.pages); extracted_text_parts = []
                    for page_num_1_indexed in list_of_pages_to_extract:
                        page_idx_0_indexed = page_num_1_indexed - 1
                        if 0 <= page_idx_0_indexed < num_pages_total:
                            page_text = reader.pages[page_idx_0_indexed].extract_text()
                            if page_text: extracted_text_parts.append(page_text)
                        else: print(f"Warning: Page {page_num_1_indexed} is invalid for '{os.path.basename(filepath)}'. Skipping.")
                    return "\n".join(extracted_text_parts).strip()
            except Exception as e:
                page_list_str = ", ".join(map(str, list_of_pages_to_extract)) if list_of_pages_to_extract else "all"
                messagebox.showerror("PDF Read Error",f"Error reading '{os.path.basename(filepath)}' (pages: {page_list_str}): {e}"); return None
        return None

    def create_settings_tab_content(self, tab):
        settings_frame = ttk.Frame(tab); settings_frame.pack(expand=True, fill='both', padx=10, pady=10)
        ttk.Label(settings_frame, text=(
                    f"Use '{CHAPTERS_PLACEHOLDER}' in your prompts to specify chapter/section names.\n"
                    "These placeholders will generate the commands to be written in AI Studio's text box.\n"
                    "The file content (PDF text or TXT file on Windows) is sent/attached separately."),
                  wraplength=700).pack(anchor="w", pady=(0,10))
        self.prompt_widgets = {}
        prompt_labels = {
            "prompt1": "Prompt 1 Template:", "prompt2": "Prompt 2 Template:",
            "prompt3": "Prompt 3 Template:", "full_book_prompt": "Full Book Prompt Template:"
        }
        for key, label_text in prompt_labels.items():
            ttk.Label(settings_frame, text=label_text, style="Header.TLabel").pack(anchor="w", pady=(10,2))
            text_widget = scrolledtext.ScrolledText(settings_frame, height=3, width=80, relief=tk.SOLID, borderwidth=1)
            text_widget.pack(fill=tk.BOTH, expand=True, pady=(0,5))
            text_widget.insert(tk.END, self.prompts.get(key, ""))
            self.prompt_widgets[key] = text_widget
        save_button = ttk.Button(settings_frame, text="Save All Prompts", command=self.save_prompts)
        save_button.pack(side=tk.LEFT, padx=(0,5), pady=(10,0))

    def load_prompts(self):
        defaults = {
            "prompt1": f"Please summarize the '{CHAPTERS_PLACEHOLDER}' section from the attached file.",
            "prompt2": f"Analyze the key points of the '{CHAPTERS_PLACEHOLDER}' section in the attached file.",
            "prompt3": f"Extract actionable items from the '{CHAPTERS_PLACEHOLDER}' section in the attached file.",
            "full_book_prompt": f"Please provide a comprehensive summary for the entire attached file, considering the following sections: {CHAPTERS_PLACEHOLDER}."
        }
        if os.path.exists(DEFAULT_TEMPLATE_FILE):
            try:
                with open(DEFAULT_TEMPLATE_FILE, 'r', encoding='utf-8') as f: loaded_prompts = json.load(f)
                for key in defaults: self.prompts[key] = loaded_prompts.get(key, defaults[key])
            except Exception as e: print(f"Error loading prompt file: {e}. Using defaults."); self.prompts = defaults.copy()
        else: self.prompts = defaults.copy()
        if hasattr(self, 'prompt_widgets'):
            for key, widget in self.prompt_widgets.items():
                widget.delete("1.0", tk.END); widget.insert(tk.END, self.prompts.get(key, ""))

    def save_prompts(self):
        if hasattr(self, 'prompt_widgets'):
            for key, widget in self.prompt_widgets.items(): self.prompts[key] = widget.get("1.0", tk.END).strip()
        try:
            with open(DEFAULT_TEMPLATE_FILE, 'w', encoding='utf-8') as f: json.dump(self.prompts, f, indent=4)
            messagebox.showinfo("Success", "Prompt templates saved.")
        except IOError: messagebox.showerror("Error", "Could not save prompt templates.")

    def save_app_state(self):
        state = {"file_items": self.file_items, "folders": self.folders, "selected_folder_id": self.selected_folder_id}
        try:
            with open(APP_STATE_FILE, 'w', encoding='utf-8') as f: json.dump(state, f, indent=4)
        except IOError as e: print(f"Could not save state: {e}")

    def load_app_state(self):
        if os.path.exists(APP_STATE_FILE):
            try:
                with open(APP_STATE_FILE, 'r', encoding='utf-8') as f: state = json.load(f)
                self.folders = state.get("folders", [])
                self.selected_folder_id = state.get("selected_folder_id", ALL_FILES_ID)
                loaded_items = state.get("file_items", state.get("pdf_items", []))
                for item in loaded_items:
                    if 'type' not in item: item['type'] = 'pdf' if item['path'].lower().endswith('.pdf') else 'txt'
                    if 'chapter_blocks' not in item: item['chapter_blocks'] = []
                    if 'total_chapters_for_full_book' not in item: item['total_chapters_for_full_book'] = 0
                    if 'folder_id' not in item: item['folder_id'] = None
                    for block in item['chapter_blocks']:
                        if 'id' not in block: block['id'] = uuid.uuid4().hex
                        if 'page_ranges_str' not in block:
                            page_s = block.pop('page_start',0); page_e = block.pop('page_end',0)
                            block['page_ranges_str'] = f"{page_s}-{page_e}" if page_s > 0 and page_e > 0 and page_s != page_e else (str(page_s) if page_s > 0 else "")
                self.file_items = [item for item in loaded_items if item['type'] in ['pdf', 'txt']]
            except Exception as e: print(f"Error loading state: {e}. Using defaults."); self.file_items = []; self.folders = []; self.selected_folder_id = ALL_FILES_ID
        else: self.file_items = []; self.folders = []; self.selected_folder_id = ALL_FILES_ID
        
    def _copy_file_to_clipboard_windows(self, file_path):
        if not PYWIN32_AVAILABLE: return False
        try:
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path): print(f"File not found: {abs_path}"); return False
            ps_command = f"Set-Clipboard -Path '{abs_path}'"
            process = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command], 
                                     capture_output=True, text=True, check=False, creationflags=subprocess.CREATE_NO_WINDOW)
            if process.returncode == 0: print(f"'{abs_path}' copied to clipboard as a file object via PowerShell."); return True
            else:
                print(f"PowerShell file copy error: {process.stderr}")
                messagebox.showwarning("File Copy Error", f"Could not copy '{os.path.basename(abs_path)}' to the clipboard as a 'file object' (PowerShell).", parent=self.root)
                return False
        except Exception as e: print(f"Windows file copy error: {e}"); messagebox.showerror("File Copy Error", f"Could not copy file to clipboard: {e}"); return False

    def _execute_ai_studio_automation(self, data_for_clipboard, is_file_object_on_clipboard, prompt_text_to_paste_after_file, item_description):
        if not PYAUTOGUI_AVAILABLE: messagebox.showerror("Error", "PyAutoGUI is not installed."); return False
        print(f"AI Studio automation for '{item_description}'.")

        # Step 1: Set clipboard (if not already set for file object)
        if not is_file_object_on_clipboard: # data_for_clipboard is text (PDF content + prompt, or non-Win TXT content + prompt)
            try:
                if PYPERCLIP_AVAILABLE: pyperclip.copy(data_for_clipboard)
                else: self.root.clipboard_clear(); self.root.clipboard_append(data_for_clipboard); self.root.update()
                print(f"Copied to clipboard (length: {len(data_for_clipboard)}): '{data_for_clipboard[:100]}...'")
            except Exception as e:
                messagebox.showerror("Clipboard Error", f"Could not copy text to clipboard: {e}"); return False
        # If is_file_object_on_clipboard is True, clipboard was already set by _copy_file_to_clipboard_windows

        # Step 2: Open browser and paste
        print(f"Opening AI Studio ({AI_STUDIO_URL})..."); webbrowser.open_new_tab(AI_STUDIO_URL)
        print(f"Waiting {BROWSER_LOAD_DELAY}s for browser to load..."); time.sleep(BROWSER_LOAD_DELAY)

        try:
            print("Pasting clipboard content (Ctrl+V)..."); pyautogui.hotkey('ctrl', 'v')
            print(f"Waiting {PASTE_DELAY}s after paste..."); time.sleep(PASTE_DELAY)

            if is_file_object_on_clipboard: # Only for Windows TXT file copy
                print(f"Waiting {FILE_UPLOAD_PROCESS_DELAY}s for file upload process..."); time.sleep(FILE_UPLOAD_PROCESS_DELAY)
                if prompt_text_to_paste_after_file:
                    print(f"Copying instructional prompt to clipboard: '{prompt_text_to_paste_after_file}'")
                    try: # Copy instructional prompt to clipboard
                        if PYPERCLIP_AVAILABLE: pyperclip.copy(prompt_text_to_paste_after_file)
                        else: self.root.clipboard_clear(); self.root.clipboard_append(prompt_text_to_paste_after_file); self.root.update()
                    except Exception as e: messagebox.showerror("Clipboard Error", f"Could not copy instructional prompt to clipboard: {e}"); return False
                    
                    print("Pasting instructional prompt (Ctrl+V)..."); pyautogui.hotkey('ctrl', 'v')
                    print(f"Waiting {PROMPT_PASTE_DELAY}s after pasting instructional prompt..."); time.sleep(PROMPT_PASTE_DELAY)
            
            print("Sending command (Ctrl+Enter)..."); pyautogui.hotkey('ctrl', 'enter')
            print(f"Waiting {SUBMIT_DELAY}s after submit..."); time.sleep(SUBMIT_DELAY)
            print(f"Automation for '{item_description}' completed."); return True
        except Exception as e:
            messagebox.showerror("Automation Error", f"PyAutoGUI error ('{item_description}'): {e}\nOperation stopped."); return False

    def _prepare_instructional_prompt(self, template, chapters_text):
        prompt = template
        if CHAPTERS_PLACEHOLDER in prompt:
            prompt = prompt.replace(CHAPTERS_PLACEHOLDER, chapters_text)
        return prompt

    def _get_chapters_text_for_template(self, file_item, chapter_block=None, chapter_indices_for_full_book=None):
        if chapter_block: # Single chapter mode
            # Use the exact text from the chapter block if it exists, otherwise "Unspecified Chapter"
            # This ensures user's exact input is used.
            return chapter_block['text'].strip() if chapter_block.get('text','').strip() else "Unspecified Chapter"
        
        if chapter_indices_for_full_book is not None: # Full book chunk mode
            target_chapter_names = []
            all_defined_blocks = file_item.get('chapter_blocks', [])
            for conceptual_idx_0_based in chapter_indices_for_full_book:
                # Use defined chapter name if available for that conceptual index
                if conceptual_idx_0_based < len(all_defined_blocks) and all_defined_blocks[conceptual_idx_0_based].get('text','').strip():
                    target_chapter_names.append(all_defined_blocks[conceptual_idx_0_based]['text'])
                else:
                    # If no defined name, use the 1-indexed number directly
                    target_chapter_names.append(str(conceptual_idx_0_based + 1)) 
            return ", ".join(target_chapter_names) if target_chapter_names else "Specified Chapters"

        return "Entire File" # Default for entire_file_context

    def perform_ai_studio_search_for_displayed_items(self, prompt_key):
        if not PYAUTOGUI_AVAILABLE: messagebox.showerror("Error", "PyAutoGUI is not installed."); return
        current_template = self.prompts.get(prompt_key)
        if not current_template: messagebox.showerror("Error", f"Prompt template for '{prompt_key}' not found."); return
        
        tasks = []
        for item in self.get_displayed_file_items():
            if not item.get('chapter_blocks'): continue
            for block in item['chapter_blocks']: tasks.append({'file_item': item, 'chapter_block': block})

        if not tasks: messagebox.showinfo("Info", "No chapters to process."); return
        if not messagebox.askyesno("Confirm", f"{len(tasks)} chapters will be processed in AI Studio with '{prompt_key}'.\nContinue?"): return

        self.root.config(cursor="watch"); self.root.update_idletasks()
        processed_count = 0
        for task in tasks:
            file_item = task['file_item']; chapter_block = task['chapter_block']
            file_path = file_item['path']; filename = file_item['filename']; file_type = file_item['type']
            page_str = chapter_block.get('page_ranges_str', "")
            
            chapters_for_template = self._get_chapters_text_for_template(file_item, chapter_block=chapter_block)
            item_description = f"{filename} - {chapters_for_template}"
            if file_type == 'pdf' and page_str: item_description += f" (Pages: {page_str})"
            instructional_prompt_text = self._prepare_instructional_prompt(current_template, chapters_for_template)

            data_for_clipboard = ""; is_file_object = False; prompt_to_paste_after = instructional_prompt_text

            if file_type == 'txt' and platform.system() == "Windows" and PYWIN32_AVAILABLE:
                if self._copy_file_to_clipboard_windows(file_path): data_for_clipboard = file_path; is_file_object = True
                else: 
                    extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
                    data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"
                    prompt_to_paste_after = "" 
            elif file_type == 'pdf':
                list_of_pages = parse_complex_page_range_string(page_str) if page_str else []
                extracted_text = self.extract_text_from_file(file_path, list_of_pages, file_type) or "[NO PDF CONTENT]"
                data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"
                prompt_to_paste_after = ""
            else: # TXT on non-Win or no pywin32
                extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
                data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"
                prompt_to_paste_after = ""

            if not self._execute_ai_studio_automation(data_for_clipboard, is_file_object, prompt_to_paste_after, item_description):
                self.root.config(cursor=""); return
            
            processed_count += 1
            if processed_count < len(tasks): time.sleep(NEXT_TAB_DELAY)
        
        self.root.config(cursor="")
        if processed_count > 0: messagebox.showinfo("Completed", f"AI Studio process initiated for {processed_count} chapters.");
        elif tasks: messagebox.showinfo("Info", "An issue occurred while processing chapters.")

    def process_full_book_for_all_displayed_files(self): # Batch processing for multiple files
        if not PYAUTOGUI_AVAILABLE: messagebox.showerror("Error", "PyAutoGUI is not installed."); return
        full_book_template = self.prompts.get("full_book_prompt")
        if not full_book_template: messagebox.showerror("Error", "Full Book Prompt template not found."); return
        
        files_to_process = [item for item in self.get_displayed_file_items() if isinstance(item.get('total_chapters_for_full_book',0),int) and item.get('total_chapters_for_full_book',0)>0]
        if not files_to_process: messagebox.showinfo("Info", "No suitable files for Full Book found among displayed files."); return
        
        confirm_msg = (f"A Full Book summary will be processed in AI Studio for {len(files_to_process)} file(s).\n"
                       "Each file will be processed in groups based on its 'Total Chapters' count.\nContinue?")
        if not messagebox.askyesno("Confirm", confirm_msg, parent=self.root): return

        self.root.config(cursor="watch"); self.root.update_idletasks()
        total_files_processed_successfully = 0

        for file_idx, file_item in enumerate(files_to_process):
            file_path = file_item['path']; filename = file_item['filename']; file_type = file_item['type']
            target_total_chapters = file_item.get('total_chapters_for_full_book', 0)
            
            print(f"Batch Full Book Process: {filename} (targeting {target_total_chapters} chapters)")
            
            chunk_size = 3
            num_chunks = (target_total_chapters + chunk_size - 1) // chunk_size
            processed_chunks_for_this_pdf = 0

            for i in range(num_chunks):
                start_conceptual_idx = i * chunk_size
                end_conceptual_idx = min((i + 1) * chunk_size, target_total_chapters)
                current_chunk_indices = list(range(start_conceptual_idx, end_conceptual_idx))

                if not current_chunk_indices: continue

                chapters_for_template = self._get_chapters_text_for_template(file_item, chapter_indices_for_full_book=current_chunk_indices)
                item_description = f"{filename} (Full Book - Group {i+1}/{num_chunks}: {chapters_for_template})"
                instructional_prompt_text = self._prepare_instructional_prompt(full_book_template, chapters_for_template)
                
                data_for_clipboard = ""; is_file_object = False; prompt_to_paste_after = instructional_prompt_text

                if file_type == 'txt' and platform.system() == "Windows" and PYWIN32_AVAILABLE:
                    if self._copy_file_to_clipboard_windows(file_path): data_for_clipboard = file_path; is_file_object = True
                    else: 
                        extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
                        data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"
                        prompt_to_paste_after = "" 
                elif file_type == 'pdf': 
                    extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO PDF CONTENT]"
                    data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"
                    prompt_to_paste_after = ""
                else: # TXT non-Win
                    extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
                    data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"
                    prompt_to_paste_after = ""

                if not self._execute_ai_studio_automation(data_for_clipboard, is_file_object, prompt_to_paste_after, item_description):
                    self.root.config(cursor=""); messagebox.showerror("Automation Error", f"Automation stopped while processing '{item_description}'."); return 
                
                processed_chunks_for_this_pdf += 1
                if processed_chunks_for_this_pdf < num_chunks : time.sleep(NEXT_TAB_DELAY)
            
            if processed_chunks_for_this_pdf > 0: total_files_processed_successfully += 1
            if file_idx < len(files_to_process) - 1:
                print(f"Waiting {NEXT_FILE_PROCESSING_DELAY}s for the next file...")
                time.sleep(NEXT_FILE_PROCESSING_DELAY)

        self.root.config(cursor="")
        if total_files_processed_successfully > 0: messagebox.showinfo("Completed", f"Full Book process completed for {total_files_processed_successfully} file(s).");
        elif files_to_process: messagebox.showinfo("Info", "Issues occurred during the Full Book process, or no files could be processed.")

    def show_context_menu(self, event):
        if self._edit_widget: self._commit_in_tree_edit()
        iid = self.file_tree.identify_row(event.y)
        if not iid: return
        self.file_tree.selection_set(iid)
        menu = tk.Menu(self.root, tearoff=0)
        tags = self.file_tree.item(iid, "tags")
        if "chapter_block_item" in tags:
            block_data, file_data = self.get_chapter_block_from_iid(iid)
            if block_data and file_data:
                menu.add_command(label="Process with Prompt 1", command=lambda b=block_data, f=file_data: self.process_single_chapter_context(b, f, 'prompt1'))
                menu.add_command(label="Process with Prompt 2", command=lambda b=block_data, f=file_data: self.process_single_chapter_context(b, f, 'prompt2'))
                menu.add_command(label="Process with Prompt 3", command=lambda b=block_data, f=file_data: self.process_single_chapter_context(b, f, 'prompt3'))
        elif "file_item" in tags:
            file_data = self.get_file_data_from_tree_iid(iid)
            if file_data:
                menu.add_command(label="Process with Prompt 1 (Entire File)", command=lambda f=file_data: self.process_entire_file_context(f, 'prompt1'))
                menu.add_command(label="Process with Prompt 2 (Entire File)", command=lambda f=file_data: self.process_entire_file_context(f, 'prompt2'))
                menu.add_command(label="Process with Prompt 3 (Entire File)", command=lambda f=file_data: self.process_entire_file_context(f, 'prompt3'))
                menu.add_separator()
                full_book_valid = self.prompts.get("full_book_prompt","").strip()!="" and file_data.get('total_chapters_for_full_book',0)>0
                menu.add_command(label="Process Full Book", command=lambda f=file_data: self.process_full_book_context(f), state=tk.NORMAL if full_book_valid else tk.DISABLED)
                menu.add_separator()
                move_menu = tk.Menu(menu, tearoff=0)
                move_menu.add_command(label="Uncategorized", command=lambda fd=file_data, fid=None: self.move_file_to_folder(fd, fid))
                if self.folders: move_menu.add_separator()
                for folder in sorted(self.folders, key=lambda f: f['name'].lower()):
                    move_menu.add_command(label=folder['name'], command=lambda fd=file_data, fid=folder['id']: self.move_file_to_folder(fd, fid))
                menu.add_cascade(label="Move to Folder", menu=move_menu)
        try: menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def move_file_to_folder(self, file_item_data, target_folder_id):
        file_to_move = self.get_file_item_by_path(file_item_data['path'])
        if file_to_move: file_to_move['folder_id'] = target_folder_id; self.update_file_treeview(); self.save_app_state()
        else: messagebox.showerror("Error", "File to be moved was not found.")

    def process_single_chapter_context(self, chapter_block, file_item, prompt_key):
        current_template = self.prompts.get(prompt_key)
        if not current_template: messagebox.showerror("Error", f"Prompt template for '{prompt_key}' not found."); return
        file_path = file_item['path']; filename = file_item['filename']; file_type = file_item['type']
        page_str = chapter_block.get('page_ranges_str', "")
        chapters_for_template = self._get_chapters_text_for_template(file_item, chapter_block=chapter_block)
        item_description = f"{filename} - {chapters_for_template}"
        if file_type == 'pdf' and page_str: item_description += f" (Pages: {page_str})"
        item_description += f" ({prompt_key})"
        instructional_prompt_text = self._prepare_instructional_prompt(current_template, chapters_for_template)
        data_for_clipboard = ""; is_file_object = False; prompt_to_paste_after = instructional_prompt_text
        if file_type == 'txt' and platform.system() == "Windows" and PYWIN32_AVAILABLE:
            if self._copy_file_to_clipboard_windows(file_path): data_for_clipboard = file_path; is_file_object = True
            else: extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"; data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
        elif file_type == 'pdf':
            list_of_pages = parse_complex_page_range_string(page_str) if page_str else []
            extracted_text = self.extract_text_from_file(file_path, list_of_pages, file_type) or "[NO PDF CONTENT]"
            data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
        else: # TXT non-Win
            extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
            data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
        self.root.config(cursor="watch"); self.root.update_idletasks()
        if self._execute_ai_studio_automation(data_for_clipboard, is_file_object, prompt_to_paste_after, item_description):
            messagebox.showinfo("Completed", f"AI Studio process initiated for '{item_description}'.")
        self.root.config(cursor="")

    def process_entire_file_context(self, file_item, prompt_key):
        current_template = self.prompts.get(prompt_key)
        if not current_template: messagebox.showerror("Error", f"Prompt template for '{prompt_key}' not found."); return
        file_path = file_item['path']; filename = file_item['filename']; file_type = file_item['type']
        chapters_for_template = self._get_chapters_text_for_template(file_item) # "Entire File"
        item_description = f"{filename} (Entire File - {prompt_key})"
        instructional_prompt_text = self._prepare_instructional_prompt(current_template, chapters_for_template)
        data_for_clipboard = ""; is_file_object = False; prompt_to_paste_after = instructional_prompt_text
        if file_type == 'txt' and platform.system() == "Windows" and PYWIN32_AVAILABLE:
            if self._copy_file_to_clipboard_windows(file_path): data_for_clipboard = file_path; is_file_object = True
            else: extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"; data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
        elif file_type == 'pdf':
            extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO PDF CONTENT]"
            data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
        else: # TXT non-Win
            extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
            data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
        self.root.config(cursor="watch"); self.root.update_idletasks()
        if self._execute_ai_studio_automation(data_for_clipboard, is_file_object, prompt_to_paste_after, item_description):
            messagebox.showinfo("Completed", f"AI Studio process initiated for '{item_description}'.")
        self.root.config(cursor="")

    def process_full_book_context(self, file_item): # For single selected file, multi-chunk
        full_book_template = self.prompts.get("full_book_prompt")
        if not full_book_template: messagebox.showerror("Error", "Full Book Prompt template not found."); return
        target_total_chapters = file_item.get('total_chapters_for_full_book',0)
        if not (isinstance(target_total_chapters,int) and target_total_chapters > 0): messagebox.showinfo("Info", f"'Total Chapters' is invalid for '{file_item['filename']}'."); return
        
        file_path = file_item['path']; filename = file_item['filename']; file_type = file_item['type']
        
        confirm_msg = (f"The Full Book process for '{filename}' will be initiated in groups, targeting {target_total_chapters} chapters.\nContinue?")
        if not messagebox.askyesno("Confirm", confirm_msg, parent=self.root): return
        
        self.root.config(cursor="watch"); self.root.update_idletasks()
        
        chunk_size = 3
        num_chunks = (target_total_chapters + chunk_size - 1) // chunk_size
        processed_chunks_count = 0

        for i in range(num_chunks):
            start_conceptual_idx = i * chunk_size
            end_conceptual_idx = min((i + 1) * chunk_size, target_total_chapters)
            current_chunk_indices = list(range(start_conceptual_idx, end_conceptual_idx))

            if not current_chunk_indices: continue

            chapters_for_template = self._get_chapters_text_for_template(file_item, chapter_indices_for_full_book=current_chunk_indices)
            item_description = f"{filename} (Full Book - Group {i+1}/{num_chunks}: {chapters_for_template})"
            instructional_prompt_text = self._prepare_instructional_prompt(full_book_template, chapters_for_template)
            
            data_for_clipboard = ""; is_file_object = False; prompt_to_paste_after = instructional_prompt_text

            if file_type == 'txt' and platform.system() == "Windows" and PYWIN32_AVAILABLE:
                if self._copy_file_to_clipboard_windows(file_path): data_for_clipboard = file_path; is_file_object = True
                else: extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"; data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
            elif file_type == 'pdf': 
                extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO PDF CONTENT]"
                data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""
            else: # TXT non-Win
                extracted_text = self.extract_text_from_file(file_path, [], file_type) or "[NO TXT CONTENT]"
                data_for_clipboard = instructional_prompt_text + f"\n\nRelevant Text:\n{extracted_text}"; prompt_to_paste_after = ""

            if not self._execute_ai_studio_automation(data_for_clipboard, is_file_object, prompt_to_paste_after, item_description):
                self.root.config(cursor=""); messagebox.showerror("Automation Error", f"Automation stopped while processing '{item_description}'."); return
            
            processed_chunks_count +=1
            if processed_chunks_count < num_chunks: time.sleep(NEXT_TAB_DELAY)
            
        self.root.config(cursor="")
        if processed_chunks_count > 0: messagebox.showinfo("Completed", f"{processed_chunks_count} chapter groups processed for '{filename}'.");
        elif num_chunks > 0 : messagebox.showinfo("Info", "An issue occurred while processing Full Book.")

if __name__ == "__main__":
    if TKINTERDND2_AVAILABLE: root = TkinterDnD.Tk()
    else: root = tk.Tk()
    app = FileProcessorApp(root)
    root.mainloop()