'''
Genie PDF/DjVu Outliner
A tool for managing and manipulating document outlines/bookmarks

desired Features:
shift group of outlines

- Import outlines from text/xml files
- Edit and beautify outlines
- Adjust page numbers
- Save outlines to text/xml files
'''

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog, PanedWindow
import os
import json
import xml.etree.ElementTree as ET  # For XML saving
from tkinter import scrolledtext
import re
import subprocess
from datetime import datetime
import tempfile

# Import necessary modules with error handling
try:
    from PyPDF2 import PdfReader, PdfWriter
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False
    print("PyPDF2 module not found. PDF functionality will be limited.")

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("Pandas module not found. Some functionality will be limited.")

# Import our outline parser
import outline_parser_replit as outline_parser


class OutlineParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Genie PDF/DjVu Outliner")
        self.last_pdf_path = None
        self.font_size = 18
        self.page_number_offset = 0
        self.outline_data = None  # store the parsed outline data
        self.imported_file_path = None  # To store the path of the imported outline
        self.is_beautified = False # flag if outline_text is beautified
        self.setup_ui()
        self.load_example()
        self.setup_bindings()

        # Expand the treeview by default
        self.tree.bind("<Map>", self.expand_treeview)
        self.root.report_callback_exception = self.report_callback_exception

    def report_callback_exception(self, *args):
        """Handles exceptions thrown by Tkinter callbacks."""
        import traceback
        err = traceback.format_exception(*args)
        messagebox.showerror('Exception', err)

    def setup_ui(self):
        # Configure main window
        self.root.geometry("1600x800")  # Increased width
        self.root.minsize(800, 400)

        # Menu Bar
        self.setup_menu()

        # PanedWindow for resizable panes
        self.paned_window = PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left Frame (Outline Input)
        left_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(left_frame)  # Initial width

        # Outline input section
        input_frame = ttk.LabelFrame(left_frame, text="Outline Input", padding=10)
        input_frame.pack(fill=tk.BOTH, expand=True)

        # Add a label with instructions
        hint_label = ttk.Label(input_frame, text="You can paste the results of OCR here or import from text/XML files")
        hint_label.pack(fill=tk.X, pady=(0, 5))

        # ScrolledText widget for the outline input
        self.outline_text = scrolledtext.ScrolledText(
            input_frame,
            height=15,
            wrap=tk.NONE,  # Enable horizontal scrolling
            font=('Consolas', self.font_size),
            tabs=4  # Set tab size to 4 spaces
        )
        self.outline_text.pack(fill=tk.BOTH, expand=True)

        # Add horizontal scrollbar
        self.xscrollbar = ttk.Scrollbar(input_frame, orient=tk.HORIZONTAL, command=self.outline_text.xview)
        self.xscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.outline_text.configure(xscrollcommand=self.xscrollbar.set)
        self.outline_text.bind("<Configure>", lambda e: self.outline_text.config(wrap=tk.NONE))

        # Bind text change event
        self.outline_text.bind("<<Modified>>", self.update_data_view)
        self.outline_text.bind("<KeyRelease>", self.update_data_view)
        self.outline_text.bind("<Tab>", self.indent_text)
        self.outline_text.bind("<Shift-Tab>", self.dedent_text)

        # Button frame
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)

        # Import button
        import_as_text_btn = ttk.Button(
            button_frame,
            text="Import Outline from text/xml file...",
            command=self.import_outline
        )
        import_as_text_btn.pack(side=tk.LEFT, padx=5)

        # Import from book button
        import_from_book_btn = ttk.Button(
            button_frame,
            text="Import Outline from pdf/djvu file...",
            command=self.import_from_book
        )
        import_from_book_btn.pack(side=tk.LEFT, padx=5)

        # Beautify button
        self.beautify_btn = ttk.Button(
            button_frame,
            text="Beautify Outline",
            command=self.beautify_outline
        )
        self.beautify_btn.pack(side=tk.LEFT, padx=5)

        # Save Outline button
        self.save_outline_btn = ttk.Button(
            button_frame,
            text="Save Outline",
            command=self.save_outline
        )
        self.save_outline_btn.pack(side=tk.LEFT, padx=5)

        # Clear button 
        clear_btn = ttk.Button(
            button_frame,
            text="Clear",
            command=self.clear_text
        )
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Page Number Offset
        offset_frame = ttk.Frame(left_frame)
        offset_frame.pack(fill=tk.X, pady=5)

        # Page Number Adjustment
        adjust_frame = ttk.Frame(left_frame)
        adjust_frame.pack(fill=tk.X, pady=5)

        minus_button = ttk.Button(adjust_frame, text="-", width=3, command=lambda: self.adjust_page_numbers(-1))
        minus_button.pack(side=tk.LEFT, padx=2)

        plus_button = ttk.Button(adjust_frame, text="+", width=3, command=lambda: self.adjust_page_numbers(1))
        plus_button.pack(side=tk.LEFT, padx=2)

        self.adjust_amount_entry = ttk.Entry(adjust_frame, width=5)
        self.adjust_amount_entry.pack(side=tk.LEFT, padx=5)
        self.adjust_amount_entry.insert(0, "1")  # Default adjust amount

        adjust_label = ttk.Label(adjust_frame, text="Page number adjustment amount")
        adjust_label.pack(side=tk.LEFT, padx=5)

        # PDF selection section
        pdf_frame = ttk.LabelFrame(left_frame, text="Target PDF/DjVu Book...", padding=10)
        pdf_frame.pack(fill=tk.X, pady=5)

        # PDF path entry with browse button
        pdf_select_frame = ttk.Frame(pdf_frame)
        pdf_select_frame.pack(fill=tk.X)

        self.pdf_path_var = tk.StringVar()
        pdf_entry = ttk.Entry(
            pdf_select_frame,
            textvariable=self.pdf_path_var,
            state='readonly'
        )
        pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_btn = ttk.Button(
            pdf_select_frame,
            text="Browse...",
            command=self.select_pdf
        )
        browse_btn.pack(side=tk.RIGHT)

        # Process button
        process_frame = ttk.Frame(left_frame)
        process_frame.pack(fill=tk.X, pady=5)

        process_btn = ttk.Button(
            process_frame,
            text="Process Outline and Save PDF/DjVu book",
            command=self.process_outline
        )
        process_btn.pack(fill=tk.X)

        # Right Frame (Data View)
        right_frame = ttk.Frame(self.paned_window)  # No fixed width, will expand
        self.paned_window.add(right_frame)

        # Tree view controls frame
        tree_controls_frame = ttk.Frame(right_frame)
        tree_controls_frame.pack(fill=tk.X, pady=5)
        
        # Expand/Collapse buttons
        expand_all_btn = ttk.Button(
            tree_controls_frame,
            text="Expand All",
            command=self.expand_treeview
        )
        expand_all_btn.pack(side=tk.LEFT, padx=5)
        
        collapse_all_btn = ttk.Button(
            tree_controls_frame,
            text="Collapse All",
            command=self.collapse_treeview
        )
        collapse_all_btn.pack(side=tk.LEFT, padx=5)

        # Data View Section
        data_frame = ttk.LabelFrame(right_frame, text="Resulting Outline Tree View", padding=10)
        data_frame.pack(fill=tk.BOTH, expand=True)

        self.data_frame = data_frame  # Store data_frame as an instance variable

        # Create a frame to hold the treeview and scrollbars
        tree_frame = ttk.Frame(data_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview for Outline Structure with drag and drop support
        self.tree = ttk.Treeview(tree_frame, show="tree")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Setup drag and drop
        self.tree.bind("<ButtonPress-1>", self.on_tree_drag_start)
        self.tree.bind("<B1-Motion>", self.on_tree_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_tree_drag_drop)
        self.drag_item = None
        self.drag_source = None

        # Vertical scrollbar for Treeview
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        # Horizontal scrollbar for the treeview
        self.tree_xscrollbar = ttk.Scrollbar(data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree_xscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=self.tree_xscrollbar.set)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def setup_menu(self):
        menu_bar = tk.Menu(self.root)
        # File Menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Import Outline...", command=self.import_outline)
        file_menu.add_command(label="Import from PDF/DjVu...", command=self.import_from_book)
        file_menu.add_command(label="Save Outline", command=self.save_outline)
        file_menu.add_command(label="Save Outline As...", command=self.save_outline_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Edit Menu
        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Beautify Outline", command=self.beautify_outline)
        edit_menu.add_command(label="Increase Page Numbers", command=lambda: self.adjust_page_numbers(1))
        edit_menu.add_command(label="Decrease Page Numbers", command=lambda: self.adjust_page_numbers(-1))
        edit_menu.add_command(label="Clear", command=self.clear_text)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)
        
        # View Menu
        view_menu = tk.Menu(menu_bar, tearoff=0)
        view_menu.add_command(label="Expand All", command=self.expand_treeview)
        view_menu.add_command(label="Collapse All", command=self.collapse_treeview)
        menu_bar.add_cascade(label="View", menu=view_menu)

        # Options Menu
        options_menu = tk.Menu(menu_bar, tearoff=0)
        options_menu.add_command(label="Djvused Path...", command=self.set_djvused_path)
        menu_bar.add_cascade(label="Options", menu=options_menu)

        # Language Menu
        language_menu = tk.Menu(menu_bar, tearoff=0)
        # language_menu.add_command(label="Russia", command=self.change_lang())
        menu_bar.add_cascade(label="Language", menu=language_menu)

        # Help Menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About...", command=self.show_about)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    def set_djvused_path(self):
        path = filedialog.askopenfilename(
            title="Select djvused executable",
            filetypes=[("Executable", "*.exe"), ("All Files", "*.*")]
        )
        if path:
            self.djvused_path = path
            self.status_var.set(f"djvused path set to: {path}")
        else:
            self.status_var.set("djvused path not set")

    def show_about(self):
        # Placeholder for About dialog
        messagebox.showinfo("About", "Genie PDF/DjVu Outliner\nVersion 0.3\n(c) 2024")

    def save_outline_as(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Outline As",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("XML Files", "*.xml"), ("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if file_path:
            self.save_outline_content(file_path)

    def save_outline(self):
        """Saves the outline to the same file it was imported from, asking for confirmation."""
        if self.imported_file_path:
            # Ask for confirmation before overwriting
            confirmed = messagebox.askyesno(
                "Confirm Save", 
                f"Are you sure you want to save changes to:\n{self.imported_file_path}?"
            )
            if confirmed:
                self.save_outline_content(self.imported_file_path)
            else:
                self.status_var.set("Save cancelled")
        else:
            # If no file was imported, prompt the user to save as
            self.save_outline_as()

    def save_outline_content(self, file_path):
        """Saves the outline content to the specified file."""
        if not file_path:
            return  # User cancelled the save dialog

        outline_content = self.outline_text.get(1.0, tk.END).strip()
        try:
            # Create a backup of the original file if it exists
            if os.path.exists(file_path):
                backup_path = f"{file_path}.bak"
                try:
                    with open(file_path, 'r', encoding='utf-8') as src:
                        with open(backup_path, 'w', encoding='utf-8') as dst:
                            dst.write(src.read())
                    self.status_var.set(f"Backup created: {os.path.basename(backup_path)}")
                except Exception as e:
                    messagebox.showwarning("Warning", f"Could not create backup: {str(e)}")

            if file_path.endswith(".xml"):
                # Save as XML
                try:
                    parsed_outline = outline_parser.parse_outline(outline_content)
                    root = ET.Element("outline")
                    self.create_xml_from_outline(parsed_outline, root)  # Use the recursive function

                    tree = ET.ElementTree(root)
                    ET.indent(tree, space="\t", level=0)  # Pretty-printing
                    tree.write(file_path, encoding="utf-8", xml_declaration=True)
                    self.status_var.set(f"Outline saved as XML: {os.path.basename(file_path)}")
                    # Update the imported file path
                    self.imported_file_path = file_path
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save as XML: {str(e)}")
                    self.status_var.set("Error saving outline as XML")
                    return

            elif file_path.endswith(".json"):
                # Save as JSON
                try:
                    parsed_outline = outline_parser.parse_outline(outline_content)
                    with open(file_path, 'w', encoding='utf-8') as file:
                        json.dump(parsed_outline, file, indent=4)
                    self.status_var.set(f"Outline saved as JSON: {os.path.basename(file_path)}")
                    # Update the imported file path
                    self.imported_file_path = file_path
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save as JSON: {str(e)}")
                    self.status_var.set("Error saving outline as JSON")
                    return

            else:
                # Save as Text
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(outline_content)
                self.status_var.set(f"Outline saved as text: {os.path.basename(file_path)}")
                # Update the imported file path
                self.imported_file_path = file_path
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save outline: {str(e)}")
            self.status_var.set("Error saving outline")

    def create_xml_from_outline(self, outline_items, parent_element):
        """Recursively create XML elements from outline items."""
        for item in outline_items:
            element = ET.SubElement(parent_element, "item")
            element.set("title", item.get("title", ""))
            element.set("page", str(item.get("page_number", 0)))
            
            if item.get("children"):
                self.create_xml_from_outline(item["children"], element)

    def setup_bindings(self):
        """Set up event bindings."""
        # Implement common keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self.import_outline())
        self.root.bind('<Control-s>', lambda e: self.save_outline())
        self.root.bind('<Control-b>', lambda e: self.beautify_outline())

    def load_example(self):
        """Load example outline text."""
        example_text = """PREFACE 1
CHAPTER 1. INTRODUCTION 2
    1.1 Background 3
    1.2 Purpose 4
CHAPTER 2. METHODOLOGY 5
    2.1 Research Design 6
    2.2 Data Collection 7
    2.3 Analysis 8
CHAPTER 3. RESULTS 9
CHAPTER 4. CONCLUSION 10
APPENDIX 11
"""
        self.outline_text.delete(1.0, tk.END)
        self.outline_text.insert(tk.END, example_text)
        self.update_data_view()

    def indent_text(self, event=None):
        """Handle tab key to add indentation."""
        # get selection range
        try:
            start_idx = self.outline_text.index("sel.first")
            end_idx = self.outline_text.index("sel.last")
            
            # get selected lines
            lineno_start = int(float(start_idx))
            lineno_end = int(float(end_idx))
            
            # add tabs in front of each line
            for lineno in range(lineno_start, lineno_end+1):
                self.outline_text.insert(f"{lineno}.0", "    ")
                
            return "break"  # prevent default tab behavior
        except:
            # if no selection, just insert a tab
            self.outline_text.insert(tk.INSERT, "    ")
            return "break"

    def dedent_text(self, event=None):
        """Handle Shift+Tab to remove indentation, even if nothing is selected."""
        try:
            # Check if there's a selection
            try:
                start_idx = self.outline_text.index("sel.first")
                end_idx = self.outline_text.index("sel.last")
                has_selection = True
            except tk.TclError:
                # No selection, use current line
                current_line = self.outline_text.index("insert").split('.')[0]
                start_idx = f"{current_line}.0"
                end_idx = f"{current_line}.end"
                has_selection = False
            
            lineno_start = int(float(start_idx))
            lineno_end = int(float(end_idx))
            
            # Process each line in the selection or just the current line
            for lineno in range(lineno_start, lineno_end+1):
                line_content = self.outline_text.get(f"{lineno}.0", f"{lineno}.end")
                if line_content.startswith("    "):  # 4 spaces
                    self.outline_text.delete(f"{lineno}.0", f"{lineno}.4")
                elif line_content.startswith("\t"):  # tab
                    self.outline_text.delete(f"{lineno}.0", f"{lineno}.1")
                elif line_content.startswith("  "):  # 2 spaces
                    self.outline_text.delete(f"{lineno}.0", f"{lineno}.2")
                elif line_content.startswith(" "):  # 1 space
                    self.outline_text.delete(f"{lineno}.0", f"{lineno}.1")
            
            # Update the data view after modification
            self.update_data_view()
            return "break"  # prevent default behavior
        except Exception as e:
            # Log error but don't interrupt user
            self.status_var.set(f"Error in dedent_text: {str(e)}")
            return "break"

    def expand_treeview(self, event=None):
        """Expand all items in the treeview."""
        def expand_all(node=""):
            # Get all children of the node
            children = self.tree.get_children(node)
            
            # For each child, set it as open and recurse
            for child in children:
                self.tree.item(child, open=True)  # Force open this node
                expand_all(child)  # Recursively expand all its children
                
        # Start from root nodes
        expand_all()
        self.status_var.set("Tree view expanded")

    def collapse_treeview(self, event=None):
        """Collapse all items in the treeview."""
        def collapse_all(item=""):
            children = self.tree.get_children(item)
            for child in children:
                self.tree.item(child, open=False)
                collapse_all(child)
                
        collapse_all()
        self.status_var.set("Tree view collapsed")

    # Drag and Drop functionality for tree items
    def on_tree_drag_start(self, event):
        """Start dragging a tree item"""
        # Get the item under the mouse
        item_id = self.tree.identify_row(event.y)
        if item_id:
            # Record the item and its current parent
            self.drag_item = item_id
            self.drag_source = self.tree.parent(item_id)
            self.status_var.set(f"Dragging item: {self.tree.item(item_id, 'text')}")

    def on_tree_drag_motion(self, event):
        """Visual feedback during dragging"""
        if self.drag_item:
            # Change cursor to indicate dragging
            self.tree.config(cursor="exchange")
            
            # Highlight the potential drop target
            target_id = self.tree.identify_row(event.y)
            if target_id and target_id != self.drag_item:
                self.tree.see(target_id)  # Scroll to make target visible

    def on_tree_drag_drop(self, event):
        """Handle the drop of a dragged item"""
        if not self.drag_item:
            return
        
        # Reset cursor
        self.tree.config(cursor="")
        
        # Get the drop target
        target_id = self.tree.identify_row(event.y)
        
        if target_id and target_id != self.drag_item:
            # Check if we're not dropping onto a descendant
            parent = target_id
            while parent:
                if parent == self.drag_item:
                    # Can't drop onto a descendant
                    messagebox.showinfo("Invalid Move", "Cannot drop an item onto its own descendant.")
                    self.drag_item = None
                    self.drag_source = None
                    return
                parent = self.tree.parent(parent)
                
            # Move the item
            # Get the current item text before moving
            item_text = self.tree.item(self.drag_item, 'text')
            
            # Create a new item at the target location
            if self.tree.parent(target_id):  # If dropping onto an item with a parent
                # Insert as a sibling after the target
                new_index = self.tree.index(target_id) + 1
                new_id = self.tree.insert(self.tree.parent(target_id), new_index, text=item_text)
            else:
                # Insert as a child of the target
                new_id = self.tree.insert(target_id, 'end', text=item_text)
            
            # Copy any children recursively
            self._copy_tree_structure(self.drag_item, new_id)
            
            # Delete the original item
            self.tree.delete(self.drag_item)
            
            # Update the outline text to match the new tree structure
            self.update_outline_from_tree()
            
            # Status update
            self.status_var.set(f"Moved item: {item_text}")
        
        # Reset drag state
        self.drag_item = None
        self.drag_source = None

    def _copy_tree_structure(self, source_item, target_item):
        """Recursively copy tree items"""
        for child in self.tree.get_children(source_item):
            # Create a new child in the target
            child_text = self.tree.item(child, 'text')
            new_child = self.tree.insert(target_item, 'end', text=child_text)
            
            # Recursively copy any children
            self._copy_tree_structure(child, new_child)

    def update_outline_from_tree(self):
        """Update the outline text based on the tree structure"""
        # Build a new outline text from the tree
        new_outline = self._build_outline_text(self.tree.get_children(), 0)
        
        # Update the text widget
        self.outline_text.delete(1.0, tk.END)
        self.outline_text.insert(tk.END, new_outline)
        
        # Apply beautification
        self.beautify_outline()

    def _build_outline_text(self, items, level):
        """Recursively build outline text from tree items"""
        result = ""
        for item in items:
            # Get item text and extract page number if present
            item_text = self.tree.item(item, 'text')
            page_match = re.search(r'\[Page: (\d+)\]', item_text)
            
            if page_match:
                title = item_text[:page_match.start()].strip()
                page = page_match.group(1)
            else:
                title = item_text
                page = "1"  # Default page
            
            # Add line with proper indentation
            indent = "    " * level
            result += f"{indent}{title} {page}\n"
            
            # Add children recursively
            children = self.tree.get_children(item)
            if children:
                result += self._build_outline_text(children, level+1)
        
        return result

    def import_outline(self):
        """Import outline from text file."""
        file_path = filedialog.askopenfilename(
            title="Import Outline",
            filetypes=[
                ("All Supported Files", "*.txt *.xml *.json *.bookmarks"),
                ("Text Files", "*.txt"),
                ("XML Files", "*.xml *.bookmarks"),
                ("JSON Files", "*.json"),
                ("All Files", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.endswith(('.xml', '.bookmarks')):
                # Parse XML bookmarks
                tree = ET.parse(file_path)
                root = tree.getroot()
                
                if root.tag == "content" and root.find("bookmarks") is not None:
                    # WinDjView bookmark format
                    self.parse_windj_bookmarks(file_path)
                elif root.tag == "outline" or root.find("item") is not None:
                    # Standard XML outline format
                    outline_text = self.convert_xml_to_text(root)
                    self.outline_text.delete(1.0, tk.END)
                    self.outline_text.insert(tk.END, outline_text)
                else:
                    messagebox.showerror("Error", "Unknown XML format")
                    return
                    
            elif file_path.endswith('.json'):
                # Parse JSON outline
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                outline_text = self.convert_json_to_text(data)
                self.outline_text.delete(1.0, tk.END)
                self.outline_text.insert(tk.END, outline_text)
                
            else:
                # Import text outline
                with open(file_path, 'r', encoding='utf-8') as f:
                    outline_text = f.read()
                self.outline_text.delete(1.0, tk.END)
                self.outline_text.insert(tk.END, outline_text)
            
            # Update the imported file path
            self.imported_file_path = file_path
            self.status_var.set(f"Imported outline from: {os.path.basename(file_path)}")
            
            # Apply beautification if needed and update the tree view
            self.beautify_outline()
            self.update_data_view()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import outline: {str(e)}")
            self.status_var.set("Error importing outline")

    def parse_windj_bookmarks(self, bookmark_file):
        """Parse WinDjView bookmarks and convert to text outline."""
        try:
            doc = ET.parse(bookmark_file)
            bookmarks = doc.getroot().find("bookmarks")
            
            if not bookmarks or not bookmarks.findall("bookmark"):
                raise ValueError("No bookmarks found in the file")
                
            # Build the outline text from bookmarks
            outline_text = ""
            
            for bookmark in bookmarks.findall("bookmark"):
                title = bookmark.get("title", "").strip()
                page = int(bookmark.get("page", "0"))
                indent = "    " * int(bookmark.get("level", "0"))
                outline_text += f"{indent}{title} {page+1}\n"  # Adjust page number (0-based to 1-based)
                
            self.outline_text.delete(1.0, tk.END)
            self.outline_text.insert(tk.END, outline_text)
            self.status_var.set(f"Imported WinDjView bookmarks")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse WinDjView bookmarks: {str(e)}")
            self.status_var.set("Error parsing bookmarks")

    def convert_xml_to_text(self, root, level=0):
        """Convert XML to indented text outline."""
        result = ""
        
        for item in root.findall("item"):
            title = item.get("title", "")
            page = item.get("page", "0")
            indent = "    " * level
            result += f"{indent}{title} {page}\n"
            
            if len(item) > 0:  # Has children
                result += self.convert_xml_to_text(item, level + 1)
                
        return result

    def convert_json_to_text(self, outline_data, level=0):
        """Convert JSON outline to indented text."""
        result = ""
        
        for item in outline_data:
            title = item.get("title", "")
            page = item.get("page_number", "0")
            indent = "    " * level
            result += f"{indent}{title} {page}\n"
            
            if item.get("children"):
                result += self.convert_json_to_text(item["children"], level + 1)
                
        return result

    def import_from_book(self):
        """Import outline from PDF or DjVu file."""
        if not HAS_PYPDF2:
            messagebox.showwarning("Warning", "PyPDF2 module not installed. PDF import functionality is not available.")
            return
            
        file_path = filedialog.askopenfilename(
            title="Import Outline from Book",
            filetypes=[
                ("Document Files", "*.pdf *.djvu"),
                ("PDF Files", "*.pdf"),
                ("DjVu Files", "*.djvu"),
                ("All Files", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.lower().endswith('.pdf'):
                self.import_pdf_outline(file_path)
            elif file_path.lower().endswith('.djvu'):
                self.import_djvu_outline(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import outline from book: {str(e)}")
            self.status_var.set("Error importing outline from book")

    def import_pdf_outline(self, pdf_path):
        """Import outline from PDF file."""
        if not HAS_PYPDF2:
            messagebox.showwarning("Warning", "PyPDF2 module not installed. PDF import functionality is not available.")
            return
            
        try:
            pdf = PdfReader(pdf_path)
            if pdf.outline is None or len(pdf.outline) == 0:
                messagebox.showinfo("Info", "No outline found in the PDF file.")
                return
                
            # Convert PDF outline to text
            outline_text = self.convert_pdf_outline_to_text(pdf.outline)
            
            self.outline_text.delete(1.0, tk.END)
            self.outline_text.insert(tk.END, outline_text)
            
            # Store the imported file path
            self.imported_file_path = None  # Reset as we're not importing from a text/XML file
            self.pdf_path_var.set(pdf_path)  # Set the PDF path
            
            self.status_var.set(f"Imported outline from PDF: {os.path.basename(pdf_path)}")
            self.update_data_view()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import PDF outline: {str(e)}")
            self.status_var.set("Error importing PDF outline")

    def convert_pdf_outline_to_text(self, outline_items, level=0):
        """Recursively convert PDF outline to text format.
        Handles complex outline item formats."""
        result = ""
        
        for item in outline_items:
            # Different possible formats for outline items
            if isinstance(item, dict):
                # Extract title
                title = item.get('/Title', '')
                
                # Handle various page reference formats
                page = 1  # Default page
                
                # Format 1: Direct page reference
                if '/Dest' in item and isinstance(item['/Dest'], int):
                    page = item['/Dest'] + 1  # Adjust 0-based page index
                
                # Format 2: Page reference in dictionary
                elif '/Page' in item and isinstance(item['/Page'], int):
                    page = item['/Page'] + 1
                
                # Format 3: Complex destination
                elif '/Dest' in item and isinstance(item['/Dest'], list) and len(item['/Dest']) > 0:
                    # First item might be the page reference
                    if isinstance(item['/Dest'][0], int):
                        page = item['/Dest'][0] + 1
                
                # Add this entry
                indent = "    " * level
                result += f"{indent}{title} {page}\n"
                
                # Process any sub-items
                if '/Kids' in item and isinstance(item['/Kids'], list):
                    result += self.convert_pdf_outline_to_text(item['/Kids'], level + 1)
                
            # Alternative format: item is a list
            # here is te culprit
            elif isinstance(item, list) and len(item) >= 2:
                title = item[0]
                
                # Second element might be page number
                if isinstance(item[1], int):
                    page = item[1] + 1  # Adjust 0-based page index
                else:
                    page = 1  # Default
                
                indent = "    " * level

                #here in this line is a problem, title is an object but it should be plain string
                result += f"{indent}{title} {page}\n"
                
                # Process any sub-items in list format
                if len(item) > 2 and isinstance(item[2], list):
                    result += self.convert_pdf_outline_to_text(item[2], level + 1)
        
        return result

    def import_djvu_outline(self, djvu_path):
        """Import outline from DjVu file using djvused command."""
        try:
            # Get the djvused path
            djvused_path = getattr(self, 'djvused_path', 'djvused')  # Default to 'djvused' if not set
            
            # Try to extract the outline from the DjVu file
            cmd = [djvused_path, '-e', 'print-outline', djvu_path]
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                if "djvused: command not found" in result.stderr or "command not found" in result.stderr:
                    # Prompt for djvused path instead of showing error
                    messagebox.showinfo("DjVu Tools Required", 
                                        "The 'djvused' command was not found.\n\n"
                                        "Please locate the djvused executable on your system.")
                    new_path = filedialog.askopenfilename(
                        title="Select djvused executable",
                        filetypes=[("Executable", "*.exe"), ("All Files", "*.*")]
                    )
                    if not new_path:
                        self.status_var.set("DjVu operation cancelled - djvused not set")
                        return
                    
                    # Save the path for future use
                    self.djvused_path = new_path
                    
                    # Try again with the new path
                    cmd = [new_path, '-e', 'print-outline', djvu_path]
                    result = subprocess.run(cmd, capture_output=True, text=True)
                    
                    if result.returncode != 0:
                        messagebox.showerror("Error", f"Failed to extract DjVu outline with provided path: {result.stderr}")
                        return
                else:
                    messagebox.showerror("Error", f"Failed to extract DjVu outline: {result.stderr}")
                    return
                
            outline_text = self.parse_djvused_outline(result.stdout)
            
            if not outline_text:
                messagebox.showinfo("Info", "No outline found in the DjVu file or outline format not recognized.")
                return
                
            self.outline_text.delete(1.0, tk.END)
            self.outline_text.insert(tk.END, outline_text)
            
            # Store the imported file path
            self.imported_file_path = None  # Reset as we're not importing from a text/XML file
            self.pdf_path_var.set(djvu_path)  # Set the DjVu path
            
            self.status_var.set(f"Imported outline from DjVu: {os.path.basename(djvu_path)}")
            self.update_data_view()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import DjVu outline: {str(e)}")
            self.status_var.set("Error importing DjVu outline")

    def parse_djvused_outline(self, outline_str):
        """Parse djvused outline output and convert to text format."""
        if not outline_str or outline_str.strip() == "()" or outline_str.strip() == "(bookmarks)":
            return ""
            
        # Parse djvused outline format: (bookmarks ("Title" "#Page") ("Title2" "#Page2") ...)
        result = ""
        level = 0
        
        # Simple state machine parser
        i = 0
        while i < len(outline_str):
            if outline_str[i] == '(':
                # Start of group
                if outline_str[i:i+10] == '(bookmarks':
                    # Skip 'bookmarks' keyword
                    i += 10
                else:
                    # Increase indentation level
                    level += 1
                    i += 1
            elif outline_str[i] == ')':
                # End of group
                level = max(0, level - 1)
                i += 1
            elif outline_str[i] == '"':
                # Title starts
                i += 1
                title_start = i
                while i < len(outline_str) and outline_str[i] != '"':
                    i += 1
                title = outline_str[title_start:i]
                
                # Find the page number
                page_match = re.search(r'"#(\d+)"', outline_str[i:i+10])
                if page_match:
                    page = page_match.group(1)
                    i += len(page_match.group(0))
                else:
                    page = ""
                    i += 1
                    
                # Add to result with proper indentation
                indent = "    " * level
                result += f"{indent}{title} {page}\n"
            else:
                i += 1
                
        return result

    def clear_text(self):
        """Clear the outline text field."""
        if messagebox.askyesno("Confirm", "Are you sure you want to clear the current outline?"):
            self.outline_text.delete(1.0, tk.END)
            self.update_data_view()
            self.status_var.set("Outline cleared")

    def beautify_outline(self):
        """Beautify the outline text."""
        # Get current outline text
        outline_text = self.outline_text.get(1.0, tk.END).strip()
        
        # Check if already beautified
        if hasattr(outline_parser, 'is_beautified') and outline_parser.is_beautified(outline_text):
            self.status_var.set("Outline is already beautified")
            return
        
        try:
            beautified_text = outline_parser.beautify_outline(outline_text)
            
            # Only update if there are changes
            if beautified_text != outline_text:
                self.outline_text.delete(1.0, tk.END)
                self.outline_text.insert(tk.END, beautified_text)
                self.status_var.set("Outline beautified")
                self.is_beautified = True
            else:
                self.status_var.set("No changes needed for beautification")
                
            # Update the tree view
            self.update_data_view()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to beautify outline: {str(e)}")
            self.status_var.set("Error beautifying outline")

    def update_data_view(self, event=None):
        """Update the tree view with the parsed outline data."""
        outline_text = self.outline_text.get(1.0, tk.END).strip()
        
        try:
            # Parse the outline and update the data
            self.outline_data = outline_parser.parse_outline(outline_text)
            
            # Clear and rebuild the tree
            self.tree.delete(*self.tree.get_children())
            self.populate_treeview(self.outline_data)
            
            # Expand all items by default
            self.expand_treeview()
            
            self.outline_text.edit_modified(False)  # Reset the modified flag
            
        except Exception as e:
            # Just clear the treeview if there's an error
            self.tree.delete(*self.tree.get_children())
            self.outline_data = None

    def populate_treeview(self, outline):
        """Populate the treeview with outline data."""
        def add_node(node, parent=""):
            # Add this node
            title = node.get("title", "")
            page = node.get("page_number", 0)
            item_text = f"{title} [Page: {page}]"
            
            item_id = self.tree.insert(parent, "end", text=item_text)
            
            # Add children
            children = node.get("children", [])
            for child in children:
                add_node(child, item_id)
                
        # Add all root items
        for item in outline:
            add_node(item)

    def adjust_page_numbers(self, direction):
        """Adjust all page numbers in the outline by the given direction and amount."""
        amount = 0
        try:
            amount = int(self.adjust_amount_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid adjustment amount. Please enter an integer.")
            return
            
        # Apply the direction to the amount
        adjustment = amount * direction
        
        # Get the current outline text
        outline_text = self.outline_text.get(1.0, tk.END).strip()
        lines = outline_text.split('\n')
        adjusted_lines = []
        
        # Find all page numbers and adjust them
        for line in lines:
            # Use regex to find the last number in each line (presumed to be the page number)
            match = re.search(r'(\s+)(\d+)(\s*)$', line)
            if match:
                spaces_before, page_num, spaces_after = match.groups()
                new_page = max(1, int(page_num) + adjustment)  # Ensure page is at least 1
                adjusted_lines.append(line[:match.start()] + spaces_before + str(new_page) + spaces_after)
            else:
                adjusted_lines.append(line)
                
        # Update the text
        adjusted_text = '\n'.join(adjusted_lines)
        if adjusted_text != outline_text:
            self.outline_text.delete(1.0, tk.END)
            self.outline_text.insert(tk.END, adjusted_text)
            
            # Beautify after the adjustment
            self.beautify_outline()
            
            self.status_var.set(f"Page numbers adjusted by {adjustment}")
        else:
            self.status_var.set("No page numbers found to adjust")

    def select_pdf(self):
        """Select a PDF or DjVu file for processing."""
        file_path = filedialog.askopenfilename(
            title="Select Target Document",
            filetypes=[
                ("Document Files", "*.pdf *.djvu"),
                ("PDF Files", "*.pdf"),
                ("DjVu Files", "*.djvu")
            ]
        )
        
        if file_path:
            self.pdf_path_var.set(file_path)
            self.last_pdf_path = file_path
            self.status_var.set(f"Selected document: {os.path.basename(file_path)}")

    def process_outline(self):
        """Process the outline and apply it to the selected PDF/DjVu."""
        pdf_path = self.pdf_path_var.get()
        if not pdf_path:
            messagebox.showerror("Error", "Please select a target PDF or DjVu file first.")
            return
            
        outline_text = self.outline_text.get(1.0, tk.END).strip()
        if not outline_text:
            messagebox.showerror("Error", "No outline to process.")
            return
            
        try:
            if pdf_path.lower().endswith('.pdf'):
                if not HAS_PYPDF2:
                    messagebox.showwarning("Warning", "PyPDF2 module not installed. PDF processing is not available.")
                    return
                self.process_pdf_outline(pdf_path, outline_text)
            elif pdf_path.lower().endswith('.djvu'):
                self.process_djvu_outline(pdf_path, outline_text)
            else:
                messagebox.showerror("Error", "Unsupported file format.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process outline: {str(e)}")
            self.status_var.set("Error processing outline")

    def process_pdf_outline(self, pdf_path, outline_text):
        """Apply the outline to the PDF file."""
        if not HAS_PYPDF2:
            messagebox.showwarning("Warning", "PyPDF2 module not installed. PDF processing is not available.")
            return
            
        try:
            # Update status
            self.status_var.set(f"Processing PDF outline...")
            self.root.update()
            
            # Check if PDF is encrypted or has rights restrictions
            try:
                pdf_reader = PdfReader(pdf_path)
                if pdf_reader.is_encrypted:
                    messagebox.showerror("Error", 
                                         "The PDF file is encrypted and cannot be modified.\n"
                                         "Please decrypt the PDF before attempting to edit its outline.")
                    self.status_var.set("PDF is encrypted - cannot modify")
                    return
            except Exception as e:
                if "encrypted" in str(e).lower() or "permission" in str(e).lower():
                    messagebox.showerror("Error", 
                                         f"Cannot modify this PDF due to security restrictions: {str(e)}\n"
                                         "Please try a different PDF or remove restrictions first.")
                    self.status_var.set("PDF has restrictions - cannot modify")
                    return
                else:
                    # Other error, continue with normal error handling
                    raise
            
            # Parse the outline text
            outline_data = outline_parser.parse_outline(outline_text)
            
            # Create a backup of the original PDF
            backup_path = f"{pdf_path}.bak"
            try:
                import shutil
                shutil.copy2(pdf_path, backup_path)
                self.status_var.set(f"Backup created: {os.path.basename(backup_path)}")
            except Exception as e:
                messagebox.showwarning("Warning", f"Could not create backup: {str(e)}")
            
            # Open the PDF
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            # Copy all pages
            self.status_var.set(f"Copying PDF pages...")
            self.root.update()
            for page in reader.pages:
                writer.add_page(page)
            
            # Add bookmarks
            self.status_var.set(f"Adding bookmarks...")
            self.root.update()
            
            # Recursive function to add bookmarks
            def add_outlines_to_pdf(outlines, parent):
                for outline in outlines:
                    title = outline.get("title", "")
                    page_number = outline.get("page_number", 0)
                    
                    # Ensure page number is valid
                    page_idx = min(max(0, page_number - 1), len(writer.pages) - 1)
                    
                    # Add bookmark
                    bookmark = writer.add_outline_item(title, page_idx, parent)
                    
                    # Add children recursively
                    if outline.get("children"):
                        add_outlines_to_pdf(outline["children"], bookmark)
            
            # Apply outlines
            add_outlines_to_pdf(outline_data, None)
            
            # Save the modified PDF
            self.status_var.set(f"Saving modified PDF...")
            self.root.update()
            
            # Try to save with error handling for permissions
            try:
                with open(pdf_path, "wb") as file:
                    writer.write(file)
                    
                self.status_var.set(f"Outline applied to PDF: {os.path.basename(pdf_path)}")
                messagebox.showinfo("Success", f"Outline successfully applied to {os.path.basename(pdf_path)}")
            except PermissionError:
                messagebox.showerror("Error", 
                                    "Permission denied: Cannot write to the PDF file.\n"
                                    "The file may be in use by another application or you don't have write permissions.")
                self.status_var.set("Permission error - cannot save PDF")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")
                self.status_var.set("Error saving PDF")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process PDF outline: {str(e)}")
            self.status_var.set("Error processing PDF outline")

    def process_djvu_outline(self, djvu_path, outline_text):
        """Apply the outline to the DjVu file."""
        try:
            # Parse the outline to ensure it's valid
            outline_data = outline_parser.parse_outline(outline_text)
            
            # Get the djvused path
            djvused_path = getattr(self, 'djvused_path', 'djvused')  # Default to 'djvused' if not set
            
            # Check if djvused is available
            try:
                # First test if djvused command is available
                test_cmd = [djvused_path, '--help']
                subprocess.run(test_cmd, capture_output=True, check=False)
            except FileNotFoundError:
                # If the command isn't found, prompt the user to provide the path
                messagebox.showinfo("DjVu Tools Required", 
                                     "The 'djvused' tool is required for DjVu outline operations but wasn't found.\n\n"
                                     "Please locate the djvused executable.")
                new_path = filedialog.askopenfilename(
                    title="Select djvused executable",
                    filetypes=[("Executable", "*.exe"), ("All Files", "*.*")]
                )
                if not new_path:
                    self.status_var.set("DjVu operation cancelled - djvused not found")
                    try:
                        os.unlink(temp_outline_path)
                    except:
                        pass
                    return
                djvused_path = new_path
                self.djvused_path = new_path  # Save for future use
            
            # Update status
            self.status_var.set("Generating DjVu outline...")
            self.root.update()
            
            # Generate the DjVu outline format
            djvu_outline = self.generate_djvu_outline(outline_data)
            
            # Create a temp file for the outline
            with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as f:
                temp_outline_path = f.name
                f.write(djvu_outline.encode('utf-8'))
                
            # Create a backup of the original file
            backup_path = f"{djvu_path}.bak"
            try:
                with open(djvu_path, 'rb') as src:
                    with open(backup_path, 'wb') as dst:
                        dst.write(src.read())
                self.status_var.set(f"Backup created: {os.path.basename(backup_path)}")
            except Exception as e:
                messagebox.showwarning("Warning", f"Could not create backup: {str(e)}")
                
            # Apply the outline using djvused
            cmd = [djvused_path, '-e', f'set-outline {temp_outline_path}', '-s', djvu_path]
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            # Clean up temp file
            try:
                os.unlink(temp_outline_path)
            except:
                pass
                
            if result.returncode != 0:
                messagebox.showerror("Error", f"Failed to apply DjVu outline: {result.stderr}")
                return
                
            self.status_var.set(f"Outline applied to DjVu: {os.path.basename(djvu_path)}")
            messagebox.showinfo("Success", f"Outline successfully applied to {os.path.basename(djvu_path)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process DjVu outline: {str(e)}")
            self.status_var.set("Error processing DjVu outline")

    #don't change this! it works! don't touch this functon
    def generate_djvu_outline(self, outline_data, level=0):
        """Generate DjVu outline format from our outline data."""
        result = ""
        
        if level == 0:
            result += "(bookmarks\n"
            
        for item in outline_data:
            title = item.get("title", "").replace('"', '\\"')  # Escape quotes
            page = item.get("page_number", 1)
            
            # Indent based on level
            indent = "  " * (level + 1)
            
            # Start item
            result += f'{indent}("{title}" "#{page}"'
            
            # Add children if any
            if item.get("children") and len(item["children"]) > 0:
                result += "\n"
                children_text = self.generate_djvu_outline(item["children"], level + 1)
                result += children_text
                result += f"{indent}"
                
            # Close item
            result += ")\n"
            
        if level == 0:
            result += ")\n"
            
        return result

if __name__ == "__main__":
    root = tk.Tk()
    app = OutlineParserGUI(root)
    root.mainloop()