import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import io
import re
import os
import random

# pyinstaller -w -F --add-binary="C:/Users/kod03/AppData/Local/Programs/Python/Python311/tcl/tkdnd2.8;tkdnd2.8" excel_compare.py


class ExcelCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Compare Tool")
        self.root.geometry("1200x800")
        
        # Create main container
        self.main_container = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left frame for input data
        self.left_frame = ttk.Frame(self.main_container)
        self.main_container.add(self.left_frame, weight=1)
        
        # Right frame for MasterDB with fixed width and padding
        self.right_frame = ttk.Frame(self.main_container, width=600, padding=(20, 20, 20, 20))  # Add 20px padding on all sides
        self.right_frame.pack_propagate(False)  # Prevent the frame from shrinking
        self.main_container.add(self.right_frame, weight=1)
        
        # Setup left frame
        self.setup_left_frame()
        
        # Setup right frame
        self.setup_right_frame()
        
        # Initialize data storage
        self.master_data = None
        self.input_data = None
        
        # Pattern for matching (4 letters followed by 7 digits)
        self.pattern = re.compile(r'[A-Za-z]{4}\d{7}')
        
        # Dictionary to store row colors
        self.row_colors = {}
        
        # Dictionary to store matching rows for each pattern
        self.matching_rows = {}
        
        # List to store all matching patterns
        self.all_matching_patterns = []

    def setup_left_frame(self):
        # Label for input area
        input_label = ttk.Label(self.left_frame, text="Input Data Area (Drag & Drop Excel File)", width=40)
        input_label.pack(pady=5)
        
        # Text widget for input
        self.input_text = tk.Text(self.left_frame, height=20)
        self.input_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Configure drag and drop
        self.input_text.drop_target_register(DND_FILES)
        self.input_text.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(self.left_frame, orient="vertical", command=self.input_text.yview)
        vsb.pack(side='right', fill='y')
        hsb = ttk.Scrollbar(self.left_frame, orient="horizontal", command=self.input_text.xview)
        hsb.pack(side='bottom', fill='x')
        
        self.input_text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    def setup_right_frame(self):
        # Label for MasterDB
        master_label = ttk.Label(self.right_frame, text="MasterDB (Paste Excel Data Here)", width=150)
        master_label.pack(pady=5)
        
        # Create a frame to hold the treeview and scrollbars
        tree_frame = ttk.Frame(self.right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create Treeview for MasterDB
        self.master_tree = ttk.Treeview(tree_frame)
        self.master_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add vertical scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.master_tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add horizontal scrollbar
        hsb = ttk.Scrollbar(self.right_frame, orient="horizontal", command=self.master_tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=5)
        
        # Configure the treeview to use the scrollbars
        self.master_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Store scrollbar references
        self.vsb = vsb
        self.hsb = hsb
        
        # Bind paste event
        self.master_tree.bind('<Control-v>', self.handle_master_paste)
        
        # Configure tag for highlighting
        self.master_tree.tag_configure('highlight', background='')

    def handle_master_paste(self, event):
        try:
            # Get clipboard data
            clipboard_data = self.root.clipboard_get()
            
            # Convert clipboard data to DataFrame
            df = pd.read_csv(io.StringIO(clipboard_data), sep='\t')
            
            # Update MasterDB display
            self.update_master_tree(df)
            
            # Store the data
            self.master_data = df
            
            # Clear row colors
            self.row_colors = {}
            
            # Clear matching rows
            self.matching_rows = {}
            
            # Clear matching patterns
            self.all_matching_patterns = []
            
            messagebox.showinfo("Success", "MasterDB data updated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error handling paste: {e}")

    def handle_drop(self, event):
        try:
            # Get the file path
            file_path = event.data
            
            # Remove curly braces if present (Windows file path format)
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            
            # Check if file exists
            if not os.path.exists(file_path):
                messagebox.showerror("Error", "File does not exist!")
                return
            
            # Read Excel file
            df = pd.read_excel(file_path)
            
            # Display data in input text
            self.input_text.delete(1.0, tk.END)
            self.input_text.insert(tk.END, df.to_string())
            
            # Store the data
            self.input_data = df
            
            # Find matches and highlight
            self.find_and_highlight_matches()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error handling file drop: {e}")

    def generate_pastel_color(self):
        # Generate a random pastel color
        # Pastel colors have high lightness and low saturation
        r = random.randint(180, 255)
        g = random.randint(180, 255)
        b = random.randint(180, 255)
        return f'#{r:02x}{g:02x}{b:02x}'

    def find_and_highlight_matches(self):
        if self.master_data is None or self.input_data is None:
            messagebox.showwarning("Warning", "Please paste data into MasterDB first!")
            return
        
        # Find all pattern matches in input data
        input_matches = []
        
        # Convert input data to string for pattern matching
        input_str = self.input_data.to_string()
        
        # Find all matches in input data
        for match in self.pattern.finditer(input_str):
            input_matches.append(match.group())
        
        if not input_matches:
            messagebox.showinfo("No Matches", "No matching patterns found in Input Data.")
            return
        
        # Find matching rows in MasterDB
        all_matching_rows = []
        
        # Dictionary to store matching rows for each pattern
        self.matching_rows = {}
        
        # List to store all matching patterns
        self.all_matching_patterns = []
        
        # Convert MasterDB to string for pattern matching
        master_str = self.master_data.to_string()
        
        # Check which input matches exist in MasterDB
        master_matches = []
        for match in input_matches:
            if match in master_str:
                master_matches.append(match)
                self.all_matching_patterns.append(match)
                # Find the row index in MasterDB that contains this match
                matching_rows_for_pattern = []
                for i, row in self.master_data.iterrows():
                    row_str = ' '.join(row.astype(str))
                    if match in row_str:
                        matching_rows_for_pattern.append(i)
                        all_matching_rows.append(i)
                
                # Store matching rows for this pattern
                self.matching_rows[match] = matching_rows_for_pattern
        
        # Show matches in a new window
        if master_matches:
            self.show_matches_window(master_matches)
            self.highlight_matching_rows(all_matching_rows)
        else:
            messagebox.showinfo("No Matches", "No matching patterns found in MasterDB.")

    def highlight_matching_rows(self, matching_rows):
        if self.master_data is None:
            return
        
        # Clear previous highlights
        for item in self.master_tree.get_children():
            self.master_tree.item(item, tags=())
        
        # Generate a pastel color for highlighting
        highlight_color = self.generate_pastel_color()
        
        # Apply color to matching rows
        for row_index in matching_rows:
            # Find the item in the tree that corresponds to this row
            for item in self.master_tree.get_children():
                # Get the index of this item
                item_index = int(item)
                if item_index == row_index:
                    # Apply the highlight color
                    self.master_tree.item(item, tags=('highlight',))
                    self.master_tree.tag_configure('highlight', background=highlight_color)
                    break

    def show_matches_window(self, matches):
        # Create a new window
        match_window = tk.Toplevel(self.root)
        match_window.title("Matching Patterns")
        match_window.geometry("1200x800")
        
        # Create a notebook for tabs
        notebook = ttk.Notebook(match_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a tab for pattern list
        pattern_tab = ttk.Frame(notebook)
        notebook.add(pattern_tab, text="Pattern List")
        
        # Add a label with match count
        match_count = len(matches)
        label = ttk.Label(pattern_tab, text=f"Found {match_count} matching patterns in MasterDB:")
        label.pack(pady=10)
        
        # Create a frame for the table
        table_frame = ttk.Frame(pattern_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a treeview for the table
        tree = ttk.Treeview(table_frame)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Configure columns - use all columns from MasterDB
        tree['columns'] = list(self.master_data.columns)
        tree['show'] = 'headings'
        
        # Set column headings
        for column in self.master_data.columns:
            tree.heading(column, text=column)
            tree.column(column, width=100)
        
        # Add data - only for rows that contain any match
        all_matching_rows = set()
        for match in matches:
            if match in self.matching_rows:
                all_matching_rows.update(self.matching_rows[match])
        
        for row_index in sorted(all_matching_rows):
            # Get all values for this row
            row_values = self.master_data.iloc[row_index].values.tolist()
            tree.insert('', 'end', values=row_values)
        
        # Add a copy button
        copy_button = ttk.Button(pattern_tab, text="Copy to Clipboard", 
                                command=lambda t=tree: self.copy_tree_to_clipboard(t))
        copy_button.pack(pady=10)
        
        # Create a tab for each pattern
        for match in matches:
            if match in self.matching_rows:
                # Create a tab for this pattern
                pattern_detail_tab = ttk.Frame(notebook)
                notebook.add(pattern_detail_tab, text=match)
                
                # Add a label with row count
                row_count = len(self.matching_rows[match])
                detail_label = ttk.Label(pattern_detail_tab, text=f"MasterDB rows containing '{match}' ({row_count} rows):")
                detail_label.pack(pady=10)
                
                # Create a frame for the table
                table_frame = ttk.Frame(pattern_detail_tab)
                table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                # Create a treeview for the table
                tree = ttk.Treeview(table_frame)
                tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                
                # Add scrollbars
                vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
                vsb.pack(side=tk.RIGHT, fill=tk.Y)
                hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
                hsb.pack(side=tk.BOTTOM, fill=tk.X)
                
                tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
                
                # Configure columns - use all columns from MasterDB
                tree['columns'] = list(self.master_data.columns)
                tree['show'] = 'headings'
                
                # Set column headings
                for column in self.master_data.columns:
                    tree.heading(column, text=column)
                    tree.column(column, width=100)
                
                # Add data - only for rows that contain the match
                for row_index in self.matching_rows[match]:
                    # Get all values for this row
                    row_values = self.master_data.iloc[row_index].values.tolist()
                    tree.insert('', 'end', values=row_values)
                
                # Add a copy button
                copy_button = ttk.Button(pattern_detail_tab, text="Copy to Clipboard", 
                                        command=lambda t=tree: self.copy_tree_to_clipboard(t))
                copy_button.pack(pady=10)
        
        # Add a close button
        close_button = ttk.Button(match_window, text="Close", command=match_window.destroy)
        close_button.pack(pady=10)

    def copy_tree_to_clipboard(self, tree):
        # Get all items
        items = tree.get_children()
        
        if not items:
            return
        
        # Get column names
        columns = tree['columns']
        
        # Create header row
        header = '\t'.join(columns)
        
        # Create data rows
        data_rows = []
        for item in items:
            values = tree.item(item)['values']
            data_rows.append('\t'.join(map(str, values)))
        
        # Combine header and data
        clipboard_text = header + '\n' + '\n'.join(data_rows)
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_text)
        
        messagebox.showinfo("Copied", "Data copied to clipboard!")

    def update_master_tree(self, df):
        # Clear existing items
        for item in self.master_tree.get_children():
            self.master_tree.delete(item)
        
        # Configure columns
        self.master_tree['columns'] = list(df.columns)
        self.master_tree['show'] = 'headings'
        
        # Set column headings
        for column in df.columns:
            self.master_tree.heading(column, text=column)
            self.master_tree.column(column, width=100)
        
        # Add data
        for i, row in df.iterrows():
            self.master_tree.insert('', 'end', iid=str(i), values=list(row))
            
        # Ensure scrollbars are properly configured
        if hasattr(self, 'vsb') and hasattr(self, 'hsb'):
            # Reconfigure the treeview to use the scrollbars
            self.master_tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
            
            # Make sure scrollbars are visible
            self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
            self.hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=5)

def main():
    root = TkinterDnD.Tk()
    app = ExcelCompareApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
