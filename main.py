import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from df_functions import process_excel_file

class ExcelViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Analyzer")
        self.root.geometry("1400x800")
        
        # Configure modern color scheme
        self.colors = {
            'bg_primary': '#2c3e50',      # Dark blue-gray
            'bg_secondary': '#34495e',     # Lighter blue-gray
            'bg_light': '#ecf0f1',        # Light gray
            'accent': '#3498db',          # Bright blue
            'accent_dark': '#2980b9',     # Darker blue
            'text_light': '#ffffff',      # White text
            'text_dark': '#2c3e50',       # Dark text
            'success': '#27ae60',         # Green
            'border': '#bdc3c7'           # Light border
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Configure ttk styles
        self.setup_styles()
        
        # Create main container
        self.create_main_interface()
        
        # Store dataframes and tabs
        self.dataframes = {}
        self.tabs = {}
        self.trees = {}
    
    def setup_styles(self):
        """Configure modern ttk styles"""
        style = ttk.Style()
        
        # Configure Notebook (tabs)
        style.configure('Modern.TNotebook', 
                       background=self.colors['bg_secondary'],
                       borderwidth=0)
        style.configure('Modern.TNotebook.Tab',
                       background=self.colors['bg_light'],
                       foreground='#000000',  # Black text
                       padding=[20, 10],
                       focuscolor='none',
                       borderwidth=0)
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', self.colors['accent']),
                            ('active', self.colors['accent_dark'])],
                 foreground=[('selected', '#000000'),  # Black text when selected
                            ('active', '#000000')])   # Black text when active
        
        # Configure Treeview
        style.configure('Modern.Treeview',
                       background='white',
                       foreground=self.colors['text_dark'],
                       fieldbackground='white',
                       borderwidth=1,
                       relief='solid')
        style.configure('Modern.Treeview.Heading',
                       background=self.colors['accent'],
                       foreground='#000000',  # Black text
                       borderwidth=1,
                       relief='flat',
                       font=('Segoe UI', 10, 'bold'))  # Bold font
        style.map('Modern.Treeview',
                 background=[('selected', self.colors['accent'])],
                 foreground=[('selected', self.colors['text_light'])])
        
        # Configure Buttons
        style.configure('Modern.TButton',
                       background=self.colors['accent'],
                       foreground='#000000',  # Black text
                       borderwidth=0,
                       focuscolor='none',
                       padding=[40, 7],
                       font=('Segoe UI', 10, 'bold'))  # Bold font
        style.map('Modern.TButton',
                 background=[('active', self.colors['accent_dark']),
                            ('pressed', self.colors['accent_dark'])])
        
        # Configure Quit Button
        style.configure('Quit.TButton',
                       background='#e74c3c',
                       foreground='#000000',  # Black text
                       borderwidth=0,
                       focuscolor='none',
                       padding=[20, 7],
                       font=('Segoe UI', 10, 'bold'))  # Bold font
        style.map('Quit.TButton',
                 background=[('active', '#c0392b'),
                            ('pressed', '#c0392b')])
        
        # Configure Labels
        style.configure('Modern.TLabel',
                       background=self.colors['bg_primary'],
                       foreground=self.colors['text_light'],
                       font=('Segoe UI', 10))
        
        # Configure Frames
        style.configure('Modern.TFrame',
                       background=self.colors['bg_primary'],
                       borderwidth=0)
        style.configure('Card.TFrame',
                       background=self.colors['bg_light'],
                       borderwidth=1,
                       relief='solid')
    
    def create_main_interface(self):
        """Create the main interface"""
        # Main container
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Header section
        self.create_header(main_container)
        
        # Content section
        self.create_content_area(main_container)
    
    def create_header(self, parent):
        """Create modern header with title and controls"""
        header_frame = tk.Frame(parent, bg=self.colors['bg_primary'], height=80)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # Title
        title_label = tk.Label(header_frame, 
                              text="📊 Excel Data Analyzer",
                              font=('Segoe UI', 24, 'bold'),
                              fg=self.colors['text_light'],
                              bg=self.colors['bg_primary'])
        title_label.pack(side='left', pady=20)
        
        # Control panel
        control_frame = tk.Frame(header_frame, bg=self.colors['bg_primary'])
        control_frame.pack(side='right', pady=20)
        
        # Load button
        self.load_btn = ttk.Button(control_frame, 
                                  text="📁 Load Excel File", 
                                  style='Modern.TButton',
                                  command=self.load_file)
        self.load_btn.pack(side='left', padx=(0, 15))
        
        # Quit button
        self.quit_btn = ttk.Button(control_frame, 
                                  text="✕ Quit", 
                                  style='Quit.TButton',
                                  command=self.root.quit)
        self.quit_btn.pack(side='left')
    
    def create_content_area(self, parent):
        """Create the main content area"""
        content_frame = tk.Frame(parent, bg=self.colors['bg_primary'])
        content_frame.pack(fill='both', expand=True)
        
        # Status bar
        self.create_status_bar(content_frame)
        
        # Main data area
        self.create_data_area(content_frame)
    
    def create_status_bar(self, parent):
        """Create status information bar"""
        status_frame = tk.Frame(parent, bg=self.colors['bg_secondary'], height=50)
        status_frame.pack(fill='x', pady=(0, 15))
        status_frame.pack_propagate(False)
        
        # Status icon and text
        self.status_icon = tk.Label(status_frame,
                                   text="⚪",
                                   font=('Segoe UI', 16),
                                   fg=self.colors['text_light'],
                                   bg=self.colors['bg_secondary'])
        self.status_icon.pack(side='left', padx=20, pady=12)
        
        self.info_label = tk.Label(status_frame,
                                  text="Ready to load Excel file...",
                                  font=('Segoe UI', 11),
                                  fg=self.colors['text_light'],
                                  bg=self.colors['bg_secondary'])
        self.info_label.pack(side='left', pady=15)
        
        # File info on the right
        self.file_info_label = tk.Label(status_frame,
                                       text="",
                                       font=('Segoe UI', 10),
                                       fg=self.colors['accent'],
                                       bg=self.colors['bg_secondary'])
        self.file_info_label.pack(side='right', padx=20, pady=15)
    
    def create_data_area(self, parent):
        """Create the tabbed data display area"""
        # Container for notebook
        notebook_container = tk.Frame(parent, bg=self.colors['bg_secondary'], 
                                     highlightbackground=self.colors['border'],
                                     highlightthickness=1)
        notebook_container.pack(fill='both', expand=True)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(notebook_container, style='Modern.TNotebook')
        self.notebook.pack(fill='both', expand=True, padx=2, pady=2)
        
        # Welcome tab
        self.create_welcome_tab()
    
    def create_welcome_tab(self):
        """Create welcome tab shown when no data is loaded"""
        welcome_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(welcome_frame, text="Welcome")
        
        # Welcome content
        welcome_content = tk.Frame(welcome_frame, bg='white')
        welcome_content.pack(expand=True, fill='both')
        
        # Center the welcome message
        center_frame = tk.Frame(welcome_content, bg='white')
        center_frame.pack(expand=True)
        
        # Welcome icon and text
        icon_label = tk.Label(center_frame,
                             text="📈",
                             font=('Segoe UI', 64),
                             bg='white',
                             fg=self.colors['accent'])
        icon_label.pack(pady=(100, 20))
        
        welcome_title = tk.Label(center_frame,
                                text="Welcome to Excel Data Analyzer",
                                font=('Segoe UI', 20, 'bold'),
                                bg='white',
                                fg=self.colors['text_dark'])
        welcome_title.pack(pady=(0, 10))
        
        welcome_subtitle = tk.Label(center_frame,
                                   text="Click 'Load Excel File' to get started",
                                   font=('Segoe UI', 12),
                                   bg='white',
                                   fg=self.colors['bg_secondary'])
        welcome_subtitle.pack()
    
    def create_treeview(self, parent):
        """Create a modern styled treeview"""
        # Configure parent
        parent.configure(bg='white')
        
        # Main container
        tree_container = tk.Frame(parent, bg='white')
        tree_container.pack(fill='both', expand=True, padx=15, pady=15)
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Create treeview with modern style
        tree = ttk.Treeview(tree_container, style='Modern.Treeview')
        tree.grid(row=0, column=0, sticky='nsew')
        
        # Modern scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        tree.configure(xscrollcommand=h_scrollbar.set)
        
        return tree
    
    def clear_tabs(self):
        """Remove all tabs except welcome"""
        for tab_id in list(self.tabs.keys()):
            self.notebook.forget(self.tabs[tab_id])
        
        self.tabs.clear()
        self.trees.clear()
        self.dataframes.clear()
    
    def load_file(self):
        """Load and process Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        # Update status
        self.status_icon.config(text="🔄", fg=self.colors['accent'])
        self.info_label.config(text="Processing Excel file...")
        self.root.update()
        
        try:
            # Process the entire Excel file
            processed_tabs = process_excel_file(file_path)
            
            if not processed_tabs:
                messagebox.showerror("Error", "No data could be processed from the Excel file")
                self.reset_status()
                return
            
            # Clear existing tabs (except welcome)
            self.clear_tabs()
            
            # Remove welcome tab
            self.notebook.forget(0)
            
            # Create tabs for each processed dataframe
            for tab_name, df in processed_tabs.items():
                if df is not None and not df.empty:
                    # Create tab
                    tab = tk.Frame(self.notebook, bg='white')
                    self.tabs[tab_name] = tab
                    self.notebook.add(tab, text=tab_name)
                    
                    # Create treeview for this tab
                    tree = self.create_treeview(tab)
                    self.trees[tab_name] = tree
                    
                    # Store dataframe
                    self.dataframes[tab_name] = df
                    
                    # Display the data
                    self.display_dataframe(tree, df)
            
            # Update status
            self.update_status_success(len(self.dataframes))
            self.file_info_label.config(text=f"📄 {file_path.split('/')[-1]}")
            
            # Show success message
            messagebox.showinfo("Success", f"Excel file processed successfully!\n{len(self.dataframes)} analysis tabs created")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
            self.reset_status()
    
    def display_dataframe(self, tree, df):
        """Display dataframe in treeview"""
        # Clear existing data
        for item in tree.get_children():
            tree.delete(item)
        
        if df is None or df.empty:
            return
        
        # Configure columns
        columns = list(df.columns)
        tree["columns"] = columns
        tree["show"] = "headings"
        
        # Configure column headings and widths
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=140, minwidth=100)
        
        # Insert data with alternating row colors
        for index, row in df.iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            if index % 2 == 0:
                tree.insert("", "end", values=values, tags=('evenrow',))
            else:
                tree.insert("", "end", values=values, tags=('oddrow',))
        
        # Configure row colors
        tree.tag_configure('evenrow', background='#f8f9fa')
        tree.tag_configure('oddrow', background='white')
    
    def update_status_success(self, tab_count):
        """Update status to show success"""
        self.status_icon.config(text="✅", fg=self.colors['success'])
        self.info_label.config(text=f"Successfully loaded {tab_count} data analysis tabs")
    
    def reset_status(self):
        """Reset status to default"""
        self.status_icon.config(text="⚪", fg=self.colors['text_light'])
        self.info_label.config(text="Ready to load Excel file...")
        self.file_info_label.config(text="")

# Create and run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewer(root)
    root.mainloop()