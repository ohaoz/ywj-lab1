import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import sqlite3
from PIL import Image, ImageTk
import os
from ttkthemes import ThemedTk
import random
import math
import numpy as np
import json
import matplotlib.patheffects as path_effects

# Chart type mapping between display names and internal names
CHART_TYPE_MAP = {
    "折线图": "line",
    "柱状图": "bar",
    "散点图": "scatter",
    "饼图": "pie",
    "热力图": "heatmap",
    "直方图": "histogram"
}

# Reverse mapping for setting default values
CHART_TYPE_REVERSE_MAP = {v: k for k, v in CHART_TYPE_MAP.items()}

class DataVisualizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("数据可视化应用")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        
        # Add proper shutdown handling
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 配置Matplotlib支持中文
        self.setup_matplotlib_chinese()
        
        # Initialize variables before creating UI
        self.df = None  # Will hold pandas DataFrame
        self.db_conn = None  # Database connection
        self.sidebar_visible = True
        self.color_mode = "light"  # Default color mode
        self.theme = tk.StringVar(value="light")  # For radio buttons
        self.chart_type = tk.StringVar(value="折线图")  # Default chart type
        self.recent_files = []  # Track recent files
        self.max_recent_files = 5
        self.x_column = tk.StringVar()
        self.y_column = tk.StringVar()
        self.search_var = tk.StringVar()  # For data search/filter
        
        # 初始化翻页相关变量
        self.current_page = 0
        self.rows_per_page = 50
        self.total_pages = 1
        self.sampled_df = None
        
        # Configure the style
        self.style = ttk.Style()
        # Fix font rendering issues on Windows
        try:
            # Try using SimHei font for Chinese characters
            self.style.configure("TFrame", background="#f0f0f0")
            self.style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=6)
            self.style.configure("Accent.TButton", font=("Microsoft YaHei UI", 11, "bold"), 
                                background="#4a6cd4", foreground="#ffffff")
            self.style.configure("TLabel", font=("Microsoft YaHei UI", 11), background="#f0f0f0")
            self.style.configure("Header.TLabel", font=("Microsoft YaHei UI", 14, "bold"), background="#f0f0f0")
        except:
            # Fall back to system default if Microsoft YaHei UI is not available
            self.style.configure("TFrame", background="#f0f0f0")
            self.style.configure("TButton", padding=6)
            self.style.configure("Accent.TButton", background="#4a6cd4", foreground="#ffffff")
            self.style.configure("TLabel", background="#f0f0f0")
            self.style.configure("Header.TLabel", background="#f0f0f0")
            
        self.style.configure("Sidebar.TFrame", background="#e8e8e8")
        self.style.configure("TLabelframe", background="#f0f0f0", borderwidth=2)
        self.style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 12, "bold"), background="#f0f0f0")
        self.style.configure("TNotebook", padding=5)
        self.style.configure("TNotebook.Tab", font=("Microsoft YaHei UI", 10), padding=[10, 4])
        
        # Create main containers with improved padding
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Load user preferences before creating UI
        self.load_user_preferences()
        
        # Update color scheme first time before creating UI elements
        self.update_color_scheme()
        
        # Create UI components
        self.create_sidebar()
        self.create_content_area()
        
        # Set up keyboard shortcuts
        self.setup_shortcuts()
        
        # 延迟设置拖放功能但不显示任何错误消息
        self.root.after(500, self.setup_drag_drop)
        
    def on_closing(self):
        """Handle application shutdown cleanly"""
        try:
            # Save user preferences
            self.save_user_preferences()
            
            # Close database connection if open
            if self.db_conn is not None:
                self.db_conn.close()
                
            # Release any other resources here
            
            # Destroy the application
            self.root.destroy()
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")
            # Force destroy if there's an error during cleanup
            self.root.destroy()
    
    def center_window(self, window, width=300, height=100):
        """Center a popup window over the main window"""
        window.geometry(f"{width}x{height}")
        
        # Get the main window's position and dimensions
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        
        # Calculate the centered position
        x = main_x + (main_width - width) // 2
        y = main_y + (main_height - height) // 2
        
        # Ensure the window is visible on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Adjust if the window would appear offscreen
        if x < 0:
            x = 0
        elif x + width > screen_width:
            x = screen_width - width
            
        if y < 0:
            y = 0
        elif y + height > screen_height:
            y = screen_height - height
        
        window.geometry(f"{width}x{height}+{x}+{y}")
        
    def update_color_scheme(self, mode=None):
        """
        Update the application's color scheme based on the selected mode.
        
        Applies dark or light mode to all UI elements by configuring ttk styles
        and updating widget colors.
        
        Args:
            mode (str, optional): The color mode to apply ("dark" or "light").
                                If None, uses the current self.color_mode.
                                
        Returns:
            None
        """
        if mode:
            self.color_mode = mode
            
        # Define colors for dark and light modes
        if self.color_mode == "dark":
            bg_color = "#2d2d2d"
            frame_bg = "#333333"
            sidebar_bg = "#252525" 
            button_bg = "#444444"
            button_fg = "#f0f0f0"  # Light gray instead of white
            accent_bg = "#4a6cd4"
            accent_fg = "#f0f0f0"  # Light gray instead of white
            text_color = "#f0f0f0"  # Light gray instead of white
            label_bg = "#333333"
            entry_bg = "#3d3d3d"
            entry_fg = "#f0f0f0"  # Light gray instead of white
            table_bg = "#333333"
            table_fg = "#f0f0f0"  # Light gray instead of white
            table_selected_bg = "#4a6cd4"
            table_selected_fg = "#f0f0f0"  # Light gray for selected text
            hover_bg = "#3a3a3a"
            chart_colors = ["#4a6cd4", "#d44a6c", "#6cd44a", "#d4c84a", "#4ad4c8"]
        else:  # light mode
            bg_color = "#f0f0f0"
            frame_bg = "#f0f0f0"
            sidebar_bg = "#e8e8e8"
            button_bg = "#f0f0f0"
            button_fg = "#000000"
            accent_bg = "#4a6cd4"
            accent_fg = "#f0f0f0"  # Light gray instead of white
            text_color = "#000000"
            label_bg = "#f0f0f0"
            entry_bg = "#ffffff"
            entry_fg = "#000000"
            table_bg = "#ffffff"
            table_fg = "#000000"
            table_selected_bg = "#4a6cd4"
            table_selected_fg = "#f0f0f0"  # Light gray for selected text
            hover_bg = "#e0e0e0"
            chart_colors = ["#4a6cd4", "#d44a6c", "#6cd44a", "#d4c84a", "#4ad4c8"]
            
        # Update styles
        self.style.configure("TFrame", background=frame_bg)
        self.style.configure("TLabelframe", background=frame_bg)
        self.style.configure("TLabelframe.Label", background=frame_bg, foreground=text_color)
        
        # Update button styles with proper text contrast
        self.style.configure("TButton", background=button_bg, foreground=button_fg)
        self.style.map("TButton",
                      foreground=[('active', button_fg), ('disabled', '#888888')],
                      background=[('active', hover_bg), ('disabled', button_bg)])
        
        # Update accent button with proper text contrast
        self.style.configure("Accent.TButton", background=accent_bg, foreground=accent_fg)
        self.style.map("Accent.TButton",
                      foreground=[('active', accent_fg), ('disabled', '#cccccc')],
                      background=[('active', "#3a5cb4"), ('disabled', "#7792e6")])
        
        self.style.configure("TLabel", background=label_bg, foreground=text_color)
        self.style.configure("Header.TLabel", background=label_bg, foreground=text_color)
        self.style.configure("Sidebar.TFrame", background=sidebar_bg)
        
        # Update notebook styles
        self.style.configure("TNotebook", background=frame_bg)
        self.style.configure("TNotebook.Tab", background=button_bg, foreground=button_fg)
        self.style.map("TNotebook.Tab",
                      foreground=[('selected', accent_fg), ('active', button_fg)],
                      background=[('selected', accent_bg), ('active', hover_bg)])
        
        # Update treeview styles
        self.style.configure("Treeview", 
                          background=table_bg, 
                          foreground=table_fg,
                          fieldbackground=table_bg)
        self.style.map("Treeview",
                     background=[('selected', table_selected_bg)],
                     foreground=[('selected', table_selected_fg)])
                     
        # Store chart colors for visualization
        self.chart_colors = chart_colors
        
        # Update main container backgrounds
        self.main_frame.configure(style="TFrame")
        
        # Force redraw of all widgets
        self.root.update_idletasks()
        
        # Save user preference (could be connected to a settings system)
        # self.settings.save_preference("color_mode", self.color_mode)
        
    def create_sidebar(self):
        """Create the sidebar with data settings"""
        # Sidebar container
        self.sidebar_frame = ttk.Frame(self.main_frame, style="Sidebar.TFrame", width=250)
        self.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        self.sidebar_frame.pack_propagate(False)  # Prevent the sidebar from shrinking
        
        # Create a sidebar header with toggle button
        sidebar_header = ttk.Frame(self.sidebar_frame, style="Sidebar.TFrame")
        sidebar_header.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
        
        # Title for the sidebar
        sidebar_title = ttk.Label(sidebar_header, text="数据控制面板", style="Header.TLabel")
        sidebar_title.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Toggle button for sidebar
        toggle_icon = "◀" if self.sidebar_visible else "▶"
        self.sidebar_toggle_btn = ttk.Button(
            sidebar_header, 
            text=toggle_icon,
            command=self.toggle_sidebar,
            width=3
        )
        self.sidebar_toggle_btn.pack(side=tk.RIGHT, padx=10, pady=10)
        
        # Container for sidebar sections (can be hidden together)
        self.sidebar_content = ttk.Frame(self.sidebar_frame, style="Sidebar.TFrame")
        self.sidebar_content.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Data section
        data_section = ttk.LabelFrame(
            self.sidebar_content, 
            text="数据", 
            padding=10
        )
        data_section.pack(fill=tk.X, padx=10, pady=5)
        
        # CSV Load button with modern styling
        load_btn = ttk.Button(
            data_section, 
            text="加载CSV文件",
            style="Accent.TButton",  # Using the accent style for primary actions
            command=self.load_csv
        )
        load_btn.pack(fill=tk.X, pady=(0, 5))
        
        # Database save button
        save_btn = ttk.Button(
            data_section, 
            text="保存到数据库",
            command=self.save_to_db
        )
        save_btn.pack(fill=tk.X, pady=5)
        
        # Column selection frame
        column_frame = ttk.Frame(data_section)
        column_frame.pack(fill=tk.X, pady=5)
        
        # X axis selection
        x_label = ttk.Label(column_frame, text="X轴列:")
        x_label.grid(row=0, column=0, sticky="w", pady=2)
        
        self.x_combobox = ttk.Combobox(column_frame, textvariable=self.x_column, state="readonly")
        self.x_combobox.grid(row=0, column=1, sticky="ew", pady=2, padx=(5, 0))
        # 添加ComboBox选择变更的绑定事件
        self.x_combobox.bind("<<ComboboxSelected>>", self.on_x_selected)
        
        # Y axis selection
        y_label = ttk.Label(column_frame, text="Y轴列:")
        y_label.grid(row=1, column=0, sticky="w", pady=2)
        
        self.y_combobox = ttk.Combobox(column_frame, textvariable=self.y_column, state="readonly")
        self.y_combobox.grid(row=1, column=1, sticky="ew", pady=2, padx=(5, 0))
        # 添加ComboBox选择变更的绑定事件
        self.y_combobox.bind("<<ComboboxSelected>>", self.on_y_selected)
        
        column_frame.columnconfigure(1, weight=1)
        
        # Chart type selection
        chart_frame = ttk.Frame(data_section)
        chart_frame.pack(fill=tk.X, pady=5)
        
        chart_label = ttk.Label(chart_frame, text="图表类型:")
        chart_label.pack(side=tk.LEFT)
        
        self.chart_combobox = ttk.Combobox(
            chart_frame, 
            textvariable=self.chart_type,
            values=["折线图", "柱状图", "散点图", "饼图", "热力图", "直方图"],
            state="readonly"
        )
        self.chart_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.chart_combobox.current(0)  # Default to first option
        
        # Create plot button
        plot_btn = ttk.Button(
            data_section, 
            text="创建图表",
            command=self.create_plot
        )
        plot_btn.pack(fill=tk.X, pady=(5, 0))
        
        # Quick visualize suggestion button
        quick_btn = ttk.Button(
            data_section, 
            text="快速可视化建议",
            command=self.suggest_visualization
        )
        quick_btn.pack(fill=tk.X, pady=5)
        
        # Recent files section
        recent_frame = ttk.LabelFrame(
            self.sidebar_content, 
            text="最近文件", 
            padding=10
        )
        recent_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Will be populated as files are opened
        # Use light gray instead of white for text color in dark mode
        bg_color = "#f8f8f8" if self.color_mode == "light" else "#333333"
        fg_color = "#000000" if self.color_mode == "light" else "#f0f0f0"
        self.recent_files_list = tk.Listbox(
            recent_frame, 
            height=5, 
            bg=bg_color,
            fg=fg_color,
            selectbackground="#4a6cd4",
            selectforeground="#f0f0f0"  # Light gray instead of white
        )
        self.recent_files_list.pack(fill=tk.X, expand=True)
        self.recent_files_list.bind('<Double-Button-1>', self.load_recent_file)
        
        # Theme toggle at the bottom of sidebar
        theme_frame = ttk.Frame(self.sidebar_content, style="Sidebar.TFrame")
        theme_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        theme_label = ttk.Label(theme_frame, text="主题:")
        theme_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Using Radiobuttons for theme selection
        light_rb = ttk.Radiobutton(
            theme_frame, 
            text="明亮",
            variable=self.theme, 
            value="light",
            command=lambda: self.update_color_scheme("light")
        )
        light_rb.pack(side=tk.LEFT, padx=5)
        
        dark_rb = ttk.Radiobutton(
            theme_frame, 
            text="暗黑",
            variable=self.theme, 
            value="dark",
            command=lambda: self.update_color_scheme("dark")
        )
        dark_rb.pack(side=tk.LEFT, padx=5)
        
    def toggle_sidebar(self):
        """Collapse or expand the sidebar"""
        if self.sidebar_visible:
            # Hide the sidebar content
            self.sidebar_content.pack_forget()
            self.sidebar_toggle_btn.configure(text="▶")
            self.sidebar_frame.configure(width=50)  # Narrow width when collapsed
            self.sidebar_visible = False
        else:
            # Show the sidebar content
            self.sidebar_content.pack(fill=tk.BOTH, expand=True)
            self.sidebar_toggle_btn.configure(text="◀")
            self.sidebar_frame.configure(width=200)  # Return to normal width
            self.sidebar_visible = True
            
    def suggest_visualization(self):
        """Suggest the best visualization based on selected data"""
        if self.df is None:
            messagebox.showwarning("警告", "没有数据可视化")
            return
            
        x_col = self.x_column.get()
        y_col = self.y_column.get()
        
        if not x_col or not y_col:
            messagebox.showwarning("警告", "请选择X轴和Y轴的列")
            return
            
        # Check column types
        x_nunique = self.df[x_col].nunique() 
        
        # See if y is numeric
        try:
            y_is_numeric = pd.to_numeric(self.df[y_col], errors='coerce').notna().any()
        except:
            y_is_numeric = False
            
        if not y_is_numeric:
            messagebox.showinfo("推荐", "Y轴不是数值列，推荐使用柱状图显示计数。")
            self.chart_type.set("bar")
            
        elif x_nunique <= 7:
            # Small number of categories
            messagebox.showinfo("推荐", "检测到少量分类变量，推荐使用饼图或柱状图。")
            self.chart_type.set("pie")
            
        elif x_nunique > 50:
            # Many unique x values - could be continuous
            try:
                # Check if x can be numeric
                x_is_numeric = pd.to_numeric(self.df[x_col], errors='coerce').notna().any()
                if x_is_numeric:
                    messagebox.showinfo("推荐", "检测到两个数值变量，推荐使用散点图。")
                    self.chart_type.set("scatter")
                else:
                    messagebox.showinfo("推荐", "检测到大量类别变量，推荐使用直方图。")
                    self.chart_type.set("histogram")
            except:
                messagebox.showinfo("推荐", "检测到大量类别变量，推荐使用直方图。")
                self.chart_type.set("histogram")
                
        else:
            # Medium number of categories
            messagebox.showinfo("推荐", "检测到中等数量的类别，推荐使用柱状图。")
            self.chart_type.set("bar")
            
    def create_content_area(self):
        # Content area with tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Data tab
        self.data_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.data_tab, text="数据视图")
        
        # Data table
        self.data_frame = ttk.Frame(self.data_tab)
        self.data_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        x_scrollbar = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL)
        y_scrollbar = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL)
        
        # Treeview for data
        self.tree = ttk.Treeview(self.data_frame, 
                                 xscrollcommand=x_scrollbar.set,
                                 yscrollcommand=y_scrollbar.set)
        
        x_scrollbar.config(command=self.tree.xview)
        y_scrollbar.config(command=self.tree.yview)
        
        # Arrange components with grid
        self.tree.grid(row=0, column=0, sticky="nsew")
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.data_frame.rowconfigure(0, weight=1)
        self.data_frame.columnconfigure(0, weight=1)
        
        # Add pagination controls
        pagination_frame = ttk.Frame(self.data_tab)
        pagination_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.prev_btn = ttk.Button(pagination_frame, text="上一页", command=self.prev_page, state="disabled")
        self.prev_btn.pack(side=tk.LEFT, padx=5)
        
        self.page_label = ttk.Label(pagination_frame, text="页 0/0")
        self.page_label.pack(side=tk.LEFT, padx=5)
        
        self.next_btn = ttk.Button(pagination_frame, text="下一页", command=self.next_page, state="disabled")
        self.next_btn.pack(side=tk.LEFT, padx=5)
        
        # Add a button to change rows per page
        self.rows_btn = ttk.Button(pagination_frame, text=f"每页{self.rows_per_page}行", command=self.change_rows_per_page)
        self.rows_btn.pack(side=tk.LEFT, padx=15)
        
        # Data info label
        self.data_info_label = ttk.Label(pagination_frame, text="没有数据")
        self.data_info_label.pack(side=tk.RIGHT, padx=5)
        
        # Search/filter frame
        search_frame = ttk.Frame(self.data_tab)
        search_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT, padx=5)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        search_entry.bind("<Return>", self.search_data)
        
        ttk.Button(search_frame, text="搜索", command=self.search_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="清除", command=self.clear_search).pack(side=tk.LEFT, padx=5)
        
        # Visualization tab
        self.viz_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.viz_tab, text="可视化视图")
        
        # Frame for the plot
        self.plot_frame = ttk.Frame(self.viz_tab)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)
        
    def search_data(self, event=None):
        """
        Search functionality for the data table.
        
        Searches all columns of the DataFrame for the given search term.
        Creates a filtered DataFrame (self.filtered_df) containing only matching rows.
        Updates the data view to display only the matching results.
        
        Args:
            event: Event object from event binding (optional)
            
        Returns:
            None
        """
        if self.df is None:
            messagebox.showwarning("警告", "没有数据可搜索")
            return
            
        search_term = self.search_var.get().strip().lower()
        if not search_term:
            # If search is empty, show all data
            self.current_page = 0
            self.update_data_view()
            return
            
        # Search in all columns
        results = []
        
        # Convert all columns to string for searching
        str_df = self.df.astype(str)
        
        # Check each row
        for idx, row in str_df.iterrows():
            # Check if any cell contains the search term
            if any(search_term in cell.lower() for cell in row.values):
                results.append(idx)
                
        if not results:
            messagebox.showinfo("搜索结果", "没有找到匹配的数据")
            return
            
        # Create filtered dataframe
        self.filtered_df = self.df.loc[results]
        
        # Update pagination for filtered data
        self.current_page = 0
        self.total_pages = max(1, math.ceil(len(self.filtered_df) / self.rows_per_page))
        
        # Display filtered data
        self.display_filtered_data()
        
    def clear_search(self):
        """Clear search and show all data"""
        self.search_var.set("")
        if hasattr(self, 'filtered_df'):
            del self.filtered_df
        self.current_page = 0
        self.update_data_view()
        
    def display_filtered_data(self):
        """Display the filtered data in the treeview"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not hasattr(self, 'filtered_df') or self.filtered_df is None:
            return
            
        # Calculate start and end index for current page
        start_idx = self.current_page * self.rows_per_page
        end_idx = min(start_idx + self.rows_per_page, len(self.filtered_df))
        
        # Get the slice of data for the current page
        page_data = self.filtered_df.iloc[start_idx:end_idx]
        
        # Insert filtered data
        for _, row in page_data.iterrows():
            # Convert all values to strings and limit their length for display
            row_values = []
            for val in row:
                if val is None:
                    row_values.append("")
                else:
                    str_val = str(val)
                    if len(str_val) > 50:  # Truncate long values
                        str_val = str_val[:47] + "..."
                    row_values.append(str_val)
                    
            self.tree.insert("", "end", values=row_values)
        
        # Update pagination controls
        self.prev_btn["state"] = "normal" if self.current_page > 0 else "disabled"
        self.next_btn["state"] = "normal" if self.current_page < self.total_pages - 1 else "disabled"
        self.page_label["text"] = f"页 {self.current_page + 1}/{self.total_pages}"
        
        # Update rows per page button text
        if hasattr(self, 'rows_btn'):
            self.rows_btn["text"] = f"每页{self.rows_per_page}行"
        
        # Update data info label
        self.data_info_label["text"] = f"搜索结果: {len(self.filtered_df):,} | 显示: {start_idx + 1}-{end_idx}"
    
    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            if hasattr(self, 'filtered_df') and self.filtered_df is not None:
                self.display_filtered_data()
            else:
                self.update_data_view()
    
    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            if hasattr(self, 'filtered_df') and self.filtered_df is not None:
                self.display_filtered_data()
            else:
                self.update_data_view()

    def setup_drag_drop(self):
        """Configure drag and drop functionality for the application"""
        # This function requires TkinterDnD2 library to be installed
        # Install with: pip install tkinterdnd2
        try:
            # Try importing tkinterdnd2 directly
            import tkinterdnd2
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # Check if TkinterDnD is already properly initialized
            # Different approach depending on whether root is already a TkinterDnD.Tk
            is_tkdnd_initialized = False
            
            try:
                # Method 1: Root is already a TkinterDnD.Tk instance
                if isinstance(self.root, TkinterDnD.Tk):
                    is_tkdnd_initialized = True
                # Method 2: Root is a ThemedTk with TkinterDnD functionality
                elif hasattr(self.root.tk, 'call') and self.root.tk.call('package', 'present', 'tkdnd') == '2.9.2':
                    is_tkdnd_initialized = True
                # Method 3: Initialize TkinterDnD for any Tk instance
                else:
                    TkinterDnD._require(self.root.tk)
                    is_tkdnd_initialized = True
            except Exception as e:
                print(f"Error initializing TkinterDnD: {str(e)}")
                # Last resort - try direct tkdnd initialization
                if hasattr(self.root, 'tk'):
                    try:
                        self.root.tk.eval('package require tkdnd')
                        is_tkdnd_initialized = True
                    except Exception as inner_e:
                        print(f"Failed to initialize tkdnd package: {str(inner_e)}")
            
            # Only proceed if we successfully initialized TkinterDnD
            if is_tkdnd_initialized:
                # Register drag events with the root window
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.handle_drop)
                
                # Also register data tab for drag and drop
                self.data_tab.drop_target_register(DND_FILES)
                self.data_tab.dnd_bind('<<Drop>>', self.handle_drop)
                
                # Create a visual indicator for drag operations
                self.drop_indicator = ttk.Label(
                    self.data_tab, 
                    text="拖放CSV文件到这里", 
                    font=("Microsoft YaHei UI", 18),
                    foreground="#888888"  # Use darker gray for better visibility
                )
                # Initially hidden, will be shown during drag
                self.data_tab.bind("<<DragEnter>>", self.show_drop_indicator)
                self.data_tab.bind("<<DragLeave>>", self.hide_drop_indicator)
                
                print("Drag and drop functionality successfully initialized")
            else:
                print("Could not initialize TkinterDnD, drag and drop disabled")
                
        except ImportError as e:
            print(f"tkinterdnd2 import error: {str(e)}")
            # 不显示消息框，静默失败
        except Exception as e:
            print(f"Error setting up drag and drop: {str(e)}")
            # 不显示消息框，静默失败
        
    def show_drop_indicator(self, event):
        """Show the drop indicator during drag operations"""
        # Position in center of data tab
        self.drop_indicator.place(relx=0.5, rely=0.5, anchor="center")
        
    def hide_drop_indicator(self, event):
        """Hide the drop indicator when drag leaves"""
        self.drop_indicator.place_forget()
        
    def handle_drop(self, event):
        """Handle file drop events"""
        # Hide the drop indicator
        if hasattr(self, 'drop_indicator'):
            self.drop_indicator.place_forget()
        
        # Get file path from the drop event
        file_path = event.data
        
        # Check if it's a CSV file
        if file_path.lower().endswith('.csv'):
            self.load_csv(file_path=file_path)
        else:
            messagebox.showwarning("警告", "只支持CSV文件。请拖放CSV文件。")
    
    def load_recent_file(self, event=None):
        """Load a file from the recent files list when double-clicked"""
        try:
            # Get the selected item index
            selected_idx = self.recent_files_list.curselection()
            if not selected_idx:
                return
                
            # Get the file path from the recent files list
            file_path = self.recent_files[selected_idx[0]]
            
            # Load the selected file
            if os.path.exists(file_path):
                self.load_csv(file_path=file_path)
            else:
                messagebox.showerror("错误", f"文件不存在: {file_path}")
                # Remove from recent files list if not found
                self.recent_files.pop(selected_idx[0])
                self.update_recent_files_list()
        except Exception as e:
            messagebox.showerror("错误", f"加载文件时出错: {str(e)}")
    
    def update_recent_files_list(self):
        """Update the recent files listbox with the current list of recent files"""
        # Clear the listbox
        self.recent_files_list.delete(0, tk.END)
        
        # Add all recent files to the listbox
        for file_path in self.recent_files:
            # Get the filename without the full path
            file_name = os.path.basename(file_path)
            self.recent_files_list.insert(tk.END, file_name)
            
    def add_to_recent_files(self, file_path):
        """Add a file to the recent files list"""
        # Don't add if already at the top of the list
        if self.recent_files and self.recent_files[0] == file_path:
            return
            
        # Remove if already in the list (to move to top)
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
            
        # Add to the beginning of the list
        self.recent_files.insert(0, file_path)
        
        # Trim list if exceeds max length
        if len(self.recent_files) > self.max_recent_files:
            self.recent_files = self.recent_files[:self.max_recent_files]
            
        # Update the listbox
        self.update_recent_files_list()
    
    def load_csv(self, event=None, file_path=None):
        """
        Load a CSV file and display its data in the application.
        
        Opens a file dialog if file_path is None, otherwise uses the provided path.
        Handles various file encodings, updates pagination, column selectors, and adds
        to recent files list.
        
        Args:
            event: Event object from event binding (optional)
            file_path: Path to CSV file (optional). If None, a file dialog is shown.
            
        Returns:
            None
        """
        try:
            # If file path is not provided, open a file dialog
            if file_path is None:
                file_path = filedialog.askopenfilename(
                    title="选择CSV文件",
                    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
                )
            else:
                # 清理从拖放获得的文件路径 (可能包含引号或多余空格)
                file_path = file_path.strip()
                if file_path.startswith('"') and file_path.endswith('"'):
                    file_path = file_path[1:-1]
                elif file_path.startswith("'") and file_path.endswith("'"):
                    file_path = file_path[1:-1]
                # 处理Windows路径转义问题
                file_path = file_path.replace("\\", "/")
                
            if not file_path:
                return  # User cancelled
                
            # Check if file exists
            if not os.path.exists(file_path):
                messagebox.showerror("错误", f"文件不存在: {file_path}")
                return
                
            # Try different encodings for reading the CSV file
            encodings_to_try = ['utf-8', 'gbk', 'gb18030', 'ISO-8859-1', 'cp936', 'big5']
            
            for encoding in encodings_to_try:
                try:
                    self.df = pd.read_csv(file_path, encoding=encoding)
                    break  # If successful, break the loop
                except UnicodeDecodeError:
                    continue  # Try the next encoding
                except Exception as e:
                    # If it's not an encoding error, re-raise
                    if not isinstance(e, UnicodeDecodeError):
                        raise
            
            # If we've tried all encodings and none worked
            if self.df is None:
                messagebox.showerror("错误", "无法读取CSV文件，请检查文件编码。尝试了：" + ", ".join(encodings_to_try))
                return
            
            # Initialize pagination variables
            self.current_page = 0
            self.rows_per_page = 50
            self.total_pages = max(1, math.ceil(len(self.df) / self.rows_per_page))
            
            # Create a sampled version for large datasets
            if len(self.df) > 1000:
                self.sampled_df = self.df.sample(n=1000, random_state=42)
            else:
                self.sampled_df = None
            
            # Add to recent files list
            self.add_to_recent_files(file_path)
            
            # Update the data view
            self.update_data_view()
            
            # Update column selectors with new data columns
            self.update_column_selectors()
            
            # Switch to the data tab if we're using a notebook
            if hasattr(self, 'notebook'):
                self.notebook.select(0)  # Select the first (data) tab
                
            # Show success message
            messagebox.showinfo("成功", f"已加载文件: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载CSV文件时出错: {str(e)}")
            print(f"Error loading CSV: {str(e)}")
    
    def load_from_db(self):
        db_path = filedialog.askopenfilename(
            title="选择SQLite数据库",
            filetypes=[("SQLite files", "*.db *.sqlite"), ("All files", "*.*")]
        )
        
        if db_path:
            try:
                # Connect to database
                self.db_conn = sqlite3.connect(db_path)
                
                # Ask user which table to load
                cursor = self.db_conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
                tables = cursor.fetchall()
                
                if not tables:
                    messagebox.showinfo("提示", "数据库中没有表")
                    return
                
                table_names = [table[0] for table in tables]
                
                # Create a simple dialog to select a table
                dialog = tk.Toplevel(self.root)
                dialog.title("选择表")
                dialog.geometry("300x200")
                dialog.transient(self.root)
                dialog.grab_set()
                
                ttk.Label(dialog, text="选择要加载的表:").pack(pady=10)
                
                table_var = tk.StringVar(value=table_names[0])
                table_selector = ttk.Combobox(dialog, textvariable=table_var, values=table_names, state="readonly")
                table_selector.pack(pady=10, padx=20, fill=tk.X)
                
                def load_selected_table():
                    selected_table = table_var.get()
                    try:
                        self.df = pd.read_sql_query(f"SELECT * FROM {selected_table}", self.db_conn)
                        self.update_data_view()
                        self.update_column_selectors()
                        messagebox.showinfo("成功", f"已加载表: {selected_table}")
                        dialog.destroy()
                    except Exception as e:
                        messagebox.showerror("错误", f"加载表时出错: {str(e)}")
                
                ttk.Button(dialog, text="加载", command=load_selected_table).pack(pady=20)
                
            except Exception as e:
                messagebox.showerror("错误", f"连接到数据库时出错: {str(e)}")
    
    def save_to_db(self):
        if self.df is None:
            messagebox.showwarning("警告", "没有数据可保存")
            return
        
        db_path = filedialog.asksaveasfilename(
            title="保存到SQLite数据库",
            defaultextension=".db",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")]
        )
        
        if db_path:
            try:
                conn = sqlite3.connect(db_path)
                
                # Ask for table name
                table_name = tk.simpledialog.askstring("表名", "输入表名:", parent=self.root)
                
                if table_name:
                    self.df.to_sql(table_name, conn, if_exists='replace', index=False)
                    messagebox.showinfo("成功", f"数据已保存到表 {table_name}")
                
                conn.close()
                
            except Exception as e:
                messagebox.showerror("错误", f"保存到数据库时出错: {str(e)}")
    
    def update_data_view(self):
        """
        Update the data table view with the current DataFrame.
        
        Handles:
        1. Clearing existing data and columns in the treeview
        2. Setting up column headers based on DataFrame columns
        3. Displaying the current page of data
        4. Updating pagination controls and data info label
        
        Uses self.df for data source and respects pagination settings
        (self.current_page, self.rows_per_page, self.total_pages).
        
        Returns:
            None
        """
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Clear columns
        for column in self.tree["columns"]:
            self.tree.heading(column, text="")
        
        if self.df is not None:
            # Configure columns
            columns = list(self.df.columns)
            self.tree.configure(columns=columns)
            
            # Set column headings
            self.tree["show"] = "headings"
            for col in columns:
                self.tree.heading(col, text=col)
                # Set a reasonable column width
                self.tree.column(col, width=100)
            
            # Calculate start and end index for current page
            start_idx = self.current_page * self.rows_per_page
            end_idx = min(start_idx + self.rows_per_page, len(self.df))
            
            # Get the slice of data for the current page
            page_data = self.df.iloc[start_idx:end_idx]
            
            # Insert data for current page
            for _, row in page_data.iterrows():
                # Convert all values to strings and limit their length for display
                row_values = []
                for val in row:
                    if val is None:
                        row_values.append("")
                    else:
                        str_val = str(val)
                        if len(str_val) > 50:  # Truncate long values
                            str_val = str_val[:47] + "..."
                        row_values.append(str_val)
                        
                self.tree.insert("", "end", values=row_values)
            
            # Update pagination controls
            self.prev_btn["state"] = "normal" if self.current_page > 0 else "disabled"
            self.next_btn["state"] = "normal" if self.current_page < self.total_pages - 1 else "disabled"
            self.page_label["text"] = f"页 {self.current_page + 1}/{self.total_pages}"
            
            # Update rows per page button text
            if hasattr(self, 'rows_btn'):
                self.rows_btn["text"] = f"每页{self.rows_per_page}行"
            
            # Update data info label
            self.data_info_label["text"] = f"总行数: {len(self.df):,} | 显示: {start_idx + 1}-{end_idx}"
        else:
            # Reset pagination if no data
            self.prev_btn["state"] = "disabled"
            self.next_btn["state"] = "disabled"
            self.page_label["text"] = "页 0/0"
            
            # Update rows per page button text
            if hasattr(self, 'rows_btn'):
                self.rows_btn["text"] = f"每页{self.rows_per_page}行"
                
            self.data_info_label["text"] = "没有数据"
    
    def update_column_selectors(self):
        """
        Update the X and Y axis column selectors with available DataFrame columns.
        
        For Y-axis, attempts to identify numeric columns and prioritizes those.
        Shows a warning if no numeric columns are detected.
        
        Returns:
            None
        """
        if self.df is not None:
            columns = list(self.df.columns)
            
            # 清除之前的值
            self.x_column.set("")
            self.y_column.set("")
            
            # All columns can be used for X-axis (categorical)
            self.x_combobox['values'] = columns
            if columns:
                self.x_combobox.current(0)
                self.x_column.set(columns[0])  # 显式设置变量值
            
            # Only numeric columns should be suggested for Y-axis
            numeric_columns = []
            for col in columns:
                # Try to identify numeric columns
                try:
                    # Check if column has any numeric values
                    if pd.to_numeric(self.df[col], errors='coerce').notna().any():
                        numeric_columns.append(col)
                except:
                    continue
            
            # If we found numeric columns, use those for Y-axis options
            if numeric_columns:
                self.y_combobox['values'] = numeric_columns
                self.y_combobox.current(0)
                self.y_column.set(numeric_columns[0])  # 显式设置变量值
            else:
                # If no numeric columns found, still allow all columns but show warning
                self.y_combobox['values'] = columns
                if len(columns) > 1:
                    self.y_combobox.current(1)
                    self.y_column.set(columns[1])  # 显式设置变量值
                elif columns:
                    self.y_combobox.current(0)
                    self.y_column.set(columns[0])  # 显式设置变量值
                messagebox.showwarning("警告", "未检测到数值列。可视化需要数值数据。")
    
    def create_plot(self):
        """
        Create and display a visualization based on the selected data and chart type.
        
        This method handles:
        1. Validating the selected columns
        2. Processing data (handling non-numeric data, outliers, etc.)
        3. Creating the appropriate chart type
        4. Displaying the chart with proper labels and settings
        
        The chart type is determined by the value in self.chart_type, which is 
        translated to the appropriate internal chart type via CHART_TYPE_MAP.
        
        Returns:
            None
        """
        if self.df is None:
            messagebox.showwarning("警告", "没有数据可视化")
            return
        
        x_col = self.x_column.get()
        y_col = self.y_column.get()
        
        # 使用ComboBox的当前值，如果get()方法没有返回值
        if not x_col and hasattr(self, 'x_combobox') and self.x_combobox.get():
            x_col = self.x_combobox.get()
        
        if not y_col and hasattr(self, 'y_combobox') and self.y_combobox.get():
            y_col = self.y_combobox.get()
        
        if not x_col or not y_col:
            messagebox.showwarning("警告", "请选择X轴和Y轴的列")
            return
            
        # 检查选择的列是否在数据集中
        if x_col not in self.df.columns:
            messagebox.showwarning("警告", f"选择的X轴列 '{x_col}' 不在数据集中")
            return
            
        if y_col not in self.df.columns:
            messagebox.showwarning("警告", f"选择的Y轴列 '{y_col}' 不在数据集中")
            return

        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("生成图表中")
        progress_window.transient(self.root)
        progress_window.grab_set()
        self.center_window(progress_window)
        ttk.Label(progress_window, text="正在生成图表，请稍候...").pack(pady=(10, 5))
        progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))
        progress_bar.start()
        progress_window.update()

        try:
            # Get the internal chart type from the display name
            chart_display_type = self.chart_type.get()
            if not chart_display_type:
                # Default to line chart if nothing selected
                chart_display_type = "折线图"
                self.chart_type.set(chart_display_type)
                
            # Map display name to internal chart type
            chart_type = CHART_TYPE_MAP.get(chart_display_type)
            
            # If mapping failed, show error and default to line
            if chart_type is None:
                print(f"未知图表类型: {chart_display_type}, 使用默认折线图")
                chart_type = "line"
            
            # Use the sampled dataframe for plotting to improve performance
            work_df = self.sampled_df if self.sampled_df is not None else self.df
            
            # Make a copy to avoid modifying original data
            plot_df = work_df.copy()
            
            # Check if Y column has any potential numeric values
            if not pd.to_numeric(plot_df[y_col], errors='coerce').notna().any():
                progress_window.destroy()
                # Special case: If user selected a non-numeric column for Y-axis, suggest using count instead
                response = messagebox.askyesno("提示", 
                    f"列 '{y_col}' 不包含数值数据。是否要使用计数统计（每个 {x_col} 值的出现次数）代替?")
                
                if response:
                    # Show progress again
                    progress_window = tk.Toplevel(self.root)
                    progress_window.title("生成图表中")
                    progress_window.transient(self.root)
                    progress_window.grab_set()
                    self.center_window(progress_window)
                    ttk.Label(progress_window, text="正在生成图表，请稍候...").pack(pady=(10, 5))
                    progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
                    progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))
                    progress_bar.start()
                    progress_window.update()
                    
                    # For count-based visualization with large datasets, we may need to limit categories
                    value_counts = plot_df[x_col].value_counts()
                    
                    # If too many categories, show only top N
                    if len(value_counts) > 30:
                        top_n = messagebox.askyesno("大量类别", 
                            f"发现 {len(value_counts)} 个不同的类别，这可能导致图表不清晰。\n是否只显示前 30 个最常见的类别?")
                        
                        if top_n:
                            value_counts = value_counts.nlargest(30)
                    
                    # Convert to dataframe for plotting
                    count_data = value_counts.reset_index()
                    count_data.columns = [x_col, 'Count']
                    plot_df = count_data
                    y_col = 'Count'  # Use the new Count column
                else:
                    progress_window.destroy()
                    messagebox.showinfo("提示", "请选择包含数值的列作为Y轴")
                    return
            else:
                # Convert Y column to numeric, coercing errors
                plot_df[y_col] = pd.to_numeric(plot_df[y_col], errors='coerce')
                
                # Check if we have any valid numeric data after conversion
                if plot_df[y_col].isna().all():
                    progress_window.destroy()
                    messagebox.showerror("错误", f"Y轴列 '{y_col}' 不包含可用的数值数据")
                    return
                
                # Drop rows with NaN values
                plot_df = plot_df.dropna(subset=[y_col])
                
                if len(plot_df) == 0:
                    progress_window.destroy()
                    messagebox.showerror("错误", f"转换为数值后没有有效数据")
                    return
                
                # Check for outliers in the data
                Q1 = plot_df[y_col].quantile(0.25)
                Q3 = plot_df[y_col].quantile(0.75)
                IQR = Q3 - Q1
                has_extreme_outliers = plot_df[y_col].max() > Q3 + 5 * IQR
                
                # Initialize outlier handling variables
                use_log_scale = False
                outlier_msg = ""
                
                if has_extreme_outliers and y_col != 'Count':
                    # Show a dialog asking if user wants to handle outliers
                    progress_window.destroy()
                    outlier_option = messagebox.askyesnocancel(
                        title="发现异常值",
                        message=("数据中存在极端值，这可能会使图表难以解读。\n\n"
                                 "• 点击 '是' 移除异常值\n"
                                 "• 点击 '否' 使用对数刻度\n"
                                 "• 点击 '取消' 保持原始数据")
                    )
                    
                    if outlier_option is None:  # User clicked cancel - keep original data
                        pass
                    elif outlier_option:  # User clicked yes - remove outliers
                        # Create new progress window
                        progress_window = tk.Toplevel(self.root)
                        progress_window.title("处理数据中")
                        progress_window.transient(self.root)
                        progress_window.grab_set()
                        self.center_window(progress_window)
                        ttk.Label(progress_window, text="正在移除异常值...").pack(pady=(10, 5))
                        progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
                        progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))
                        progress_bar.start()
                        progress_window.update()
                        
                        # Use 1.5 * IQR method to filter outliers
                        lower_bound = Q1 - 1.5 * IQR
                        upper_bound = Q3 + 1.5 * IQR
                        old_count = len(plot_df)
                        plot_df = plot_df[(plot_df[y_col] >= lower_bound) & (plot_df[y_col] <= upper_bound)]
                        
                        if len(plot_df) < 10:  # If we filtered too much, use a more lenient approach
                            plot_df = work_df.copy()
                            plot_df[y_col] = pd.to_numeric(plot_df[y_col], errors='coerce')
                            plot_df = plot_df.dropna(subset=[y_col])
                            
                            # Just remove the most extreme values
                            lower_bound = plot_df[y_col].quantile(0.01)  # 1st percentile
                            upper_bound = plot_df[y_col].quantile(0.99)  # 99th percentile
                            plot_df = plot_df[(plot_df[y_col] >= lower_bound) & (plot_df[y_col] <= upper_bound)]
                        
                        outlier_msg = f"已移除 {old_count - len(plot_df)} 个异常值。"
                    else:  # User clicked no - use log scale
                        # Create new progress window
                        progress_window = tk.Toplevel(self.root)
                        progress_window.title("处理数据中")
                        progress_window.transient(self.root)
                        progress_window.grab_set()
                        self.center_window(progress_window)
                        ttk.Label(progress_window, text="正在应用对数刻度...").pack(pady=(10, 5))
                        progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
                        progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))
                        progress_bar.start()
                        progress_window.update()
                        
                        # Remove negative or zero values which can't be log-scaled
                        if (plot_df[y_col] <= 0).any():
                            min_positive = plot_df[plot_df[y_col] > 0][y_col].min()
                            if pd.notna(min_positive):
                                # Replace zeros and negatives with a small fraction of the minimum positive value
                                plot_df.loc[plot_df[y_col] <= 0, y_col] = min_positive * 0.1
                        
                        # Set flag to use log scale
                        use_log_scale = True
                        outlier_msg = "已应用对数刻度。"
                
                # If there are too many data points for a meaningful visualization, sample or aggregate
                if len(plot_df) > 1000 and chart_type in ["scatter", "line"]:
                    # For scatter plots with many points, sample for better performance
                    plot_df = plot_df.sample(n=1000, random_state=42)
            
            # Clear previous plot
            for widget in self.plot_frame.winfo_children():
                widget.destroy()
            
            # Create figure and axis with higher resolution for large datasets
            fig, ax = plt.subplots(figsize=(10, 6), dpi=120)
            
            # 确保中文字体设置生效
            try:
                if hasattr(self, 'apply_chinese_font_to_plot'):
                    self.apply_chinese_font_to_plot(ax)
            except Exception as font_error:
                print(f"应用中文字体失败，将使用默认字体: {str(font_error)}")
                # 设置一个保底字体
                try:
                    import matplotlib.font_manager as fm
                    default_font = fm.FontProperties(family='sans-serif')
                    ax.set_title(ax.get_title(), fontproperties=default_font)
                    ax.set_xlabel(ax.get_xlabel(), fontproperties=default_font)
                    ax.set_ylabel(ax.get_ylabel(), fontproperties=default_font)
                except:
                    pass  # 如果连这个也失败，继续执行
            
            # For large datasets with many x values, aggregate or rotate labels
            x_values_count = plot_df[x_col].nunique()
            
            if chart_type != "pie" and chart_type != "heatmap" and x_values_count > 20:
                # Rotate x labels for better readability
                plt.xticks(rotation=45, ha='right')
                
                # If still too many values, thin out the labels
                if x_values_count > 50:
                    # Display only every nth label
                    nth_label = max(1, x_values_count // 20)
                    for i, label in enumerate(ax.xaxis.get_ticklabels()):
                        if i % nth_label != 0:
                            label.set_visible(False)
            
            if chart_type == "bar":
                # For bar charts with many categories, limit to top N
                if x_values_count > 30 and y_col != 'Count':
                    # Group by x and calculate mean of y
                    grouped = plot_df.groupby(x_col)[y_col].mean().sort_values(ascending=False)
                    top_n = grouped.head(30).index.tolist()
                    plot_df = plot_df[plot_df[x_col].isin(top_n)]
                
                # Check if columns exist in plot_df
                if x_col not in plot_df.columns or y_col not in plot_df.columns:
                    progress_window.destroy()
                    messagebox.showerror("错误", f"选择的列 '{x_col}' 或 '{y_col}' 在处理后的数据中不存在。")
                    return
                
                plot_df.plot(kind='bar', x=x_col, y=y_col, ax=ax, color='skyblue')
                
                # Apply log scale if needed
                if use_log_scale:
                    ax.set_yscale('log')
                    
            elif chart_type == "line":
                # For line charts, data needs to be sorted by x
                try:
                    # Check if columns exist in plot_df
                    if x_col not in plot_df.columns or y_col not in plot_df.columns:
                        progress_window.destroy()
                        messagebox.showerror("错误", f"选择的列 '{x_col}' 或 '{y_col}' 在处理后的数据中不存在。")
                        return
                    
                    # Check if we have enough data for plotting
                    if len(plot_df) < 2:
                        progress_window.destroy()
                        messagebox.showerror("错误", "没有足够的数据点创建折线图")
                        return
                        
                    # Try to identify if X is a date column
                    is_date = False
                    try:
                        # Check for a few date values (not all, for performance)
                        sample = plot_df[x_col].head(5)
                        pd.to_datetime(sample, errors='raise')
                        # If no error, it's likely a date
                        plot_df['temp_date'] = pd.to_datetime(plot_df[x_col], errors='coerce')
                        # Count how many valid dates we got
                        date_count = plot_df['temp_date'].notna().sum()
                        if date_count > len(plot_df) * 0.5:  # If more than half converted successfully
                            is_date = True
                            # Sort by the date
                            plot_df = plot_df.sort_values(by='temp_date')
                            # Drop rows where date conversion failed
                            plot_df = plot_df.dropna(subset=['temp_date'])
                    except:
                        # Not a date, continue with normal processing
                        pass
                        
                    if not is_date:
                        # Try to convert x to numeric for sorting if possible
                        try:
                            x_numeric = pd.to_numeric(plot_df[x_col], errors='coerce')
                            # If at least some values converted, consider it numeric
                            if x_numeric.notna().sum() > len(plot_df) * 0.5:
                                # Sort by the numeric values
                                plot_df['temp_num'] = x_numeric
                                plot_df = plot_df.sort_values(by='temp_num')
                                # Drop rows where numeric conversion failed
                                plot_df = plot_df.dropna(subset=['temp_num'])
                            else:
                                # Not enough numeric values, sort as strings
                                plot_df = plot_df.sort_values(by=x_col)
                        except:
                            # Fall back to string sort
                            plot_df = plot_df.sort_values(by=x_col)
                    
                    # Check if we still have enough data after conversions
                    if len(plot_df) < 2:
                        progress_window.destroy()
                        messagebox.showerror("错误", "处理后没有足够的数据点创建折线图")
                        return
                    
                    # Create the line plot
                    plot_df.plot(kind='line', x=x_col, y=y_col, ax=ax, color='green', marker='o', markersize=4)
                    
                    # Apply log scale if needed
                    if use_log_scale:
                        ax.set_yscale('log')
                        
                    # Rotate x-axis labels if there are many points or they are long strings
                    if len(plot_df) > 10 or (isinstance(plot_df[x_col].iloc[0], str) and len(str(plot_df[x_col].iloc[0])) > 8):
                        plt.xticks(rotation=45, ha='right')
                        
                    # Limit x-ticks if there are too many points
                    if len(plot_df) > 20:
                        # Show only every nth tick
                        ax.xaxis.set_major_locator(plt.MaxNLocator(20))
                    
                except Exception as line_error:
                    progress_window.destroy()
                    messagebox.showerror("错误", f"创建折线图时出错: {str(line_error)}")
                    return
            
            elif chart_type == "scatter":
                # Check if columns exist in plot_df
                if x_col not in plot_df.columns or y_col not in plot_df.columns:
                    progress_window.destroy()
                    messagebox.showerror("错误", f"选择的列 '{x_col}' 或 '{y_col}' 在处理后的数据中不存在。")
                    return
                
                plot_df.plot(kind='scatter', x=x_col, y=y_col, ax=ax, color='purple', alpha=0.6)
                
                # Apply log scale if needed
                if use_log_scale:
                    ax.set_yscale('log')
                    
            elif chart_type == "pie":
                try:
                    # For pie charts, we need to handle them specially
                    # Check if columns exist in plot_df
                    if x_col not in plot_df.columns:
                        progress_window.destroy()
                        messagebox.showerror("错误", f"选择的列 '{x_col}' 在处理后的数据中不存在。")
                        return
                    
                    if y_col != 'Count' and y_col not in plot_df.columns:
                        progress_window.destroy()
                        messagebox.showerror("错误", f"选择的列 '{y_col}' 在处理后的数据中不存在。")
                        return
                    
                    # Prepare pie data
                    if y_col == 'Count':  # If we're using our count-based approach
                        pie_data = plot_df.set_index(x_col)[y_col]
                    else:
                        # Group by x and sum y values
                        pie_data = plot_df.groupby(x_col)[y_col].sum()
                    
                    # Handle negative values which can't be shown in a pie chart
                    if (pie_data < 0).any():
                        progress_window.destroy()
                        messagebox.showwarning("警告", "饼图不能包含负值。将使用绝对值。")
                        pie_data = pie_data.abs()
                        
                    # Make sure we have data to plot
                    if len(pie_data) == 0 or pie_data.sum() == 0:
                        progress_window.destroy()
                        messagebox.showerror("错误", "没有有效数据可用于饼图")
                        return
                    
                    # If pie chart has too many categories, limit to top N plus "Others"
                    if len(pie_data) > 10:
                        top_n = pie_data.nlargest(9)
                        others_sum = pie_data[~pie_data.index.isin(top_n.index)].sum()
                        
                        # Create a new Series with top_n values and an "Others" category
                        if others_sum > 0:
                            pie_data = pd.concat([top_n, pd.Series({"其他": others_sum})])
                    
                    # Create the pie chart with percentage labels and a legend
                    wedges, texts, autotexts = pie_data.plot(
                        kind='pie', 
                        ax=ax, 
                        autopct='%1.1f%%',
                        startangle=90,
                        counterclock=False,
                        shadow=True,
                        explode=[0.05] * len(pie_data),  # Slightly separate all pieces
                        textprops={'fontsize': 10}
                    )
                    
                    # Format autopct labels to be more visible
                    for autotext in autotexts:
                        # 使用对比颜色以提高可读性，避免使用白色（可能在浅色背景下不可见）
                        autotext.set_fontsize(10)
                        autotext.set_weight('bold')
                        # 添加轮廓以提高对比度
                        autotext.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                    
                    # Format pie slice labels better
                    for i, text in enumerate(texts):
                        text.set_fontsize(9)
                        # 避免标签过长导致的显示问题
                        label_text = text.get_text()
                        if len(label_text) > 15:
                            text.set_text(label_text[:12] + '...')
                    
                    # If we have a legend and many categories, adjust legend position
                    if len(pie_data) > 5:
                        legend = ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
                        # 为图例应用字体设置
                        for text in legend.get_texts():
                            text.set_fontsize(9)
                
                except Exception as pie_error:
                    progress_window.destroy()
                    messagebox.showerror("错误", f"创建饼图时出错: {str(pie_error)}")
                    print(f"饼图错误: {str(pie_error)}")
                    import traceback
                    traceback.print_exc()
                    return
            
            elif chart_type == "heatmap":
                # For heatmaps, we need two categorical columns and one numeric column for the color intensity
                
                # Check if we have enough columns in the DataFrame
                if len(plot_df.columns) < 3:
                    # Need to select another column for the heatmap's second dimension
                    progress_window.destroy()
                    # Show a dialog to select another column
                    cols = [col for col in plot_df.columns if col != x_col and col != y_col]
                    if not cols:
                        messagebox.showerror("错误", "数据中没有足够的列来创建热力图")
                        return
                        
                    dialog = tk.Toplevel(self.root)
                    dialog.title("选择热力图分组列")
                    dialog.transient(self.root)
                    dialog.grab_set()
                    self.center_window(dialog, width=350, height=180)
                    
                    ttk.Label(dialog, text="请选择热力图的第二个维度（分组列）:").pack(pady=(15, 5))
                    
                    group_var = tk.StringVar(value=cols[0])
                    group_combo = ttk.Combobox(dialog, textvariable=group_var, values=cols, state="readonly")
                    group_combo.pack(pady=5, padx=20, fill=tk.X)
                    
                    def on_group_selected():
                        nonlocal group_col
                        group_col = group_var.get()
                        dialog.destroy()
                        
                        # Restart plot creation
                        self.create_heatmap_plot(plot_df, x_col, y_col, group_col)
                    
                    group_col = None
                    ttk.Button(dialog, text="确定", command=on_group_selected).pack(pady=15)
                    return
                else:
                    # Use the third column as grouping variable if available
                    group_candidates = [col for col in plot_df.columns if col != x_col and col != y_col]
                    if group_candidates:
                        group_col = group_candidates[0]
                        self.create_heatmap_plot(plot_df, x_col, y_col, group_col)
                        return
                    else:
                        progress_window.destroy()
                        messagebox.showerror("错误", "没有足够的列创建热力图")
                        return
                    
            elif chart_type == "histogram":
                # Check if y column exists in plot_df
                if y_col not in plot_df.columns:
                    progress_window.destroy()
                    messagebox.showerror("错误", f"选择的列 '{y_col}' 在处理后的数据中不存在。")
                    return
                
                # For histograms, calculate an appropriate number of bins based on data distribution
                if len(plot_df) > 0:
                    try:
                        # Calculate Freedman-Diaconis rule for bin width
                        q75, q25 = np.percentile(plot_df[y_col], [75, 25])
                        iqr = q75 - q25
                        if iqr > 0:
                            bin_width = 2 * iqr / (len(plot_df) ** (1/3))
                            data_range = plot_df[y_col].max() - plot_df[y_col].min()
                            # Ensure a reasonable number of bins
                            num_bins = max(10, min(50, int(data_range / bin_width) if bin_width > 0 else 20))
                        else:
                            # If IQR is zero, use square root rule
                            num_bins = min(30, max(5, int(math.sqrt(len(plot_df)))))
                        
                        # Check for outliers and adjust bin range if necessary
                        if 'has_extreme_outliers' in locals() and has_extreme_outliers and use_log_scale:
                            # Use log scale for histogram with positive values
                            if (plot_df[y_col] > 0).all():
                                ax.set_xscale('log')
                                plot_df[y_col].plot(kind='hist', ax=ax, bins=num_bins, color='orange', edgecolor='black', alpha=0.7)
                            else:
                                # Handle potential negative values by shifting data
                                min_val = plot_df[y_col].min()
                                if min_val <= 0:
                                    shift = abs(min_val) + 1
                                    shifted_data = plot_df[y_col] + shift
                                    shifted_data.plot(kind='hist', ax=ax, bins=num_bins, color='orange', edgecolor='black', alpha=0.7)
                                    # Adjust x-axis labels to show original values
                                    import matplotlib.ticker as ticker
                                    def format_fn(tick_val, tick_pos):
                                        return str(int(tick_val - shift))
                                    ax.xaxis.set_major_formatter(ticker.FuncFormatter(format_fn))
                                else:
                                    plot_df[y_col].plot(kind='hist', ax=ax, bins=num_bins, color='orange', edgecolor='black', alpha=0.7)
                        else:
                            # Normal histogram with automatic bins
                            plot_df[y_col].plot(kind='hist', ax=ax, bins=num_bins, color='orange', edgecolor='black', alpha=0.7)
                    
                    except Exception as histogram_error:
                        # Fallback to simpler histogram settings in case of error
                        print(f"直方图生成错误，使用备选方案: {str(histogram_error)}")
                        try:
                            # Simple histogram with fixed number of bins
                            plot_df[y_col].plot(kind='hist', ax=ax, bins=15, color='orange', alpha=0.7)
                        except Exception as e:
                            progress_window.destroy()
                            messagebox.showerror("错误", f"无法创建直方图: {str(e)}")
                            return
                else:
                    progress_window.destroy()
                    messagebox.showerror("错误", "没有足够的数据创建直方图")
                    return
            
            # Set appropriate title based on whether we're using counts
            if y_col == 'Count':
                ax.set_title(f"{x_col} 计数")
            else:
                if outlier_msg:
                    ax.set_title(f"{y_col} vs {x_col}\n{outlier_msg}")
                else:
                    ax.set_title(f"{y_col} vs {x_col}")
                
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
            
            # Add gridlines for better readability
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # Tight layout to make sure everything fits
            plt.tight_layout()
            
            # Create a toolbar with save option
            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            
            # Create a canvas for the plot
            canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
            canvas.draw()
            
            # Add toolbar
            toolbar_frame = ttk.Frame(self.plot_frame)
            toolbar_frame.pack(side=tk.TOP, fill=tk.X)
            toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
            toolbar.update()
            
            # Pack the canvas after the toolbar
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Close progress window
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()
            
            # Switch to visualization tab
            self.notebook.select(self.viz_tab)
            
        except Exception as e:
            # Make sure to destroy the progress window
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()
                
            messagebox.showerror("错误", f"创建图表时出错: {str(e)}")
            # Print stack trace for debugging
            import traceback
            traceback.print_exc()
    
    def create_heatmap_plot(self, plot_df, x_col, y_col, group_col):
        """
        Create and display a heatmap visualization.
        
        Creates a pivot table using the provided columns and displays it as a heatmap.
        Handles large datasets by limiting the number of ticks displayed.
        
        Args:
            plot_df: DataFrame containing the data to plot
            x_col: Column to use for the x-axis
            y_col: Column to use for the color values (numeric)
            group_col: Column to use for the y-axis grouping
            
        Returns:
            None
        """
        try:
            # Show progress window
            progress_window = tk.Toplevel(self.root)
            progress_window.title("生成热力图中")
            progress_window.transient(self.root)
            progress_window.grab_set()
            self.center_window(progress_window)
            ttk.Label(progress_window, text="正在生成热力图，请稍候...").pack(pady=(10, 5))
            progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
            progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))
            progress_bar.start()
            progress_window.update()
            
            # Create pivot table for heatmap
            pivot_data = plot_df.pivot_table(
                values=y_col,
                index=x_col,
                columns=group_col,
                aggfunc=np.mean,
                fill_value=0
            )
            
            # Clear previous plot
            for widget in self.plot_frame.winfo_children():
                widget.destroy()
                
            # Create figure and axis
            fig, ax = plt.subplots(figsize=(10, 8), dpi=120)
            
            # Apply Chinese font if available
            if hasattr(self, 'apply_chinese_font_to_plot'):
                self.apply_chinese_font_to_plot(ax)
                
            # Create heatmap
            cax = ax.matshow(pivot_data, cmap='viridis', aspect='auto')
            
            # Add colorbar
            fig.colorbar(cax, ax=ax, label=y_col)
            
            # Set x and y labels
            ax.set_xlabel(group_col)
            ax.set_ylabel(x_col)
            
            # Set x and y ticks
            x_labels = pivot_data.columns
            y_labels = pivot_data.index
            
            # Limit the number of ticks if there are too many
            if len(x_labels) > 20:
                x_step = max(1, len(x_labels) // 20)
                x_positions = np.arange(0, len(x_labels), x_step)
                x_labels = [x_labels[i] for i in x_positions]
            else:
                x_positions = np.arange(len(x_labels))
                
            if len(y_labels) > 20:
                y_step = max(1, len(y_labels) // 20)
                y_positions = np.arange(0, len(y_labels), y_step)
                y_labels = [y_labels[i] for i in y_positions]
            else:
                y_positions = np.arange(len(y_labels))
                
            ax.set_xticks(x_positions)
            ax.set_xticklabels(x_labels, rotation=45, ha='left')
            ax.set_yticks(y_positions)
            ax.set_yticklabels(y_labels)
            
            # Set title
            ax.set_title(f"{y_col} 热力图 ({x_col} vs {group_col})")
            
            # Adjust layout
            plt.tight_layout()
            
            # Create canvas for the plot
            canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
            canvas.draw()
            
            # Add toolbar
            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            toolbar_frame = ttk.Frame(self.plot_frame)
            toolbar_frame.pack(side=tk.TOP, fill=tk.X)
            toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
            toolbar.update()
            
            # Pack the canvas
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Close progress window
            if progress_window.winfo_exists():
                progress_window.destroy()
                
            # Switch to visualization tab
            self.notebook.select(self.viz_tab)
            
        except Exception as e:
            # Make sure progress window is closed
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()
                
            messagebox.showerror("错误", f"创建热力图时出错: {str(e)}")
            import traceback
            traceback.print_exc()

    def change_theme(self, event=None):
        selected_theme = self.theme.get()
        try:
            self.root.set_theme(selected_theme)
        except:
            pass

    def setup_shortcuts(self):
        """Set up keyboard shortcuts for common operations"""
        # File operations
        self.root.bind("<Control-o>", lambda e: self.load_csv())  # Ctrl+O to open file
        self.root.bind("<Control-s>", lambda e: self.save_to_db())  # Ctrl+S to save
        
        # Visualization 
        self.root.bind("<Control-p>", lambda e: self.create_plot())  # Ctrl+P to plot
        self.root.bind("<Control-q>", lambda e: self.suggest_visualization())  # Ctrl+Q for quick viz
        
        # Navigation
        self.root.bind("<Control-Tab>", lambda e: self.switch_tab())  # Ctrl+Tab to switch tabs
        self.root.bind("<F11>", lambda e: self.toggle_sidebar())  # F11 to toggle sidebar
        
        # Search
        self.root.bind("<Control-f>", lambda e: self.focus_search())  # Ctrl+F to focus search box
        
        # Page navigation
        self.root.bind("<Control-Right>", lambda e: self.next_page() if self.next_btn["state"] == "normal" else None)
        self.root.bind("<Control-Left>", lambda e: self.prev_page() if self.prev_btn["state"] == "normal" else None)
        
    def switch_tab(self):
        """Switch between data and visualization tabs"""
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab < len(self.notebook.tabs()) - 1:
            self.notebook.select(current_tab + 1)
        else:
            self.notebook.select(0)
            
    def focus_search(self):
        """Focus the search box"""
        # Find the search entry widget and focus it
        for widget in self.data_tab.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Entry):
                        child.focus_set()
                        return

    def on_x_selected(self, event=None):
        """处理X轴列选择变更事件"""
        selected = self.x_combobox.get()
        if selected:
            self.x_column.set(selected)
            print(f"已选择X轴列: {selected}")
    
    def on_y_selected(self, event=None):
        """处理Y轴列选择变更事件"""
        selected = self.y_combobox.get()
        if selected:
            self.y_column.set(selected)
            print(f"已选择Y轴列: {selected}")

    def setup_matplotlib_chinese(self):
        """配置Matplotlib以支持中文显示"""
        try:
            import matplotlib.pyplot as plt
            import matplotlib.font_manager as font_manager
            import platform
            
            # 不同操作系统使用不同的默认中文字体
            system = platform.system()
            
            # 设置全局字体为更通用的名称
            plt.rcParams['font.family'] = ['sans-serif']
            plt.rcParams['axes.unicode_minus'] = False  # 正确显示负号
            
            if system == 'Windows':
                # Windows系统尝试使用微软雅黑、宋体和黑体
                chinese_fonts = ['Microsoft YaHei', 'SimSun', 'SimHei', 'Arial Unicode MS']
            elif system == 'Darwin':  # macOS
                # macOS系统尝试使用苹方和华文黑体
                chinese_fonts = ['PingFang SC', 'STHeiti', 'Arial Unicode MS']
            elif system == 'Linux':
                # Linux系统尝试使用文泉驿和思源黑体
                chinese_fonts = ['WenQuanYi Micro Hei', 'Noto Sans CJK SC', 'Droid Sans Fallback']
            else:
                chinese_fonts = ['Arial Unicode MS', 'DejaVu Sans']  # 通用备选
            
            # 寻找第一个可用的字体
            font_found = False
            for font_name in chinese_fonts:
                for font in font_manager.findSystemFonts():
                    try:
                        if font_name.lower() in os.path.basename(font).lower():
                            plt.rcParams['font.sans-serif'] = [font_name] + plt.rcParams.get('font.sans-serif', [])
                            print(f"使用字体: {font_name}")
                            font_found = True
                            break
                    except:
                        continue
                if font_found:
                    break
            
            # 如果未找到中文字体，使用系统默认无衬线字体
            if not font_found:
                plt.rcParams['font.sans-serif'] = ['Arial', 'DejaVu Sans'] + plt.rcParams.get('font.sans-serif', [])
                print("未找到中文字体，使用默认无衬线字体")
            
            # 强制使用一个通用回退字体以防万一
            plt.rcParams['font.sans-serif'] = plt.rcParams.get('font.sans-serif', []) + ['DejaVu Sans']
            
        except Exception as e:
            print(f"配置Matplotlib字体失败: {str(e)}")
            # 出错时设置一个安全的默认值
            try:
                plt.rcParams['font.family'] = ['sans-serif']
                plt.rcParams['font.sans-serif'] = ['Arial', 'DejaVu Sans']
                plt.rcParams['axes.unicode_minus'] = False
            except:
                pass  # 如果连这个也失败，就放弃

    def apply_chinese_font_to_plot(self, ax):
        """为单个图表设置中文字体"""
        try:
            import matplotlib.font_manager as font_manager
            import platform
            
            # 获取全局已配置的字体设置作为首选
            global_fonts = plt.rcParams.get('font.sans-serif', [])
            
            # 使用系统特定的备选字体
            system = platform.system()
            if system == 'Windows':
                font_names = ['Microsoft YaHei', 'SimSun', 'SimHei', 'Arial Unicode MS']
            elif system == 'Darwin':  # macOS
                font_names = ['PingFang SC', 'STHeiti', 'Arial Unicode MS']
            elif system == 'Linux':
                font_names = ['WenQuanYi Micro Hei', 'Noto Sans CJK SC', 'Droid Sans Fallback']
            else:
                font_names = ['Arial Unicode MS', 'DejaVu Sans']
            
            # 首先尝试使用全局字体，然后再尝试备选字体
            all_fonts = global_fonts + [f for f in font_names if f not in global_fonts]
            
            # 查找可用的中文字体
            chinese_font = None
            
            # 先尝试直接使用全局配置的字体
            try:
                first_font = all_fonts[0] if all_fonts else 'Arial'
                chinese_font = font_manager.FontProperties(family=first_font)
                print(f"直接使用全局字体: {first_font}")
            except:
                pass
                
            # 如果直接使用全局字体失败，尝试按名称查找字体文件
            if chinese_font is None:
                for font_name in all_fonts:
                    try:
                        font_path = None
                        for font in font_manager.findSystemFonts():
                            try:
                                if font_name.lower() in os.path.basename(font).lower():
                                    font_path = font
                                    break
                            except:
                                continue
                        
                        if font_path:
                            chinese_font = font_manager.FontProperties(fname=font_path)
                            print(f"已应用中文字体到图表: {font_name}")
                            break
                    except Exception as font_error:
                        print(f"加载字体 {font_name} 时出错: {str(font_error)}")
                        continue
            
            # 如果找不到任何中文字体，使用系统默认字体
            if chinese_font is None:
                chinese_font = font_manager.FontProperties(family='sans-serif')
                print("未找到中文字体，使用默认无衬线字体")
            
            # 应用字体到图表中的各元素
            if ax.get_title():
                ax.set_title(ax.get_title(), fontproperties=chinese_font)
            
            ax.set_xlabel(ax.get_xlabel(), fontproperties=chinese_font)
            ax.set_ylabel(ax.get_ylabel(), fontproperties=chinese_font)
            
            # 设置X轴刻度标签的字体
            for label in ax.get_xticklabels():
                label.set_fontproperties(chinese_font)
            
            # 设置Y轴刻度标签的字体
            for label in ax.get_yticklabels():
                label.set_fontproperties(chinese_font)
            
            # 设置图例的字体
            legend = ax.get_legend()
            if legend:
                for text in legend.get_texts():
                    text.set_fontproperties(chinese_font)
            
        except Exception as e:
            print(f"应用中文字体到图表失败: {str(e)}")
            # 出错时不中断程序，只记录错误

    def save_user_preferences(self):
        """
        Save user preferences to a JSON file for persistence between sessions.
        
        Saves:
        - Recent files list
        - Color/theme mode
        - Default rows per page
        - Last used directory
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Create preferences dictionary
            preferences = {
                "recent_files": self.recent_files,
                "color_mode": self.color_mode,
                "rows_per_page": self.rows_per_page,
                "last_directory": os.path.dirname(self.recent_files[0]) if self.recent_files else ""
            }
            
            # Create preferences directory if it doesn't exist
            prefs_dir = os.path.join(os.path.expanduser("~"), ".data_vis_app")
            os.makedirs(prefs_dir, exist_ok=True)
            
            # Save to JSON file
            prefs_file = os.path.join(prefs_dir, "preferences.json")
            with open(prefs_file, 'w', encoding='utf-8') as f:
                json.dump(preferences, f, ensure_ascii=False, indent=2)
                
            return True
        except Exception as e:
            print(f"Error saving preferences: {str(e)}")
            return False
            
    def load_user_preferences(self):
        """
        Load user preferences from a JSON file.
        
        Loads:
        - Recent files list
        - Color/theme mode
        - Default rows per page
        - Last used directory
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Check if preferences file exists
            prefs_file = os.path.join(os.path.expanduser("~"), ".data_vis_app", "preferences.json")
            if not os.path.exists(prefs_file):
                return False
                
            # Load from JSON file
            with open(prefs_file, 'r', encoding='utf-8') as f:
                preferences = json.load(f)
                
            # Apply preferences
            if "recent_files" in preferences:
                # Filter only existing files
                self.recent_files = [file for file in preferences["recent_files"] 
                                    if os.path.exists(file)]
                self.update_recent_files_list()
                
            if "color_mode" in preferences:
                self.color_mode = preferences["color_mode"]
                self.theme.set(self.color_mode)
                self.update_color_scheme()
                
            if "rows_per_page" in preferences:
                self.rows_per_page = preferences["rows_per_page"]
                
            return True
        except Exception as e:
            print(f"Error loading preferences: {str(e)}")
            return False

    def change_rows_per_page(self):
        """
        Let users change the number of rows displayed per page in the data view.
        
        Opens a dialog for the user to input a new value, then updates the view
        and saves the preference.
        
        Returns:
            None
        """
        try:
            # Ask user for new rows per page
            new_rows = simpledialog.askinteger(
                "设置每页行数", 
                "请输入每页显示的行数 (10-200):", 
                parent=self.root,
                minvalue=10, 
                maxvalue=200,
                initialvalue=self.rows_per_page
            )
            
            if new_rows is not None:
                # Update rows per page setting
                self.rows_per_page = new_rows
                
                # Recalculate total pages
                if self.df is not None:
                    self.total_pages = max(1, math.ceil(len(self.df) / self.rows_per_page))
                    self.current_page = min(self.current_page, self.total_pages - 1)
                    
                    # Update data view
                    self.update_data_view()
                
                # Save preferences
                self.save_user_preferences()
                
        except Exception as e:
            messagebox.showerror("错误", f"设置每页行数时出错: {str(e)}")
            
# Tooltip class implementation for enhanced UX
class CreateToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        
    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        # Create a toplevel window
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = ttk.Label(self.tooltip, text=self.text, wraplength=180,
                       background="#ffffca", relief="solid", borderwidth=1)
        label.pack(padx=1, pady=1)
        
    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

if __name__ == "__main__":
    # 通过命令行参数控制是否启用拖放功能
    import sys
    enable_dnd = "--enable-dnd" in sys.argv
    
    # 使用普通的Tk作为后备选项
    root = None
    
    # 首先尝试使用ThemedTk
    try:
        root = ThemedTk(theme="arc")
        print("Successfully initialized ThemedTk")
    except Exception as e:
        print(f"Error initializing ThemedTk: {str(e)}")
        root = tk.Tk()
        print("Falling back to standard tkinter")
    
    # 如果启用了拖放功能，在创建界面前预加载tkinterdnd2
    if enable_dnd:
        try:
            import tkinterdnd2
            print("Successfully imported tkinterdnd2")
        except ImportError:
            print("tkinterdnd2 not available, drag and drop will be disabled")
        except Exception as e:
            print(f"Error importing tkinterdnd2: {str(e)}")
    
    # 创建应用实例
    app = DataVisualizationApp(root)
    root.mainloop() 