"""
GUI Application for Image Uploader
Provides a user-friendly interface for uploading and managing images
Includes intelligent sorting and grouping by identifier
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import os
from datetime import datetime
from image_uploader import ImageUploader, ImageIdentifier


class ImageUploaderGUI:
    """GUI Application for Image Uploader"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Image Uploader")
        self.root.geometry("1200x900")
        
        # Initialize uploader
        self.uploader = ImageUploader()
        
        # Store thumbnail images to prevent garbage collection
        self.thumbnail_images = {}
        
        # Track if preview is currently shown
        self.preview_visible = False
        
        # Store custom layout profiles
        self.layout_profiles = {}  # {group_label: layout_config}
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('aqua' if os.name == 'posix' else 'clam')
        
        # Setup GUI
        self.setup_ui()
        self.refresh_image_list()
        self.update_stats()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_container = ttk.Frame(self.root, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=1)
        main_container.rowconfigure(2, weight=1)  # Updated to row 2 for the main content area
        
        # Title
        title_label = ttk.Label(main_container, text="PowerPoint Image Uploader", 
                               font=('Helvetica', 20, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Actions Banner (horizontal toolbar across the top)
        actions_frame = ttk.Frame(main_container, padding="8")
        actions_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Add a subtle border and background
        actions_container = ttk.LabelFrame(actions_frame, text="Quick Actions", padding="10")
        actions_container.pack(fill=tk.X)
        
        # Create horizontal layout with organized button groups
        # Group 1: Data Management
        data_group = ttk.Frame(actions_container)
        data_group.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(data_group, text="🔄 Refresh", 
                  command=self.refresh_image_list, width=10).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(data_group, text="✏️ Edit Group", 
                  command=self.edit_group, width=11).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(data_group, text="🎨 Edit Type", 
                  command=self.edit_type, width=10).pack(side=tk.LEFT, padx=2)
        
        # Group 2: Creation Tools
        create_group = ttk.Frame(actions_container)
        create_group.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(create_group, text="🎨 Design Layout", 
                  command=self.design_layout, width=13).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(create_group, text="📊 Generate PPT", 
                  command=self.generate_powerpoint, width=13).pack(side=tk.LEFT, padx=2)
        
        # Group 3: File Operations
        file_group = ttk.Frame(actions_container)
        file_group.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(file_group, text="� Details", 
                  command=self.show_details, width=9).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(file_group, text="� Export", 
                  command=self.export_index, width=9).pack(side=tk.LEFT, padx=2)
        
        # Group 4: Deletion (separated for safety)
        delete_group = ttk.Frame(actions_container)
        delete_group.pack(side=tk.LEFT)
        
        ttk.Button(delete_group, text="�️ Remove", 
                  command=self.remove_selected, width=9).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(delete_group, text="🗑️ Clear All", 
                  command=self.clear_index, width=10).pack(side=tk.LEFT, padx=2)
        
        # Left Panel - Controls
        left_panel = ttk.Frame(main_container, padding="5")
        left_panel.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Upload Section
        upload_frame = ttk.LabelFrame(left_panel, text="Upload Images", padding="10")
        upload_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(upload_frame, text="📁 Select Images", 
                  command=self.select_images, width=25).grid(row=0, column=0, pady=5)
        
        ttk.Button(upload_frame, text="📂 Select Folder", 
                  command=self.select_folder, width=25).grid(row=1, column=0, pady=5)
        
        self.copy_files_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(upload_frame, text="Copy files to upload directory", 
                       variable=self.copy_files_var).grid(row=2, column=0, pady=5)
        
        # Statistics Section
        stats_frame = ttk.LabelFrame(left_panel, text="Statistics", padding="10")
        stats_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.stats_text = scrolledtext.ScrolledText(stats_frame, height=8, width=30, 
                                                    wrap=tk.WORD, state='disabled')
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Search Section
        search_frame = ttk.LabelFrame(left_panel, text="Search & Filter", padding="10")
        search_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Search by text
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        
        # Filter by group
        ttk.Label(search_frame, text="Group:").grid(row=2, column=0, sticky=tk.W, pady=(10, 2))
        self.group_filter_var = tk.StringVar(value="All Groups")
        self.group_combo = ttk.Combobox(search_frame, textvariable=self.group_filter_var, 
                                       state='readonly', width=27)
        self.group_combo['values'] = ["All Groups"]
        self.group_combo.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=2)
        self.group_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        
        # Filter by identifier
        ttk.Label(search_frame, text="Identifier:").grid(row=4, column=0, sticky=tk.W, pady=(10, 2))
        self.identifier_filter_var = tk.StringVar(value="All Types")
        self.identifier_combo = ttk.Combobox(search_frame, textvariable=self.identifier_filter_var,
                                            state='readonly', width=27)
        self.identifier_combo['values'] = ["All Types"] + ImageIdentifier.ALL_IDENTIFIERS
        self.identifier_combo.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=2)
        self.identifier_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        
        # Filter by format
        ttk.Label(search_frame, text="Format:").grid(row=6, column=0, sticky=tk.W, pady=(10, 2))
        self.format_filter_var = tk.StringVar(value="All Formats")
        self.format_combo = ttk.Combobox(search_frame, textvariable=self.format_filter_var,
                                        state='readonly', width=27)
        self.format_combo['values'] = ["All Formats"]
        self.format_combo.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=2)
        self.format_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        
        ttk.Button(search_frame, text="🔍 Apply Filters", 
                  command=self.apply_filters).grid(row=8, column=0, pady=(10, 5))
        
        ttk.Button(search_frame, text="Clear All Filters", 
                  command=self.clear_filters).grid(row=9, column=0, pady=5)
        
        search_frame.columnconfigure(0, weight=1)
        
        # Right Panel - Image List
        right_panel = ttk.Frame(main_container, padding="5")
        right_panel.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)
        
        # Image list label with selection instructions
        list_label = ttk.Label(right_panel, text="Uploaded Images", 
                              font=('Helvetica', 14, 'bold'))
        list_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        
        # Selection help text and status
        help_frame = ttk.Frame(right_panel)
        help_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(18, 3))
        help_frame.columnconfigure(0, weight=1)
        
        help_label = ttk.Label(help_frame, 
                              text="💡 Click & drag to select multiple images, Ctrl+Click to add/remove", 
                              font=('Helvetica', 9), foreground="gray")
        help_label.grid(row=0, column=0, sticky=tk.W)
        
        # Selection status label
        self.selection_status_label = ttk.Label(help_frame, text="", 
                                               font=('Helvetica', 9), foreground="blue")
        self.selection_status_label.grid(row=0, column=1, sticky=tk.E)
        
        # Treeview for images
        tree_frame = ttk.Frame(right_panel)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview with multi-select capability
        self.tree = ttk.Treeview(tree_frame, columns=("Group", "ID", "Identifier", "Filename", 
                                                      "Format", "Dimensions", "Size", "Date"),
                                show="headings", yscrollcommand=vsb.set, 
                                xscrollcommand=hsb.set, selectmode="extended")
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Configure columns
        self.tree.heading("Group", text="Group")
        self.tree.heading("ID", text="ID")
        self.tree.heading("Identifier", text="Type")
        self.tree.heading("Filename", text="Filename")
        self.tree.heading("Format", text="Format")
        self.tree.heading("Dimensions", text="Dimensions")
        self.tree.heading("Size", text="Size (KB)")
        self.tree.heading("Date", text="Date Added")
        
        self.tree.column("Group", width=70, anchor=tk.CENTER)
        self.tree.column("ID", width=50, anchor=tk.CENTER)
        self.tree.column("Identifier", width=80, anchor=tk.CENTER)
        self.tree.column("Filename", width=220)
        self.tree.column("Format", width=70, anchor=tk.CENTER)
        self.tree.column("Dimensions", width=110, anchor=tk.CENTER)
        self.tree.column("Size", width=90, anchor=tk.E)
        self.tree.column("Date", width=140)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Bind double-click event to show preview initially
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # Bind single-click to update preview if already visible
        self.tree.bind("<Button-1>", self.on_single_click)
        
        # Bind selection change to update preview when using arrow keys
        self.tree.bind("<<TreeviewSelect>>", self.on_selection_change)
        
        # Add drag-to-select functionality
        self.drag_start_y = None
        self.drag_start_item = None
        self.is_dragging = False
        
        # Bind drag selection events
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_end)
        
        # Preview panel with close button
        preview_frame = ttk.LabelFrame(right_panel, text="Preview", padding="10")
        preview_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Create a frame to hold the preview and close button
        preview_container = ttk.Frame(preview_frame)
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        self.preview_label = ttk.Label(preview_container, text="Double-click an image to preview")
        self.preview_label.pack()
        
        # Close button in top right - initially hidden
        self.preview_close_btn = ttk.Button(preview_container, text="Close Preview",
                                           command=self.clear_preview)
        self.preview_close_btn.pack(side=tk.TOP, anchor=tk.NE, padx=2, pady=2)
        self.preview_close_btn.pack_forget()  # Hide initially
    
    def select_images(self):
        """Open file dialog to select images"""
        filetypes = [
            ("All Images", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.tif *.webp"),
            ("JPEG Images", "*.jpg *.jpeg"),
            ("PNG Images", "*.png"),
            ("TIFF Images", "*.tiff *.tif"),
            ("All Files", "*.*")
        ]
        
        filepaths = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=filetypes
        )
        
        if filepaths:
            self.upload_images(list(filepaths))
    
    def select_folder(self):
        """Open dialog to select a folder containing images"""
        folder_path = filedialog.askdirectory(title="Select Folder with Images")
        
        if folder_path:
            # Find all image files in folder
            image_files = []
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    filepath = os.path.join(root, file)
                    if self.uploader.is_supported_format(filepath):
                        image_files.append(filepath)
            
            if image_files:
                self.upload_images(image_files)
            else:
                messagebox.showinfo("No Images", "No supported image files found in the selected folder.")
    
    def upload_images(self, filepaths):
        """Upload multiple images"""
        if not filepaths:
            return
        
        success_count = 0
        error_count = 0
        errors = []
        
        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Uploading Images")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        
        ttk.Label(progress_window, text="Uploading images...", 
                 font=('Helvetica', 12)).pack(pady=10)
        
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var, 
                                      maximum=len(filepaths), length=350)
        progress_bar.pack(pady=10)
        
        status_label = ttk.Label(progress_window, text="")
        status_label.pack(pady=5)
        
        # Upload images
        for i, filepath in enumerate(filepaths):
            try:
                filename = os.path.basename(filepath)
                status_label.config(text=f"Uploading: {filename}")
                progress_window.update()
                
                self.uploader.upload_image(filepath, 
                                          copy_to_upload_dir=self.copy_files_var.get())
                success_count += 1
            except Exception as e:
                error_count += 1
                errors.append(f"{os.path.basename(filepath)}: {str(e)}")
            
            progress_var.set(i + 1)
            progress_window.update()
        
        progress_window.destroy()
        
        # Show results
        message = f"Successfully uploaded {success_count} image(s)."
        if error_count > 0:
            message += f"\n{error_count} error(s) occurred."
            if errors:
                error_details = "\n".join(errors[:5])
                if len(errors) > 5:
                    error_details += f"\n... and {len(errors) - 5} more errors"
                message += f"\n\nErrors:\n{error_details}"
        
        messagebox.showinfo("Upload Complete", message)
        
        # Refresh display
        self.refresh_image_list()
        self.update_stats()
    
    def refresh_image_list(self):
        """Refresh the image list display with grouping and sorting"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get sorted images
        images = self.uploader.list_images()
        
        # Update group filter dropdown with all formatted labels that appear in the data
        from image_uploader import ImageIdentifier
        
        # Collect all unique group labels that actually appear in the image list
        group_labels = set()
        for img in images:
            metadata = img.get('metadata', {})
            numerical_prefix = metadata.get('numerical_prefix')
            identifier = metadata.get('identifier')
            
            if numerical_prefix:  # Only add if image has a group
                formatted_label = ImageIdentifier.format_group_label(numerical_prefix, identifier)
                if formatted_label:
                    group_labels.add(formatted_label)
        
        # Sort the group labels
        sorted_group_labels = sorted(group_labels)
        
        group_values = ["All Groups"] + sorted_group_labels
        if sorted_group_labels:  # Only add "Ungrouped" if there are any groups
            group_values.append("Ungrouped")
        self.group_combo['values'] = group_values
        
        # Update identifier filter dropdown with all types that actually appear in the data
        identifier_types = set()
        for img in images:
            identifier = img.get('metadata', {}).get('identifier')
            if identifier:
                identifier_types.add(identifier)
        
        sorted_identifiers = sorted(identifier_types)
        self.identifier_combo['values'] = ["All Types"] + sorted_identifiers
        
        # Update format filter dropdown with all formats that actually appear in the data
        format_types = set()
        for img in images:
            img_format = img.get('format')
            if img_format:
                format_types.add(img_format)
        
        sorted_formats = sorted(format_types)
        self.format_combo['values'] = ["All Formats"] + sorted_formats
        
        # Populate tree
        for img in images:
            metadata = img.get('metadata', {})
            numerical_prefix = metadata.get('numerical_prefix')
            identifier = metadata.get('identifier', 'N/A')
            
            # Format group label (MAP1, MAP2 for map groups, 0001, 0002 for others)
            from image_uploader import ImageIdentifier
            group_label = ImageIdentifier.format_group_label(numerical_prefix, identifier) or 'N/A'
            
            dimensions = f"{img['width']}x{img['height']}"
            size_kb = round(img['size_bytes'] / 1024, 2)
            date_added = datetime.fromisoformat(img['added_date']).strftime("%Y-%m-%d %H:%M")
            
            self.tree.insert("", tk.END, values=(
                group_label,
                img['id'],
                identifier,
                img['filename'],
                img['format'],
                dimensions,
                size_kb,
                date_added
            ))
    
    def update_stats(self):
        """Update statistics display"""
        stats = self.uploader.get_statistics()
        
        stats_text = f"Total Images: {stats['total_images']}\n"
        stats_text += f"Total Size: {stats['total_size_mb']} MB\n\n"
        
        # Add group information
        groups = self.uploader.index.get_all_groups()
        if groups:
            stats_text += f"Groups: {len(groups)}\n\n"
        
        stats_text += "Formats:\n"
        for format_type, count in stats['formats'].items():
            stats_text += f"  {format_type}: {count}\n"
        
        self.stats_text.config(state='normal')
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, stats_text)
        self.stats_text.config(state='disabled')
    
    def apply_filters(self):
        """Apply search query and filters"""
        from image_uploader import ImageIdentifier
        
        query = self.search_var.get().strip()
        group_filter = self.group_filter_var.get()
        identifier_filter = self.identifier_filter_var.get()
        format_filter = self.format_filter_var.get()
        
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get all images
        images = self.uploader.list_images()
        
        # Apply filters
        filtered_images = []
        for img in images:
            metadata = img.get('metadata', {})
            numerical_prefix = metadata.get('numerical_prefix', '')
            identifier = metadata.get('identifier', '')
            img_format = img.get('format', '')
            
            # Format group label for comparison
            group_label = ImageIdentifier.format_group_label(numerical_prefix, identifier)
            
            # Apply group filter
            if group_filter != "All Groups":
                if group_filter == "Ungrouped" and numerical_prefix:
                    continue
                elif group_filter != "Ungrouped" and group_label != group_filter:
                    continue
            
            # Apply identifier filter
            if identifier_filter != "All Types" and identifier != identifier_filter:
                continue
            
            # Apply format filter
            if format_filter != "All Formats" and img_format != format_filter:
                continue
            
            # Apply text search
            if query:
                query_lower = query.lower()
                if not (query_lower in img['filename'].lower() or 
                       query_lower in str(metadata).lower()):
                    continue
            
            filtered_images.append(img)
        
        # Display filtered results
        for img in filtered_images:
            metadata = img.get('metadata', {})
            numerical_prefix = metadata.get('numerical_prefix')
            identifier = metadata.get('identifier', 'N/A')
            
            # Format group label
            group_label = ImageIdentifier.format_group_label(numerical_prefix, identifier) or 'N/A'
            
            dimensions = f"{img['width']}x{img['height']}"
            size_kb = round(img['size_bytes'] / 1024, 2)
            date_added = datetime.fromisoformat(img['added_date']).strftime("%Y-%m-%d %H:%M")
            
            self.tree.insert("", tk.END, values=(
                group_label,
                img['id'],
                identifier,
                img['filename'],
                img['format'],
                dimensions,
                size_kb,
                date_added
            ))
    
    def clear_filters(self):
        """Clear all filters and show all images"""
        self.search_var.set("")
        self.group_filter_var.set("All Groups")
        self.identifier_filter_var.set("All Types")
        self.format_filter_var.set("All Formats")
        self.refresh_image_list()
    
    def edit_group(self):
        """Allow user to manually edit the group assignment of selected image(s)"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select one or more images to edit.")
            return
        
        # Get selected image IDs
        selected_ids = []
        for item in selection:
            values = self.tree.item(item)['values']
            selected_ids.append(int(values[1]))  # ID is column index 1
        
        # Create edit dialog
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Group Assignment")
        edit_window.geometry("550x350")
        edit_window.transient(self.root)
        
        # Main frame
        main_frame = ttk.Frame(edit_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text=f"Editing {len(selected_ids)} image(s)", 
                 font=('Helvetica', 12, 'bold')).pack(pady=(0, 10))
        
        # Get identifier type of first selected image to provide better guidance
        first_img = self.uploader.index.get_image(selected_ids[0])
        identifier = first_img.get('metadata', {}).get('identifier', '') if first_img else ''
        
        from image_uploader import ImageIdentifier
        
        # Group label input
        group_frame = ttk.Frame(main_frame)
        group_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(group_frame, text="New Group Label:").pack(anchor='w', pady=(0, 5))
        
        entry_frame = ttk.Frame(group_frame)
        entry_frame.pack(fill=tk.X)
        
        group_var = tk.StringVar()
        group_entry = ttk.Entry(entry_frame, textvariable=group_var, width=30)
        group_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(entry_frame, text="(or leave blank to remove group)", foreground="gray").pack(side=tk.LEFT)
        
        # Info text
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=10)
        
        # Determine example text based on identifier type
        if identifier == 'Spectrum' or identifier in ImageIdentifier.SPECTRUM_IDENTIFIERS:
            auto_format = "SPEC1, SPEC2, etc."
        elif identifier == 'Maps' or identifier in ['Map', 'Maps', 'Electron Image']:
            auto_format = "MAP1, MAP2, etc."
        else:
            auto_format = "0001, 0002, etc."
        
        info_text = (f"Examples:\n"
                    f"  • Numeric: 0001, 0002, 1, 2 → Auto-formatted as {auto_format}\n"
                    f"  • Formatted: MAP1, SPEC1 → Stored with extracted number\n"
                    f"  • Custom: MyGroup, Sample-A, Test_1 → Used as-is\n\n"
                    f"Note: Numeric values will be auto-formatted based on identifier type.\n"
                    f"Custom labels allow you to create unique grouping schemes.")
        
        ttk.Label(info_frame, text=info_text, foreground="gray", justify=tk.LEFT).pack(anchor='w')
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        def apply_changes():
            new_group_input = group_var.get().strip()
            
            if not new_group_input:
                # Blank = remove group
                new_group = None
            else:
                # Accept any input - parse for numeric or use as-is
                import re
                
                # Try to extract number from formatted labels like MAP1, SPEC1, etc.
                formatted_match = re.match(r'(?:MAP|SPEC|map|spec)(\d+)$', new_group_input, re.IGNORECASE)
                if formatted_match:
                    # Extract the number and pad with zeros (for auto-formatting)
                    new_group = formatted_match.group(1).zfill(4)
                elif new_group_input.isdigit():
                    # Pure numeric input - pad with zeros (for auto-formatting)
                    new_group = new_group_input.zfill(4)
                else:
                    # Custom label - use as-is (alphanumeric, dashes, underscores allowed)
                    # Validate that it's a reasonable label
                    if len(new_group_input) > 50:
                        messagebox.showerror("Invalid Input", "Group label is too long (max 50 characters).")
                        return
                    new_group = new_group_input
            
            # Update each selected image
            for img_id in selected_ids:
                img = self.uploader.index.get_image(img_id)
                if img:
                    img['metadata']['numerical_prefix'] = new_group
            
            # Save changes
            self.uploader.index.save_index()
            
            # Refresh display
            self.refresh_image_list()
            self.update_stats()
            
            messagebox.showinfo("Success", f"Updated group for {len(selected_ids)} image(s) to '{new_group if new_group else 'No Group'}'.")
            edit_window.destroy()
        
        ttk.Button(button_frame, text="Apply", command=apply_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=edit_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # Focus on entry
        group_entry.focus()
    
    def edit_type(self):
        """Allow user to manually edit the image type/identifier of selected image(s)"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select one or more images to edit.")
            return
        
        # Get selected image IDs
        selected_ids = []
        for item in selection:
            values = self.tree.item(item)['values']
            selected_ids.append(int(values[1]))  # ID is column index 1
        
        # Create edit dialog
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Image Type")
        edit_window.geometry("500x400")
        edit_window.transient(self.root)
        
        # Main frame
        main_frame = ttk.Frame(edit_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text=f"Editing {len(selected_ids)} image(s)", 
                 font=('Helvetica', 12, 'bold')).pack(pady=(0, 10))
        
        from image_uploader import ImageIdentifier
        
        # Type selection
        type_frame = ttk.Frame(main_frame)
        type_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(type_frame, text="New Image Type:").pack(anchor='w', pady=(0, 5))
        
        type_var = tk.StringVar()
        type_combo = ttk.Combobox(type_frame, textvariable=type_var, 
                                 values=ImageIdentifier.ALL_IDENTIFIERS, 
                                 state='readonly', width=40)
        type_combo.pack(fill=tk.X)
        
        # Set current type if all selected images have the same type
        current_types = set()
        for img_id in selected_ids:
            img = self.uploader.index.get_image(img_id)
            if img:
                current_type = img.get('metadata', {}).get('identifier', '')
                if current_type:
                    current_types.add(current_type)
        
        if len(current_types) == 1:
            type_var.set(list(current_types)[0])
        
        # Info text
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=10)
        
        info_text = ("Available image types:\n\n"
                    "• Spectrum: For spectroscopy data\n"
                    "• Map/Maps: For elemental maps\n"
                    "• Electron Image: For SEM/TEM images\n"
                    "• HAADF: High-angle annular dark field\n"
                    "• BF: Bright field\n"
                    "• DF: Dark field\n"
                    "• And more...\n\n"
                    "The image type affects how images are grouped and\n"
                    "formatted in PowerPoint presentations.")
        
        ttk.Label(info_frame, text=info_text, foreground="gray", justify=tk.LEFT).pack(anchor='w')
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        def apply_changes():
            new_type = type_var.get().strip()
            
            if not new_type:
                messagebox.showwarning("No Type Selected", "Please select an image type.")
                return
            
            if new_type not in ImageIdentifier.ALL_IDENTIFIERS:
                messagebox.showerror("Invalid Type", "Selected type is not valid.")
                return
            
            # Update each selected image
            for img_id in selected_ids:
                img = self.uploader.index.get_image(img_id)
                if img:
                    img['metadata']['identifier'] = new_type
            
            # Save changes
            self.uploader.index.save_index()
            
            # Refresh display
            self.refresh_image_list()
            self.update_stats()
            
            messagebox.showinfo("Success", f"Updated image type for {len(selected_ids)} image(s) to '{new_type}'.")
            edit_window.destroy()
        
        ttk.Button(button_frame, text="Apply", command=apply_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=edit_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # Focus on combobox
        type_combo.focus()
    
    def clear_preview(self):
        """Clear the preview panel"""
        self.preview_label.config(text="Double-click an image to preview", image="")
        self.preview_close_btn.pack_forget()  # Hide the close button
        self.preview_visible = False  # Mark preview as hidden
    
    def on_single_click(self, event):
        """Handle single-click to update preview if already visible"""
        # Only update preview if it's already visible
        if not self.preview_visible:
            return
        
        # Small delay to allow tree selection to update
        self.root.after(10, self._update_preview_from_click)
    
    def on_selection_change(self, event):
        """Handle selection change (from arrow keys or any other navigation)"""
        # Update selection status
        selection = self.tree.selection()
        if len(selection) > 1:
            self.selection_status_label.config(text=f"{len(selection)} images selected")
        else:
            self.selection_status_label.config(text="")
        
        # Only update preview if it's already visible
        if not self.preview_visible:
            return
        
        # Update preview with small delay
        self.root.after(10, self._update_preview_from_click)
    
    def _update_preview_from_click(self):
        """Update preview after selection change"""
        selection = self.tree.selection()
        if not selection:
            return
        
        # Get selected image ID
        item = self.tree.item(selection[0])
        image_id = int(item['values'][1])  # ID is in column index 1
        
        # Get image info
        img_info = self.uploader.get_image_info(image_id)
        if not img_info:
            return
        
        # Load and display thumbnail
        try:
            filepath = img_info['filepath']
            if os.path.exists(filepath):
                image = Image.open(filepath)
                
                # Create thumbnail
                image.thumbnail((300, 300), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(image)
                
                # Store reference and update label
                self.thumbnail_images[image_id] = photo
                self.preview_label.config(image=photo, text="")
            else:
                self.preview_label.config(text="Image file not found", image="")
        except Exception as e:
            self.preview_label.config(text=f"Error loading preview: {str(e)}", image="")
    
    def on_double_click(self, event):
        """Handle image double-click to show preview"""
        selection = self.tree.selection()
        if not selection:
            return
        
        # Get selected image ID (now in column index 1)
        item = self.tree.item(selection[0])
        image_id = int(item['values'][1])  # ID is now the second column
        
        # Get image info
        img_info = self.uploader.get_image_info(image_id)
        if not img_info:
            return
        
        # Load and display thumbnail
        try:
            filepath = img_info['filepath']
            if os.path.exists(filepath):
                image = Image.open(filepath)
                
                # Create thumbnail
                image.thumbnail((300, 300), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(image)
                
                # Store reference and update label
                self.thumbnail_images[image_id] = photo
                self.preview_label.config(image=photo, text="")
                
                # Show the close button and mark preview as visible
                self.preview_close_btn.pack(side=tk.TOP, anchor=tk.NE, padx=2, pady=2)
                self.preview_visible = True
            else:
                self.preview_label.config(text="Image file not found", image="")
                self.preview_close_btn.pack_forget()
                self.preview_visible = False
        except Exception as e:
            self.preview_label.config(text=f"Error loading preview: {str(e)}", image="")
            self.preview_close_btn.pack_forget()
            self.preview_visible = False
    
    def on_drag_start(self, event):
        """Handle the start of a drag selection"""
        # Get the item at the click position
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_start_item = item
            self.drag_start_y = event.y
            self.is_dragging = False  # Not dragging yet, just clicked
            
            # Check if Ctrl or Cmd is held for multi-select
            if event.state & 0x4:  # Control key held - don't change selection yet
                # Let on_drag_end handle the toggle logic
                pass
            else:  # Control key not held
                # Clear current selection and select this item
                self.tree.selection_set(item)
        
        # Call the original single-click handler only if Ctrl is not held
        if not (event.state & 0x4):
            self.on_single_click(event)
    
    def on_drag_motion(self, event):
        """Handle drag motion to select multiple items"""
        if self.drag_start_item and self.drag_start_y is not None:
            # Start dragging if we've moved enough
            if not self.is_dragging and abs(event.y - self.drag_start_y) > 5:
                self.is_dragging = True
            
            if self.is_dragging:
                # Get all items between start and current position
                start_y = self.drag_start_y
                end_y = event.y
                
                # Ensure start_y is always less than end_y
                if start_y > end_y:
                    start_y, end_y = end_y, start_y
                
                # Get all visible items in the tree
                children = self.tree.get_children()
                selected_items = []
                
                for child in children:
                    # Get the bounding box of the item
                    try:
                        bbox = self.tree.bbox(child)
                        if bbox:  # bbox returns None if item is not visible
                            item_y = bbox[1] + bbox[3] / 2  # Middle of the item
                            # Check if item is within selection range
                            if start_y <= item_y <= end_y:
                                selected_items.append(child)
                    except:
                        # If bbox fails, skip this item
                        continue
                
                # Update selection
                if selected_items:
                    self.tree.selection_set(selected_items)
    
    def on_drag_end(self, event):
        """Handle the end of drag selection"""
        if self.is_dragging:
            # Dragging completed, selection is already set in on_drag_motion
            pass
        else:
            # Just a click, not a drag - handle normal click selection
            item = self.tree.identify_row(event.y)
            if item:
                # Check if Ctrl/Cmd is held for adding to selection
                if event.state & 0x4:  # Control key held
                    # Toggle selection of this item
                    current_selection = self.tree.selection()
                    if item in current_selection:
                        # Remove from selection
                        remaining = [i for i in current_selection if i != item]
                        self.tree.selection_set(remaining)
                    else:
                        # Add to selection
                        self.tree.selection_add(item)
                else:
                    # Normal click - select just this item
                    self.tree.selection_set(item)
        
        # Update selection status
        selection = self.tree.selection()
        if len(selection) > 1:
            self.selection_status_label.config(text=f"{len(selection)} images selected")
        else:
            self.selection_status_label.config(text="")
        
        # Reset drag state
        self.drag_start_item = None
        self.drag_start_y = None
        self.is_dragging = False
    
    def show_details(self):
        """Show detailed information about selected image(s)"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select one or more images first.")
            return
        
        # Create details window
        details_window = tk.Toplevel(self.root)
        if len(selection) == 1:
            # Single selection - get specific info
            item = self.tree.item(selection[0])
            image_id = int(item['values'][1])
            img_info = self.uploader.get_image_info(image_id)
            if not img_info:
                messagebox.showerror("Error", "Image information not found.")
                return
            details_window.title(f"Image Details - {img_info['filename']}")
        else:
            # Multiple selection
            details_window.title(f"Details for {len(selection)} Selected Images")
        
        details_window.geometry("600x500")
        details_window.transient(self.root)
        
        # Details text
        details_frame = ttk.Frame(details_window, padding="20")
        details_frame.pack(fill=tk.BOTH, expand=True)
        
        details_text = scrolledtext.ScrolledText(details_frame, wrap=tk.WORD, 
                                                 font=('Courier', 10))
        details_text.pack(fill=tk.BOTH, expand=True)
        
        # Format details for single or multiple images
        if len(selection) == 1:
            # Single image details (existing behavior)
            item = self.tree.item(selection[0])
            image_id = int(item['values'][1])
            img_info = self.uploader.get_image_info(image_id)
            
            details_str = f"Image Details\n{'='*50}\n\n"
            for key, value in img_info.items():
                if key == 'metadata':
                    details_str += f"\nMetadata:\n"
                    if value:
                        for mk, mv in value.items():
                            details_str += f"  {mk.replace('_', ' ').title()}: {mv}\n"
                    else:
                        details_str += "  None\n"
                else:
                    details_str += f"{key.replace('_', ' ').title()}: {value}\n"
        else:
            # Multiple images summary
            details_str = f"Summary of {len(selection)} Selected Images\n{'='*50}\n\n"
            
            # Collect statistics
            total_size = 0
            formats = {}
            identifiers = {}
            groups = {}
            
            for item in selection:
                values = self.tree.item(item)['values']
                image_id = int(values[1])
                img_info = self.uploader.get_image_info(image_id)
                
                if img_info:
                    # Size
                    total_size += img_info['size_bytes']
                    
                    # Format
                    fmt = img_info['format']
                    formats[fmt] = formats.get(fmt, 0) + 1
                    
                    # Identifier
                    identifier = img_info.get('metadata', {}).get('identifier', 'N/A')
                    identifiers[identifier] = identifiers.get(identifier, 0) + 1
                    
                    # Group
                    group = values[0]  # Group column
                    groups[group] = groups.get(group, 0) + 1
                    
                    # Add individual file info
                    details_str += f"• {img_info['filename']}\n"
                    details_str += f"  Type: {identifier}, Group: {group}\n"
                    details_str += f"  Size: {round(img_info['size_bytes']/1024, 2)} KB\n\n"
            
            # Add summary statistics
            details_str += f"\nSummary Statistics:\n{'-'*30}\n"
            details_str += f"Total Files: {len(selection)}\n"
            details_str += f"Total Size: {round(total_size/1024/1024, 2)} MB\n\n"
            
            details_str += "Formats:\n"
            for fmt, count in sorted(formats.items()):
                details_str += f"  {fmt}: {count}\n"
            
            details_str += "\nImage Types:\n"
            for identifier, count in sorted(identifiers.items()):
                details_str += f"  {identifier}: {count}\n"
            
            details_str += "\nGroups:\n"
            for group, count in sorted(groups.items()):
                details_str += f"  {group}: {count}\n"
        
        details_text.insert(1.0, details_str)
        details_text.config(state='disabled')
        
        # Close button
        ttk.Button(details_window, text="Close", 
                  command=details_window.destroy).pack(pady=10)
    
    def remove_selected(self):
        """Remove selected image(s) from index"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select one or more images to remove.")
            return
        
        # Get selected image IDs and filenames
        selected_items = []
        for item in selection:
            values = self.tree.item(item)['values']
            image_id = int(values[1])  # ID is column index 1
            filename = values[3]  # Filename is column index 3
            selected_items.append((image_id, filename))
        
        # Confirm deletion
        if len(selected_items) == 1:
            result = messagebox.askyesno("Confirm Removal", 
                                         f"Remove '{selected_items[0][1]}' from index?\n\n"
                                         "The file will not be deleted from disk.")
        else:
            result = messagebox.askyesno("Confirm Removal", 
                                         f"Remove {len(selected_items)} images from index?\n\n"
                                         "The files will not be deleted from disk.")
        
        if result:
            success_count = 0
            for image_id, filename in selected_items:
                if self.uploader.remove_image(image_id):
                    success_count += 1
            
            if success_count == len(selected_items):
                if len(selected_items) == 1:
                    messagebox.showinfo("Success", "Image removed from index.")
                else:
                    messagebox.showinfo("Success", f"{success_count} images removed from index.")
            else:
                messagebox.showwarning("Partial Success", 
                                     f"{success_count} of {len(selected_items)} images removed from index.")
            
            self.refresh_image_list()
            self.update_stats()
            self.clear_preview()
    
    def export_index(self):
        """Export index to a text file"""
        filepath = filedialog.asksaveasfilename(
            title="Export Index",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if filepath:
            try:
                images = self.uploader.list_images()
                
                with open(filepath, 'w') as f:
                    f.write("Image Index Export\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("="*70 + "\n\n")
                    
                    for img in images:
                        f.write(f"ID: {img['id']}\n")
                        f.write(f"Filename: {img['filename']}\n")
                        f.write(f"Path: {img['filepath']}\n")
                        f.write(f"Format: {img['format']}\n")
                        f.write(f"Dimensions: {img['width']}x{img['height']}\n")
                        f.write(f"Size: {round(img['size_bytes']/1024, 2)} KB\n")
                        f.write(f"Added: {img['added_date']}\n")
                        f.write("-"*70 + "\n\n")
                
                messagebox.showinfo("Success", f"Index exported to:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export index:\n{str(e)}")
    
    def clear_index(self):
        """Clear all images from the index with double confirmation"""
        # First confirmation
        result1 = messagebox.askyesno(
            "Clear Index - Confirmation",
            "⚠️ WARNING ⚠️\n\n"
            "You are about to clear ALL images from the index.\n\n"
            "This will:\n"
            "• Remove all indexed images\n"
            "• Delete all grouping information\n"
            "• Clear all custom layouts\n\n"
            "Image files on disk will NOT be deleted.\n\n"
            "Do you want to continue?",
            icon='warning'
        )
        
        if not result1:
            return
        
        # Second confirmation (failsafe)
        images = self.uploader.list_images()
        num_images = len(images)
        
        result2 = messagebox.askyesno(
            "Clear Index - Final Confirmation",
            f"⚠️ FINAL WARNING ⚠️\n\n"
            f"You are about to permanently remove {num_images} image(s) from the index.\n\n"
            f"This action CANNOT be undone!\n\n"
            f"Are you absolutely sure you want to proceed?",
            icon='warning'
        )
        
        if not result2:
            messagebox.showinfo("Cancelled", "Index clear operation cancelled.")
            return
        
        # Clear the index
        try:
            # Clear the JSON file by saving empty list
            self.uploader.index.images = []
            self.uploader.index.save_index()
            
            # Clear layout profiles
            self.layout_profiles.clear()
            
            # Refresh the display
            self.refresh_image_list()
            self.update_stats()
            self.clear_preview()
            
            messagebox.showinfo("Success", 
                              f"Index cleared successfully!\n"
                              f"{num_images} image(s) removed from index.\n\n"
                              f"Image files on disk remain unchanged.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear index:\n{str(e)}")
    
    def design_layout(self):
        """Open layout designer for custom slide layouts"""
        from image_uploader import ImageIdentifier
        
        # Get all groups
        images = self.uploader.list_images()
        grouped_images = [img for img in images if img.get('metadata', {}).get('numerical_prefix')]
        
        if not grouped_images:
            messagebox.showwarning("No Groups", "No grouped images found. Please create groups first.")
            return
        
        # Get unique groups with their identifiers
        groups_dict = {}
        for img in grouped_images:
            metadata = img.get('metadata', {})
            numerical_prefix = metadata.get('numerical_prefix')
            identifier = metadata.get('identifier')
            
            if numerical_prefix:
                group_label = ImageIdentifier.format_group_label(numerical_prefix, identifier)
                if group_label not in groups_dict:
                    groups_dict[group_label] = []
                groups_dict[group_label].append(img)
        
        # Create designer window
        designer_window = tk.Toplevel(self.root)
        designer_window.title("Layout Designer")
        designer_window.geometry("1000x800")
        designer_window.transient(self.root)
        
        # Main frame
        main_frame = ttk.Frame(designer_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Custom Slide Layout Designer", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Design custom layouts by drawing regions on the canvas").pack(pady=(0, 10))
        
        # Create horizontal split: left for controls, right for canvas
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        left_controls = ttk.Frame(content_frame)
        left_controls.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        right_canvas_frame = ttk.Frame(content_frame)
        right_canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Group selection
        group_frame = ttk.LabelFrame(left_controls, text="Select Group", padding="10")
        group_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(group_frame, text="Group:").pack(anchor='w', pady=(0, 5))
        
        selected_group = tk.StringVar()
        group_combo = ttk.Combobox(group_frame, textvariable=selected_group, 
                                   values=sorted(groups_dict.keys()), width=20, state='readonly')
        group_combo.pack(fill=tk.X)
        
        group_info_label = ttk.Label(group_frame, text="", foreground="gray", wraplength=200)
        group_info_label.pack(pady=(5, 0))
        
        # Region tools
        tools_frame = ttk.LabelFrame(left_controls, text="Region Tools", padding="10")
        tools_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(tools_frame, text="Click and drag on canvas\nto create regions", 
                 foreground="gray").pack(pady=(0, 10))
        
        # Quick layouts
        quick_frame = ttk.LabelFrame(left_controls, text="Quick Layouts", padding="10")
        quick_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Region list
        regions_frame = ttk.LabelFrame(left_controls, text="Regions", padding="10")
        regions_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        regions_listbox = tk.Listbox(regions_frame, height=8)
        regions_listbox.pack(fill=tk.BOTH, expand=True)
        
        regions_buttons = ttk.Frame(regions_frame)
        regions_buttons.pack(fill=tk.X, pady=(5, 0))
        
        # Canvas for visual layout design
        canvas_container = ttk.LabelFrame(right_canvas_frame, text="Slide Layout (16:9)", padding="10")
        canvas_container.pack(fill=tk.BOTH, expand=True)
        
        # Canvas representing slide (aspect ratio 16:9)
        canvas_width = 640
        canvas_height = 360
        
        layout_canvas = tk.Canvas(canvas_container, width=canvas_width, height=canvas_height, 
                                  bg='white', relief='solid', borderwidth=1)
        layout_canvas.pack(pady=10)
        
        # Store regions and drawing state
        canvas_regions = []  # List of {id, rect_id, x1, y1, x2, y2, identifier, color}
        current_rect = None
        start_x = None
        start_y = None
        region_colors = ['#FFB3BA', '#BAFFC9', '#BAE1FF', '#FFFFBA', '#FFD9BA', '#E0BBE4']
        color_index = [0]
        
        def start_draw(event):
            nonlocal start_x, start_y, current_rect
            start_x = event.x
            start_y = event.y
            current_rect = layout_canvas.create_rectangle(start_x, start_y, start_x, start_y, 
                                                          outline='black', width=2, fill='', stipple='gray50')
        
        def draw_rect(event):
            nonlocal current_rect
            if current_rect:
                layout_canvas.coords(current_rect, start_x, start_y, event.x, event.y)
        
        def end_draw(event):
            nonlocal current_rect, start_x, start_y
            if current_rect and start_x and start_y:
                x1, y1 = min(start_x, event.x), min(start_y, event.y)
                x2, y2 = max(start_x, event.x), max(start_y, event.y)
                
                # Minimum size check
                if abs(x2 - x1) < 20 or abs(y2 - y1) < 20:
                    layout_canvas.delete(current_rect)
                    current_rect = None
                    return
                
                # Get identifier from selected group
                group = selected_group.get()
                identifiers = []
                if group and group in groups_dict:
                    identifiers = sorted(set(img.get('metadata', {}).get('identifier', 'Unknown') 
                                           for img in groups_dict[group]))
                
                # Prompt for identifier
                if identifiers:
                    # Create simple selection dialog
                    identifier = identifiers[0] if identifiers else "Image"
                    if len(identifiers) > 1:
                        # Simple dialog to pick identifier
                        choice_win = tk.Toplevel(designer_window)
                        choice_win.title("Select Image Type")
                        choice_win.geometry("300x200")
                        choice_win.transient(designer_window)
                        
                        ttk.Label(choice_win, text="Assign this region to:", 
                                 font=('Helvetica', 10, 'bold')).pack(pady=10)
                        
                        selected_id = tk.StringVar(value=identifiers[0])
                        
                        for id_type in identifiers:
                            ttk.Radiobutton(choice_win, text=id_type, variable=selected_id, 
                                          value=id_type).pack(anchor='w', padx=20, pady=2)
                        
                        def confirm():
                            choice_win.destroy()
                        
                        ttk.Button(choice_win, text="OK", command=confirm).pack(pady=10)
                        choice_win.wait_window()
                        identifier = selected_id.get()
                else:
                    identifier = "Image"
                
                # Color the region
                color = region_colors[color_index[0] % len(region_colors)]
                color_index[0] += 1
                
                layout_canvas.itemconfig(current_rect, fill=color, stipple='')
                
                # Add label
                text_id = layout_canvas.create_text((x1 + x2) / 2, (y1 + y2) / 2, 
                                                   text=identifier, font=('Helvetica', 10, 'bold'))
                
                # Store region
                region_id = len(canvas_regions) + 1
                canvas_regions.append({
                    'id': region_id,
                    'rect_id': current_rect,
                    'text_id': text_id,
                    'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                    'identifier': identifier,
                    'color': color
                })
                
                update_regions_list()
                current_rect = None
        
        layout_canvas.bind('<Button-1>', start_draw)
        layout_canvas.bind('<B1-Motion>', draw_rect)
        layout_canvas.bind('<ButtonRelease-1>', end_draw)
        
        def update_regions_list():
            regions_listbox.delete(0, tk.END)
            for region in canvas_regions:
                width_pct = int((region['x2'] - region['x1']) / canvas_width * 100)
                height_pct = int((region['y2'] - region['y1']) / canvas_height * 100)
                regions_listbox.insert(tk.END, 
                    f"#{region['id']}: {region['identifier']} ({width_pct}% × {height_pct}%)")
        
        def delete_selected_region():
            selection = regions_listbox.curselection()
            if selection:
                idx = selection[0]
                region = canvas_regions[idx]
                layout_canvas.delete(region['rect_id'])
                layout_canvas.delete(region['text_id'])
                canvas_regions.pop(idx)
                update_regions_list()
        
        def clear_all_regions():
            for region in canvas_regions:
                layout_canvas.delete(region['rect_id'])
                layout_canvas.delete(region['text_id'])
            canvas_regions.clear()
            update_regions_list()
        
        def apply_quick_layout(layout_type):
            clear_all_regions()
            
            group = selected_group.get()
            if not group or group not in groups_dict:
                return
            
            identifiers = sorted(set(img.get('metadata', {}).get('identifier', 'Unknown') 
                                   for img in groups_dict[group]))
            
            if layout_type == 'split_lr' and len(identifiers) >= 2:
                # Left-right split
                create_region(0, 0, canvas_width//2, canvas_height, identifiers[0])
                create_region(canvas_width//2, 0, canvas_width, canvas_height, identifiers[1])
            elif layout_type == 'split_tb' and len(identifiers) >= 2:
                # Top-bottom split
                create_region(0, 0, canvas_width, canvas_height//2, identifiers[0])
                create_region(0, canvas_height//2, canvas_width, canvas_height, identifiers[1])
            elif layout_type == 'thirds':
                # Left 1/3, right 2/3
                split = canvas_width // 3
                if len(identifiers) >= 2:
                    create_region(0, 0, split, canvas_height, identifiers[0])
                    create_region(split, 0, canvas_width, canvas_height, identifiers[1])
                else:
                    create_region(0, 0, split, canvas_height, identifiers[0] if identifiers else "Image")
                    create_region(split, 0, canvas_width, canvas_height, identifiers[0] if identifiers else "Image")
        
        def create_region(x1, y1, x2, y2, identifier):
            color = region_colors[color_index[0] % len(region_colors)]
            color_index[0] += 1
            
            rect_id = layout_canvas.create_rectangle(x1, y1, x2, y2, 
                                                    fill=color, outline='black', width=2)
            text_id = layout_canvas.create_text((x1 + x2) / 2, (y1 + y2) / 2, 
                                               text=identifier, font=('Helvetica', 10, 'bold'))
            
            region_id = len(canvas_regions) + 1
            canvas_regions.append({
                'id': region_id,
                'rect_id': rect_id,
                'text_id': text_id,
                'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                'identifier': identifier,
                'color': color
            })
            update_regions_list()
        
        ttk.Button(quick_frame, text="Split Left/Right", 
                  command=lambda: apply_quick_layout('split_lr'), width=20).pack(fill=tk.X, pady=2)
        ttk.Button(quick_frame, text="Split Top/Bottom", 
                  command=lambda: apply_quick_layout('split_tb'), width=20).pack(fill=tk.X, pady=2)
        ttk.Button(quick_frame, text="1/3 - 2/3 Split", 
                  command=lambda: apply_quick_layout('thirds'), width=20).pack(fill=tk.X, pady=2)
        
        ttk.Button(regions_buttons, text="Delete", command=delete_selected_region, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(regions_buttons, text="Clear All", command=clear_all_regions, width=10).pack(side=tk.LEFT, padx=2)
        
        def update_group_info(*args):
            """Update info when group is selected"""
            group = selected_group.get()
            if group and group in groups_dict:
                imgs = groups_dict[group]
                identifiers = set(img.get('metadata', {}).get('identifier', 'Unknown') for img in imgs)
                group_info_label.config(text=f"{len(imgs)} images\nTypes: {', '.join(sorted(identifiers))}")
        
        group_combo.bind('<<ComboboxSelected>>', update_group_info)
        
        # Buttons at bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        def save_layout():
            group = selected_group.get()
            if not group:
                messagebox.showwarning("No Group", "Please select a group first.")
                return
            
            if not canvas_regions:
                messagebox.showwarning("No Regions", "Please draw at least one region on the canvas.")
                return
            
            # Build layout configuration with region data
            config = {
                'type': 'visual',
                'group': group,
                'canvas_width': canvas_width,
                'canvas_height': canvas_height,
                'regions': []
            }
            
            for region in canvas_regions:
                # Convert pixel coordinates to percentages
                config['regions'].append({
                    'identifier': region['identifier'],
                    'x1': region['x1'] / canvas_width,
                    'y1': region['y1'] / canvas_height,
                    'x2': region['x2'] / canvas_width,
                    'y2': region['y2'] / canvas_height
                })
            
            # Save to profiles
            self.layout_profiles[group] = config
            
            messagebox.showinfo("Success", f"Visual layout profile saved for group {group}!")
            designer_window.destroy()
        
        def clear_layout():
            group = selected_group.get()
            if group and group in self.layout_profiles:
                del self.layout_profiles[group]
                messagebox.showinfo("Success", f"Layout profile cleared for group {group}.")
            else:
                messagebox.showinfo("Info", "No layout profile to clear for this group.")
        
        ttk.Button(button_frame, text="Save Layout", command=save_layout, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Layout", command=clear_layout, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=designer_window.destroy, width=15).pack(side=tk.LEFT, padx=5)
    
    def generate_powerpoint(self):
        """Generate PowerPoint presentation from grouped images"""
        from ppt_generator import PowerPointGenerator
        from image_uploader import ImageIdentifier
        
        # Check if there are any grouped images
        images = self.uploader.list_images()
        grouped_images = [img for img in images if img.get('metadata', {}).get('numerical_prefix')]
        
        if not grouped_images:
            messagebox.showwarning("No Groups", "No grouped images found. Please upload and group images first.")
            return
        
        # Create selection dialog
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Generate PowerPoint")
        selection_window.geometry("500x600")
        selection_window.transient(self.root)
        
        # Main frame
        main_frame = ttk.Frame(selection_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Select Groups to Include", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Choose which groups to add to the presentation:").pack(pady=(0, 10))
        
        # Scrollable frame for checkboxes
        canvas = tk.Canvas(main_frame, height=300)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Get all groups with formatted labels
        groups_dict = {}
        for img in grouped_images:
            metadata = img.get('metadata', {})
            numerical_prefix = metadata.get('numerical_prefix')
            identifier = metadata.get('identifier')
            
            if numerical_prefix:
                group_label = ImageIdentifier.format_group_label(numerical_prefix, identifier)
                if group_label not in groups_dict:
                    groups_dict[group_label] = []
                groups_dict[group_label].append(img)
        
        sorted_groups = sorted(groups_dict.keys())
        
        # Create checkboxes for each group
        group_vars = {}
        for group_label in sorted_groups:
            var = tk.BooleanVar(value=True)
            group_vars[group_label] = var
            
            count = len(groups_dict[group_label])
            cb = ttk.Checkbutton(
                scrollable_frame,
                text=f"{group_label} ({count} image{'s' if count != 1 else ''})",
                variable=var
            )
            cb.pack(anchor='w', pady=2)
        
        canvas.pack(side="left", fill="both", expand=True, pady=10)
        scrollbar.pack(side="right", fill="y", pady=10)
        
        # Select/Deselect all buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        def select_all():
            for var in group_vars.values():
                var.set(True)
        
        def deselect_all():
            for var in group_vars.values():
                var.set(False)
        
        ttk.Button(button_frame, text="Select All", command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Deselect All", command=deselect_all).pack(side=tk.LEFT, padx=5)
        
        # Output file selection
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(output_frame, text="Output file:").pack(anchor='w', pady=(0, 5))
        
        file_select_frame = ttk.Frame(output_frame)
        file_select_frame.pack(fill=tk.X)
        
        output_path_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "presentation.pptx"))
        output_entry = ttk.Entry(file_select_frame, textvariable=output_path_var, width=50)
        output_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def browse_output():
            filepath = filedialog.asksaveasfilename(
                title="Save PowerPoint As",
                defaultextension=".pptx",
                initialdir=os.path.expanduser("~/Desktop"),
                initialfile="presentation.pptx",
                filetypes=[("PowerPoint Files", "*.pptx"), ("All Files", "*.*")]
            )
            if filepath:
                output_path_var.set(filepath)
        
        ttk.Button(file_select_frame, text="Browse...", command=browse_output).pack(side=tk.LEFT)
        
        # Preview and Generate buttons
        def preview_slides():
            # Get selected groups
            selected_groups = [label for label, var in group_vars.items() if var.get()]
            
            if not selected_groups:
                messagebox.showwarning("No Selection", "Please select at least one group.")
                return
            
            # Create preview window
            preview_window = tk.Toplevel(selection_window)
            preview_window.title("Slide Preview")
            preview_window.geometry("1100x700")
            
            # Main frame
            preview_main_frame = ttk.Frame(preview_window, padding="10")
            preview_main_frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(preview_main_frame, text="Slide Layout Preview", 
                     font=('Helvetica', 14, 'bold')).pack(pady=(0, 10))
            
            # Info text
            info_text = f"Total slides: {len(selected_groups)} | Click on a slide to see details"
            ttk.Label(preview_main_frame, text=info_text, justify=tk.LEFT).pack(pady=(0, 10))
            
            # Scrollable frame for slide thumbnails
            canvas = tk.Canvas(preview_main_frame, bg='white')
            scrollbar = ttk.Scrollbar(preview_main_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # Display visual preview for each selected group
            from ppt_generator import PowerPointGenerator
            temp_generator = PowerPointGenerator(self.uploader.index)
            temp_generator.layout_profiles = self.layout_profiles
            
            # Store preview images to prevent garbage collection
            preview_window.preview_images = {}
            
            def show_slide_details(group_label, images):
                """Show detailed information about a slide"""
                detail_text = f"Group: {group_label}\n"
                detail_text += f"Total Images: {len(images)}\n\n"
                detail_text += "Files:\n"
                for img in images:
                    detail_text += f"  • {img['filename']}\n"
                messagebox.showinfo(f"Slide Details - {group_label}", detail_text)
            
            for idx, group_label in enumerate(sorted(selected_groups), 1):
                group_images = groups_dict[group_label]
                num_images = len(group_images)
                
                # Determine layout type
                layout_type = "Grid"
                if group_label in self.layout_profiles:
                    profile = self.layout_profiles[group_label]
                    if profile.get('type') == 'visual':
                        layout_type = "Visual Custom"
                    else:
                        layout_type = profile.get('type', 'auto').title()
                else:
                    # Auto-detect layout type
                    identifiers = [img.get('metadata', {}).get('identifier', '') for img in group_images]
                    has_spectrum = any(id == 'Spectrum' or id in ImageIdentifier.SPECTRUM_IDENTIFIERS for id in identifiers)
                    has_other = any(id != 'Spectrum' and id not in ImageIdentifier.SPECTRUM_IDENTIFIERS for id in identifiers)
                    
                    if has_spectrum and has_other:
                        layout_type = "Mixed"
                    elif has_spectrum:
                        layout_type = "Horizontal"
                    else:
                        rows, cols = temp_generator.calculate_grid_layout(num_images)
                        layout_type = f"Grid {rows}×{cols}"
                
                # Create slide preview frame
                slide_frame = ttk.LabelFrame(scrollable_frame, 
                                            text=f"Slide {idx}: {group_label} ({layout_type})", 
                                            padding="10")
                slide_frame.pack(fill=tk.X, pady=5, padx=5)
                
                # Create canvas for visual preview (16:9 aspect ratio matching PowerPoint)
                preview_width = 640
                preview_height = 480  # 10:7.5 ratio like PowerPoint
                slide_canvas = tk.Canvas(slide_frame, width=preview_width, height=preview_height, 
                                       bg='white', highlightthickness=1, highlightbackground='gray',
                                       cursor='hand2')
                slide_canvas.pack(pady=5)
                
                # Bind click event to show details
                slide_canvas.bind('<Button-1>', lambda e, gl=group_label, gi=group_images: show_slide_details(gl, gi))
                
                # Draw slide preview matching PowerPoint formatting
                try:
                    # PowerPoint dimensions scaled to canvas
                    # Slide: 10" x 7.5", Margins: L/R=0.5", T=0.75", B=0.5"
                    scale = preview_width / 10.0  # pixels per inch
                    
                    margin_left = int(0.5 * scale)
                    margin_right = int(0.5 * scale)
                    margin_top = int(0.75 * scale)
                    margin_bottom = int(0.5 * scale)
                    title_height = int(0.4 * scale)
                    title_top = int(0.25 * scale)
                    label_height = int(0.4 * scale)
                    label_spacing = int(0.1 * scale)
                    
                    available_width = preview_width - margin_left - margin_right
                    available_height = preview_height - margin_top - margin_bottom
                    
                    # Draw border
                    slide_canvas.create_rectangle(0, 0, preview_width, preview_height, 
                                                 outline='black', width=2)
                    
                    # Draw title (matching PowerPoint title position)
                    slide_canvas.create_text(margin_left + available_width // 2, title_top + title_height // 2, 
                                           text=f"Group {group_label}", 
                                           font=('Arial', 18, 'bold'))
                    
                    # Draw image placeholders based on layout
                    if group_label in self.layout_profiles and self.layout_profiles[group_label].get('type') == 'visual':
                        # Visual custom layout - show actual images in regions
                        regions = self.layout_profiles[group_label].get('regions', [])
                        
                        # Group images by their identifier
                        images_by_identifier = {}
                        for img in group_images:
                            identifier = img.get('metadata', {}).get('identifier', 'Unknown')
                            if identifier not in images_by_identifier:
                                images_by_identifier[identifier] = []
                            images_by_identifier[identifier].append(img)
                        
                        # Draw each region with its images
                        for region_idx, region in enumerate(regions):
                            # Scale region coordinates to match PowerPoint margins
                            x1 = margin_left + region['x1'] * available_width
                            y1 = margin_top + region['y1'] * available_height
                            x2 = margin_left + region['x2'] * available_width
                            y2 = margin_top + region['y2'] * available_height
                            
                            region_identifier = region['identifier']
                            region_images = images_by_identifier.get(region_identifier, [])
                            
                            # Draw region border
                            slide_canvas.create_rectangle(x1, y1, x2, y2, 
                                                        outline='#666', width=1)
                            
                            # If there are images for this identifier, display them
                            if region_images:
                                # Calculate grid for images in this region
                                num_region_images = len(region_images)
                                region_rows, region_cols = temp_generator.calculate_grid_layout(num_region_images)
                                
                                region_width = x2 - x1
                                region_height = y2 - y1
                                
                                cell_width = region_width / region_cols
                                cell_height = region_height / region_rows
                                
                                for img_idx, img_info in enumerate(region_images):
                                    row = img_idx // region_cols
                                    col = img_idx % region_cols
                                    
                                    cell_x = x1 + col * cell_width
                                    cell_y = y1 + row * cell_height
                                    
                                    # Get image aspect ratio
                                    img_aspect_ratio = img_info['width'] / img_info['height'] if img_info['height'] > 0 else 1.0
                                    
                                    # Calculate image dimensions with label space
                                    available_cell_height = cell_height - label_height - label_spacing
                                    
                                    # Fit image maintaining aspect ratio
                                    if cell_width / available_cell_height > img_aspect_ratio:
                                        img_height = available_cell_height
                                        img_width = img_height * img_aspect_ratio
                                    else:
                                        img_width = cell_width * 0.95
                                        img_height = img_width / img_aspect_ratio
                                        if img_height > available_cell_height:
                                            img_height = available_cell_height
                                            img_width = img_height * img_aspect_ratio
                                    
                                    # Center image in cell
                                    img_x = cell_x + (cell_width - img_width) / 2
                                    img_y = cell_y + (available_cell_height - img_height) / 2
                                    
                                    # Load and display image thumbnail
                                    try:
                                        filepath = img_info['filepath']
                                        if os.path.exists(filepath):
                                            img = Image.open(filepath)
                                            img.thumbnail((int(img_width), int(img_height)), Image.Resampling.LANCZOS)
                                            photo = ImageTk.PhotoImage(img)
                                            preview_window.preview_images[f"{idx}_{region_idx}_{img_idx}"] = photo
                                            
                                            slide_canvas.create_image(img_x + img_width / 2, img_y + img_height / 2, image=photo)
                                            
                                            # Add filename label
                                            label_y = cell_y + available_cell_height + label_spacing
                                            filename = img_info['filename']
                                            if len(filename) > 20:
                                                filename = filename[:17] + "..."
                                            slide_canvas.create_text(cell_x + cell_width / 2, label_y + label_height / 2,
                                                                   text=filename, font=('Arial', 6), width=cell_width)
                                        else:
                                            slide_canvas.create_rectangle(img_x, img_y, img_x + img_width, img_y + img_height,
                                                                        fill='#E8E8E8', outline='#999')
                                            slide_canvas.create_text(img_x + img_width / 2, img_y + img_height / 2,
                                                                   text="Not Found", font=('Arial', 8))
                                    except Exception as img_err:
                                        slide_canvas.create_rectangle(img_x, img_y, img_x + img_width, img_y + img_height,
                                                                    fill='#E8E8E8', outline='#999')
                                        slide_canvas.create_text(img_x + img_width / 2, img_y + img_height / 2,
                                                               text="Error", font=('Arial', 8))
                            else:
                                # No images for this identifier - show label
                                slide_canvas.create_text((x1+x2)//2, (y1+y2)//2, 
                                                       text=f"{region_identifier}\n(No images)", 
                                                       font=('Arial', 10), fill='gray')
                    else:
                        # Standard grid layout matching PowerPoint
                        rows, cols = temp_generator.calculate_grid_layout(num_images)
                        
                        cell_width = available_width / cols
                        cell_height = available_height / rows
                        
                        for i in range(min(num_images, rows * cols)):
                            if i >= len(group_images):
                                break
                                
                            row = i // cols
                            col = i % cols
                            
                            # Calculate cell position
                            cell_x = margin_left + col * cell_width
                            cell_y = margin_top + row * cell_height
                            
                            # Get image info
                            img_info = group_images[i]
                            img_aspect_ratio = img_info['width'] / img_info['height'] if img_info['height'] > 0 else 1.0
                            
                            # Calculate image dimensions (leaving space for label)
                            available_cell_height = cell_height - label_height - label_spacing
                            
                            # Fit image in cell while maintaining aspect ratio
                            if cell_width / available_cell_height > img_aspect_ratio:
                                # Height constrained
                                img_height = available_cell_height
                                img_width = img_height * img_aspect_ratio
                            else:
                                # Width constrained
                                img_width = cell_width * 0.95  # 95% to add small padding
                                img_height = img_width / img_aspect_ratio
                                if img_height > available_cell_height:
                                    img_height = available_cell_height
                                    img_width = img_height * img_aspect_ratio
                            
                            # Center image in cell (above label area)
                            img_x = cell_x + (cell_width - img_width) / 2
                            img_y = cell_y + (available_cell_height - img_height) / 2
                            
                            # Load and display image thumbnail
                            try:
                                filepath = img_info['filepath']
                                if os.path.exists(filepath):
                                    img = Image.open(filepath)
                                    img.thumbnail((int(img_width), int(img_height)), Image.Resampling.LANCZOS)
                                    photo = ImageTk.PhotoImage(img)
                                    preview_window.preview_images[f"{idx}_{i}"] = photo
                                    
                                    slide_canvas.create_image(img_x + img_width / 2, img_y + img_height / 2, image=photo)
                                    
                                    # Add filename label below image (matching PowerPoint)
                                    label_y = cell_y + available_cell_height + label_spacing
                                    # Truncate filename if too long
                                    filename = img_info['filename']
                                    if len(filename) > 20:
                                        filename = filename[:17] + "..."
                                    slide_canvas.create_text(cell_x + cell_width / 2, label_y + label_height / 2,
                                                           text=filename, font=('Arial', 6), width=cell_width)
                                else:
                                    slide_canvas.create_rectangle(img_x, img_y, img_x + img_width, img_y + img_height,
                                                                fill='#E8E8E8', outline='#999')
                                    slide_canvas.create_text(img_x + img_width / 2, img_y + img_height / 2,
                                                           text="Not Found", font=('Arial', 8))
                            except Exception as img_err:
                                # Draw placeholder rectangle
                                slide_canvas.create_rectangle(img_x, img_y, img_x + img_width, img_y + img_height,
                                                            fill='#E8E8E8', outline='#999')
                                slide_canvas.create_text(img_x + img_width / 2, img_y + img_height / 2,
                                                       text="Error", font=('Arial', 8))
                
                except Exception as e:
                    slide_canvas.create_text(preview_width//2, preview_height//2, 
                                           text=f"Preview error: {str(e)}", 
                                           font=('Arial', 10))
                
                # Info label below canvas
                info = f"{num_images} image{'s' if num_images != 1 else ''} | "
                info += f"Files: {', '.join([img['filename'] for img in group_images[:2]])}"
                if num_images > 2:
                    info += f" ... +{num_images - 2} more"
                
                ttk.Label(slide_frame, text=info, foreground='gray').pack(anchor='w', pady=(5, 0))
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Close button
            ttk.Button(preview_main_frame, text="Close", 
                      command=preview_window.destroy, width=15).pack(pady=10)
        
        def generate():
            # Get selected groups
            selected_groups = [label for label, var in group_vars.items() if var.get()]
            
            if not selected_groups:
                messagebox.showwarning("No Selection", "Please select at least one group.")
                return
            
            output_path = output_path_var.get().strip()
            if not output_path:
                messagebox.showwarning("No Output File", "Please specify an output file.")
                return
            
            # Validate output directory exists
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                messagebox.showerror("Invalid Path", f"Directory does not exist:\n{output_dir}")
                return
            
            # Close selection window
            selection_window.destroy()
            
            # Generate presentation
            try:
                generator = PowerPointGenerator(self.uploader.index)
                # Pass layout profiles to generator
                generator.layout_profiles = self.layout_profiles
                success = generator.generate_presentation(output_path, selected_groups)
                
                if success:
                    result = messagebox.askquestion("Success", 
                                      f"PowerPoint generated successfully!\n\n"
                                      f"File: {output_path}\n"
                                      f"Slides: {len(selected_groups)}\n\n"
                                      f"Would you like to open the file location?")
                    
                    if result == 'yes':
                        # Open file location in Finder/Explorer
                        import subprocess
                        if os.name == 'posix':  # macOS/Linux
                            subprocess.run(['open', '-R', output_path])
                        elif os.name == 'nt':  # Windows
                            subprocess.run(['explorer', '/select,', output_path])
                else:
                    messagebox.showerror("Error", "Failed to generate PowerPoint. Check console for details.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate PowerPoint:\n{str(e)}")
        
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(pady=10)
        
        ttk.Button(action_frame, text="Preview", command=preview_slides, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Generate", command=generate, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Cancel", command=selection_window.destroy, width=15).pack(side=tk.LEFT, padx=5)


def main():
    """Launch the GUI application"""
    root = tk.Tk()
    app = ImageUploaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
