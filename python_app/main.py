#!/usr/bin/env python3
"""
Microsoft Graph API URL Filter UI - Enhanced with Dynamic Filter Builders
A Python GUI application with customizable dynamic filters for Microsoft Graph endpoints
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import webbrowser
import pyperclip
from typing import List, Dict, Any, Optional
import json
import configparser
import os

# MSAL integration (optional - will show placeholder if not available)
try:
    import msal
    import requests
    from bs4 import BeautifulSoup
    MSAL_AVAILABLE = True
except ImportError:
    MSAL_AVAILABLE = False

class GraphEndpoint:
    def __init__(self, name: str, url: str, method: str, category: str, 
                 scopes: List[str], description: str, version: str, filters: List[Any] = None): # type: ignore
        self.name = name
        self.url = url
        self.method = method
        self.category = category
        self.scopes = scopes
        self.description = description
        self.version = version
        self.filters = filters or []

class MicrosoftEntraApp:
    """MSAL integration for Microsoft Graph authentication"""
    
    def __init__(self):
        if not MSAL_AVAILABLE:
            self.app = None
            return
            
        # Default config - users can modify this
        self.CLIENT_ID = "your-client-id"  # Users need to set this
        self.TENANT_ID = "your-tenant-id"  # Users need to set this
        AUTHORITY = f"https://login.microsoftonline.com/{self.TENANT_ID}"
        self.SCOPES = ["Mail.Read", "Mail.ReadWrite"]
        
        try:
            # Try to load config if it exists
            if os.path.exists("config.cfg"):
                config = configparser.ConfigParser()
                config.read("config.cfg")
                if "azure" in config:
                    self.CLIENT_ID = config["azure"].get("clientId", self.CLIENT_ID)
                    self.TENANT_ID = config["azure"].get("tenantId", self.TENANT_ID)
                    AUTHORITY = f"https://login.microsoftonline.com/{self.TENANT_ID}"
            
            self.app = msal.PublicClientApplication(self.CLIENT_ID, authority=AUTHORITY)
        except Exception as e:
            print(f"Error initializing MSAL: {e}")
            self.app = None

    def acquire_token(self) -> Optional[str]:
        """Acquire access token for Microsoft Graph"""
        if not self.app:
            return None
            
        try:
            accounts = self.app.get_accounts()
            result = self.app.acquire_token_silent(self.SCOPES, account=accounts[0]) if accounts else None
            if not result:
                result = self.app.acquire_token_interactive(scopes=self.SCOPES)
            return result.get("access_token") if result and "access_token" in result else None
        except Exception as e:
            print(f"Error acquiring token: {e}")
            return None
    
    def get_mail_folders(self, headers: Dict[str, str]) -> List[Dict]:
        """Get user's mail folders"""
        try:
            url = "https://graph.microsoft.com/v1.0/me/mailFolders"
            response = requests.get(url, headers=headers)
            if response.ok:
                return response.json().get("value", [])
        except Exception as e:
            print(f"Error getting mail folders: {e}")
        return []

class MicrosoftGraphExplorer:
    def __init__(self, root):
        self.root = root
        self.root.title("Microsoft Graph API Explorer - Dynamic Filters")
        self.root.geometry("1700x1000")
        self.root.configure(bg="#000000")
        
        # Configure ttk styles for dark theme
        self.setup_dark_theme()
        
        # Initialize data
        self.endpoints = self.load_endpoints()
        self.filtered_endpoints = self.endpoints.copy()
        
        # Filter variables
        self.search_var = tk.StringVar()
        self.category_var = tk.StringVar(value="All")
        self.version_var = tk.StringVar(value="All")
        self.method_var = tk.StringVar(value="All")
        self.selected_scopes = set()
        self.selected_endpoint = None
        self.selected_filters = []
        
        # MSAL integration and folder selection
        self.msal_app = MicrosoftEntraApp() if MSAL_AVAILABLE else None
        self.access_token = None
        self.user_folders = []
        self.selected_folder = None
        
        # Bind search
        self.search_var.trace_add("write", self.filter_endpoints)
        
        self.setup_ui()
        self.filter_endpoints()
    
    def setup_dark_theme(self):
        """Configure dark theme for ttk widgets"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('TCombobox',
                       fieldbackground='#2a2a2a',
                       background='#2a2a2a',
                       foreground='white',
                       arrowcolor='white',
                       bordercolor='#555555',
                       lightcolor='#2a2a2a',
                       darkcolor='#2a2a2a')
        
        style.map('TCombobox', 
                 fieldbackground=[('readonly', '#2a2a2a')],
                 selectbackground=[('readonly', '#0078d4')])
    
    def load_endpoints(self) -> List[GraphEndpoint]:
        """Load Microsoft Graph endpoints data with dynamic filter configurations"""
        endpoints_data = [
            {
                "name": "List Messages",
                "url": "https://graph.microsoft.com/v1.0/me/messages",
                "method": "GET",
                "category": "Mail",
                "scopes": ["Mail.Read", "Mail.ReadWrite"],
                "description": "Get the messages in the signed-in user's mailbox",
                "version": "v1.0",
                "filters": [
                    {
                        "name": "Filter by Read Status",
                        "template": "?$filter=isRead eq {value}",
                        "type": "boolean",
                        "options": ["true", "false"],
                        "default": "false",
                        "description": "Show read or unread messages"
                    },
                    {
                        "name": "Filter by Date Range",
                        "template": "?$filter=receivedDateTime ge {datetime}",
                        "type": "datetime",
                        "default": "2025-01-15T00:00:00Z",
                        "description": "Messages received after this date"
                    },
                    {
                        "name": "Limit Results",
                        "template": "?$top={number}",
                        "type": "number",
                        "default": "25",
                        "min": 1,
                        "max": 1000,
                        "description": "Maximum number of messages to return"
                    },
                    {
                        "name": "Filter by Attachments",
                        "template": "?$filter=hasAttachments eq {value}",
                        "type": "boolean",
                        "options": ["true", "false"],
                        "default": "true",
                        "description": "Show messages with or without attachments"
                    },
                    {
                        "name": "Filter by Importance",
                        "template": "?$filter=importance eq '{value}'",
                        "type": "select",
                        "options": ["low", "normal", "high"],
                        "default": "high",
                        "description": "Filter by message importance level"
                    },
                    {
                        "name": "Order Results",
                        "template": "?$orderBy={field} {direction}",
                        "type": "compound",
                        "fields": {
                            "field": {
                                "type": "select",
                                "options": ["receivedDateTime", "subject", "from", "importance"],
                                "default": "receivedDateTime"
                            },
                            "direction": {
                                "type": "select", 
                                "options": ["asc", "desc"],
                                "default": "desc"
                            }
                        },
                        "description": "Sort messages by field and direction"
                    },
                    {
                        "name": "Select Fields",
                        "template": "?$select={fields}",
                        "type": "multiselect",
                        "options": ["subject", "from", "receivedDateTime", "isRead", "hasAttachments", "importance", "body"],
                        "default": ["subject", "from", "receivedDateTime", "isRead"],
                        "description": "Choose which fields to include in response"
                    },
                    {
                        "name": "Search Content",
                        "template": "?$search=\"{text}\"",
                        "type": "text",
                        "default": "project update",
                        "description": "Search message content for specific text"
                    },
                    {
                        "name": "Filter by Sender",
                        "template": "?$filter=from/emailAddress/address eq '{email}'",
                        "type": "email",
                        "default": "example@company.com",
                        "description": "Show messages from specific email address"
                    }
                ]
            },
            {
                "name": "List Mail Folders",
                "url": "https://graph.microsoft.com/v1.0/me/mailFolders",
                "method": "GET",
                "category": "Mail",
                "scopes": ["Mail.Read", "Mail.ReadWrite"],
                "description": "Get the mail folders in the signed-in user's mailbox",
                "version": "v1.0",
                "filters": [
                    {
                        "name": "Filter by Folder Name",
                        "template": "?$filter=displayName eq '{value}'",
                        "type": "text",
                        "default": "Inbox",
                        "description": "Filter folders by exact name match"
                    },
                    {
                        "name": "Select Folder Fields",
                        "template": "?$select={fields}",
                        "type": "multiselect",
                        "options": ["displayName", "unreadItemCount", "totalItemCount", "id", "parentFolderId"],
                        "default": ["displayName", "unreadItemCount", "totalItemCount"],
                        "description": "Choose which folder properties to return"
                    },
                    {
                        "name": "Order Folders",
                        "template": "?$orderBy={field} {direction}",
                        "type": "compound",
                        "fields": {
                            "field": {
                                "type": "select",
                                "options": ["displayName", "unreadItemCount", "totalItemCount"],
                                "default": "displayName"
                            },
                            "direction": {
                                "type": "select",
                                "options": ["asc", "desc"],
                                "default": "asc"
                            }
                        },
                        "description": "Sort folders by field and direction"
                    },
                    {
                        "name": "Expand Child Folders",
                        "template": "?$expand=childFolders",
                        "type": "static",
                        "description": "Include child folders in the response"
                    },
                    {
                        "name": "Limit Results",
                        "template": "?$top={number}",
                        "type": "number",
                        "default": "50",
                        "min": 1,
                        "max": 500,
                        "description": "Maximum number of folders to return"
                    }
                ]
            }
        ]
        
        return [GraphEndpoint(**data) for data in endpoints_data]
    
    def setup_ui(self):
        """Setup the user interface with dark theme"""
        # Title frame (top 15%)
        title_frame = tk.Frame(self.root, bg="#1a1a1a", height=120)
        title_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame, 
            text="🔷 Microsoft Graph API Explorer - Dynamic Filters", 
            font=("Segoe UI", 16, "bold"),
            fg="white",
            bg="#1a1a1a"
        )
        title_label.pack(pady=(15, 5))
        
        subtitle_label = tk.Label(
            title_frame,
            text="Build custom Microsoft Graph queries with dynamic, interactive filter builders",
            font=("Segoe UI", 10),
            fg="#cccccc",
            bg="#1a1a1a"
        )
        subtitle_label.pack()
        
        # Authentication and folder selection frame
        auth_folder_frame = tk.Frame(title_frame, bg="#1a1a1a")
        auth_folder_frame.pack(pady=8)
        
        # MSAL status and auth button
        auth_frame = tk.Frame(auth_folder_frame, bg="#1a1a1a")
        auth_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        self.auth_status_label = tk.Label(
            auth_frame,
            text="MSAL: Not authenticated" if MSAL_AVAILABLE else "MSAL: Not available",
            font=("Segoe UI", 9),
            fg="#ff6b6b" if not self.access_token else "#51cf66",
            bg="#1a1a1a"
        )
        self.auth_status_label.pack(side=tk.LEFT, padx=(0, 10))
        
        if MSAL_AVAILABLE:
            auth_btn = tk.Button(
                auth_frame,
                text="🔐 Authenticate",
                font=("Segoe UI", 9),
                bg="#2a2a2a",
                fg="white",
                padx=15,
                relief=tk.RAISED,
                bd=1,
                highlightthickness=0,
                activebackground="#3a3a3a",
                activeforeground="white",
                command=self.authenticate_msal
            )
            auth_btn.pack(side=tk.LEFT)
        
        # Folder selection frame (initially hidden)
        self.folder_frame = tk.Frame(auth_folder_frame, bg="#1a1a1a")
        
        folder_label = tk.Label(
            self.folder_frame,
            text="📁 Selected Folder:",
            font=("Segoe UI", 9, "bold"),
            fg="white",
            bg="#1a1a1a"
        )
        folder_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.selected_folder_label = tk.Label(
            self.folder_frame,
            text="None",
            font=("Segoe UI", 9),
            fg="#ffd43b",
            bg="#1a1a1a"
        )
        self.selected_folder_label.pack(side=tk.LEFT, padx=(0, 10))
        
        browse_folders_btn = tk.Button(
            self.folder_frame,
            text="🗂️ Browse Folders",
            font=("Segoe UI", 9),
            bg="#2a2a2a",
            fg="white",
            padx=15,
            relief=tk.RAISED,
            bd=1,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=self.browse_folders
        )
        browse_folders_btn.pack(side=tk.LEFT)
        
        # Filters frame (15%)
        self.setup_filters(self.root)
        
        # Main content frame (70%) - split equally
        main_frame = tk.Frame(self.root, bg="#000000")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left half: API Endpoints
        self.setup_results(main_frame)
        
        # Right half: Dynamic Filter Builders
        self.setup_filtering_panel(main_frame)
    
    def setup_filters(self, parent):
        """Setup basic filter controls"""
        filters_frame = tk.LabelFrame(parent, text="Basic Filters", font=("Segoe UI", 10, "bold"), 
                                     bg="#000000", fg="white", pady=5)
        filters_frame.pack(fill=tk.X, pady=5, padx=10)
        
        # Row 1: Search and basic filters
        row1 = tk.Frame(filters_frame, bg="#000000")
        row1.pack(fill=tk.X, padx=10, pady=5)
        
        # Search
        tk.Label(row1, text="Search:", font=("Segoe UI", 9), bg="#000000", fg="white").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        search_entry = tk.Entry(row1, textvariable=self.search_var, font=("Segoe UI", 9), width=25, 
                               bg="#2a2a2a", fg="white", insertbackground="white")
        search_entry.grid(row=0, column=1, padx=(0, 15), sticky=tk.W)
        
        # Category filter
        tk.Label(row1, text="Category:", font=("Segoe UI", 9), bg="#000000", fg="white").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        categories = ["All", "Users", "Groups", "Mail", "Calendar", "Files", "Applications", "Teams"]
        category_combo = ttk.Combobox(row1, textvariable=self.category_var, values=categories, 
                                     state="readonly", width=12)
        category_combo.grid(row=0, column=3, padx=(0, 15))
        category_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_endpoints())
        
        # Version filter
        tk.Label(row1, text="Version:", font=("Segoe UI", 9), bg="#000000", fg="white").grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        versions = ["All", "v1.0", "beta"]
        version_combo = ttk.Combobox(row1, textvariable=self.version_var, values=versions, 
                                    state="readonly", width=8)
        version_combo.grid(row=0, column=5)
        version_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_endpoints())
    
    def setup_results(self, parent):
        """Setup results display - left half"""
        # Left panel for endpoints (50% width)
        left_frame = tk.Frame(parent, bg="#000000")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        results_frame = tk.LabelFrame(left_frame, text="API Endpoints", font=("Segoe UI", 12, "bold"), 
                                     bg="#000000", fg="white")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Results count
        self.count_label = tk.Label(results_frame, text="", font=("Segoe UI", 10), bg="#000000", fg="white")
        self.count_label.pack(anchor=tk.W, padx=10, pady=5)
        
        # Scrollable results area
        canvas = tk.Canvas(results_frame, bg="#1a1a1a", highlightthickness=0)
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg="#1a1a1a")
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)
        
        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
    
    def setup_filtering_panel(self, parent):
        """Setup dynamic filtering panel - right half"""
        # Right panel for filtering options (50% width)
        right_frame = tk.Frame(parent, bg="#000000")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        self.filtering_frame = tk.LabelFrame(right_frame, text="Dynamic Filter Builders", 
                                           font=("Segoe UI", 12, "bold"), 
                                           bg="#000000", fg="white")
        self.filtering_frame.pack(fill=tk.BOTH, expand=True)
        
        # Default message when no endpoint is selected
        self.no_selection_label = tk.Label(
            self.filtering_frame,
            text="Select an endpoint from the left panel\nto build custom filters",
            font=("Segoe UI", 11),
            fg="#cccccc",
            bg="#000000",
            justify=tk.CENTER
        )
        self.no_selection_label.pack(expand=True)
        
        # Scrollable filtering content (initially hidden)
        self.filter_canvas = tk.Canvas(self.filtering_frame, bg="#1a1a1a", highlightthickness=0)
        self.filter_scrollbar = ttk.Scrollbar(self.filtering_frame, orient=tk.VERTICAL, command=self.filter_canvas.yview)
        self.filter_content_frame = tk.Frame(self.filter_canvas, bg="#1a1a1a")
        
        self.filter_content_frame.bind(
            "<Configure>",
            lambda e: self.filter_canvas.configure(scrollregion=self.filter_canvas.bbox("all"))
        )
        
        self.filter_canvas.create_window((0, 0), window=self.filter_content_frame, anchor="nw")
        self.filter_canvas.configure(yscrollcommand=self.filter_scrollbar.set)
        
        # Mouse wheel scrolling for filter panel
        def _on_filter_mousewheel(event):
            self.filter_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.filter_canvas.bind("<MouseWheel>", _on_filter_mousewheel)
    
    def authenticate_msal(self):
        """Authenticate with MSAL and get access token"""
        if not self.msal_app:
            messagebox.showerror("Error", "MSAL not available. Please install required packages.")
            return
        
        try:
            self.access_token = self.msal_app.acquire_token()
            if self.access_token:
                self.auth_status_label.config(text="MSAL: Authenticated ✓", fg="#51cf66")
                
                # Load user folders and show folder selection
                self.load_user_folders()
                self.folder_frame.pack(side=tk.LEFT)
                
                messagebox.showinfo("Success", "Successfully authenticated with Microsoft Graph!")
            else:
                self.auth_status_label.config(text="MSAL: Authentication failed", fg="#ff6b6b")
                messagebox.showerror("Error", "Failed to authenticate. Please check your credentials.")
        except Exception as e:
            messagebox.showerror("Error", f"Authentication error: {str(e)}")
    
    def load_user_folders(self):
        """Load user's mail folders after authentication, including all nested folders"""
        if not self.access_token:
            return
        
        try:
            headers = {"Authorization": f"Bearer {self.access_token}"}
            # Start with top-level folders
            top_level_folders = self.msal_app.get_mail_folders(headers) # type: ignore
            self.user_folders = []
            
            # Recursively load all folders and their children
            for folder in top_level_folders:
                # Add depth and display info to top-level folders
                folder['depth'] = 0
                folder['indentDisplay'] = f"📁 {folder.get('displayName', 'Unknown')}"
                folder['hierarchyName'] = folder.get('displayName', 'Unknown')
                self.user_folders.append(folder)
                self._load_child_folders_recursive(folder, headers, depth=0)
                        
        except Exception as e:
            print(f"Error loading folders: {e}")
            self.user_folders = []
    
    def _load_child_folders_recursive(self, parent_folder: Dict, headers: Dict[str, str], depth: int = 0, max_depth: int = 10):
        """Recursively load child folders with depth limiting to prevent infinite loops"""
        if depth >= max_depth:
            return
        
        try:
            parent_id = parent_folder.get('id')
            parent_name = parent_folder.get('displayName', 'Unknown')
            
            child_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{parent_id}/childFolders"
            child_response = requests.get(child_url, headers=headers)
            
            if child_response.ok:
                child_folders = child_response.json().get("value", [])
                
                for child in child_folders:
                    # Create hierarchy display name with proper indentation
                    indent = "  " * (depth + 1)
                    original_name = child.get('displayName', 'Unknown')
                    
                    # Build full hierarchy path
                    parent_hierarchy = parent_folder.get('hierarchyName', parent_name)
                    child['hierarchyName'] = f"{parent_hierarchy}/{original_name}"
                    
                    child['parentName'] = parent_name
                    child['parentId'] = parent_id
                    child['depth'] = depth + 1
                    child['indentDisplay'] = f"{indent}📁 {original_name}"
                    
                    self.user_folders.append(child)
                    
                    # Recursively load children of this child folder
                    self._load_child_folders_recursive(child, headers, depth + 1, max_depth)
                    
        except Exception as e:
            print(f"Error loading child folders for {parent_folder.get('displayName', 'Unknown')}: {e}")
    
    def browse_folders(self):
        """Show folder browser popup"""
        if not self.user_folders:
            messagebox.showwarning("Warning", "No folders loaded. Please authenticate first.")
            return
        
        self.show_folder_browser()
    
    def show_folder_browser(self):
        """Show popup window for folder selection with nested folder support"""
        folder_popup = tk.Toplevel(self.root)
        folder_popup.title("Select Email Folder")
        folder_popup.geometry("600x500")
        folder_popup.configure(bg="#000000")
        
        # Header
        header_frame = tk.Frame(folder_popup, bg="#1a1a1a", height=60)
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        header_frame.pack_propagate(False)
        
        tk.Label(
            header_frame,
            text="📁 Select Email Folder for Filtering",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#1a1a1a"
        ).pack(pady=15)
        
        # Current selection display
        current_frame = tk.Frame(folder_popup, bg="#2a2a2a", relief=tk.RAISED, bd=1)
        current_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        
        tk.Label(
            current_frame,
            text="Currently Selected:",
            font=("Segoe UI", 10, "bold"),
            bg="#2a2a2a",
            fg="white"
        ).pack(anchor=tk.W, padx=10, pady=(5, 2))
        
        current_folder_text = self.selected_folder.get('hierarchyName', 'No folder selected') if self.selected_folder else "No folder selected"
        current_folder_label = tk.Label(
            current_frame,
            text=current_folder_text,
            font=("Segoe UI", 10),
            bg="#2a2a2a",
            fg="#ffd43b"
        )
        current_folder_label.pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        # Folders list
        folders_frame = tk.LabelFrame(folder_popup, text="Available Folders", font=("Segoe UI", 12, "bold"), 
                                     bg="#000000", fg="white")
        folders_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollable folders list
        folders_canvas = tk.Canvas(folders_frame, bg="#1a1a1a", highlightthickness=0)
        folders_scrollbar = ttk.Scrollbar(folders_frame, orient=tk.VERTICAL, command=folders_canvas.yview)
        folders_content_frame = tk.Frame(folders_canvas, bg="#1a1a1a")
        
        folders_content_frame.bind(
            "<Configure>",
            lambda e: folders_canvas.configure(scrollregion=folders_canvas.bbox("all"))
        )
        
        folders_canvas.create_window((0, 0), window=folders_content_frame, anchor="nw")
        folders_canvas.configure(yscrollcommand=folders_scrollbar.set)
        
        folders_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        folders_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
        
        # Mouse wheel scrolling
        def _on_folders_mousewheel(event):
            folders_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        folders_canvas.bind("<MouseWheel>", _on_folders_mousewheel)
        
        # Add folder cards
        for folder in self.user_folders:
            self.create_folder_card(folder, folders_content_frame, folder_popup, current_folder_label)
        
        # Close button
        close_frame = tk.Frame(folder_popup, bg="#000000")
        close_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        close_btn = tk.Button(
            close_frame,
            text="✕ Close",
            font=("Segoe UI", 10),
            bg="#2a2a2a",
            fg="white",
            padx=20,
            relief=tk.RAISED,
            bd=1,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=folder_popup.destroy
        )
        close_btn.pack(side=tk.RIGHT)
    
    def create_folder_card(self, folder: Dict, parent_frame: tk.Frame, popup_window: tk.Toplevel, current_label: tk.Label):
        """Create a selectable card for each folder"""
        depth = folder.get('depth', 0)
        depth_colors = {0: "#2a2a2a", 1: "#353535", 2: "#404040"}
        bg_color = depth_colors.get(depth, "#4a4a4a")
        
        card_frame = tk.Frame(parent_frame, bg=bg_color, relief=tk.RAISED, bd=1)
        card_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Folder info
        info_frame = tk.Frame(card_frame, bg=bg_color)
        info_frame.pack(fill=tk.X, padx=15, pady=8)
        
        # Use indented display name
        indent_display = folder.get('indentDisplay', f"📁 {folder.get('displayName', 'Unknown')}")
        
        name_label = tk.Label(
            info_frame,
            text=indent_display,
            font=("Segoe UI", 10 if depth == 0 else 9, "bold" if depth == 0 else "normal"),
            bg=bg_color,
            fg="white",
            anchor="w"
        )
        name_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Folder stats
        total_items = folder.get('totalItemCount', 0)
        unread_items = folder.get('unreadItemCount', 0)
        stats_text = f"{total_items}/{unread_items}"
        
        stats_label = tk.Label(
            info_frame,
            text=stats_text,
            font=("Segoe UI", 8),
            bg=bg_color,
            fg="#cccccc"
        )
        stats_label.pack(side=tk.LEFT, padx=(10, 10))
        
        # Select button
        def select_folder():
            self.selected_folder = folder
            display_name = folder.get('hierarchyName', folder.get('displayName', 'Unknown'))
            self.selected_folder_label.config(text=display_name)
            current_label.config(text=display_name)
            messagebox.showinfo("Folder Selected", f"Selected: {display_name}")
            popup_window.destroy()
        
        select_btn = tk.Button(
            info_frame,
            text="✓",
            font=("Segoe UI", 8),
            bg="#2a2a2a",
            fg="white",
            padx=8,
            relief=tk.RAISED,
            bd=1,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=select_folder
        )
        select_btn.pack(side=tk.RIGHT)
        
        # Make card clickable
        def on_card_click(event=None):
            select_folder()
        
        for widget in [card_frame, info_frame, name_label, stats_label]:
            widget.bind("<Button-1>", on_card_click)
    
    def show_filtering_options(self, endpoint: GraphEndpoint):
        """Display dynamic filtering options for the selected endpoint"""
        self.selected_endpoint = endpoint
        self.selected_filters = []
        
        # Hide default message
        self.no_selection_label.pack_forget()
        
        # Show filtering content
        self.filter_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.filter_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
        
        # Clear previous content
        for widget in self.filter_content_frame.winfo_children():
            widget.destroy()
        
        # Endpoint info
        info_frame = tk.Frame(self.filter_content_frame, bg="#1a1a1a")
        info_frame.pack(fill=tk.X, padx=10, pady=(10, 20))
        
        tk.Label(
            info_frame,
            text=f"Dynamic Filters for: {endpoint.name}",
            font=("Segoe UI", 12, "bold"),
            bg="#1a1a1a",
            fg="#66b3ff"
        ).pack(anchor=tk.W)
        
        # Folder selection status for mail endpoints
        if "messages" in endpoint.url.lower():
            folder_status_frame = tk.Frame(self.filter_content_frame, bg="#2a2a2a", relief=tk.RAISED, bd=1)
            folder_status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
            
            tk.Label(
                folder_status_frame,
                text="Target Folder:",
                font=("Segoe UI", 9, "bold"),
                bg="#2a2a2a",
                fg="white"
            ).pack(anchor=tk.W, padx=10, pady=(5, 2))
            
            folder_text = self.selected_folder.get('hierarchyName', 'All Messages') if self.selected_folder else "All Messages"
            tk.Label(
                folder_status_frame,
                text=folder_text,
                font=("Segoe UI", 9),
                bg="#2a2a2a",
                fg="#ffd43b"
            ).pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        # Selected filters display
        self.selected_filters_frame = tk.Frame(self.filter_content_frame, bg="#1a1a1a")
        self.selected_filters_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        self.selected_filters_label = tk.Label(
            self.selected_filters_frame,
            text="Selected Filters: None",
            font=("Segoe UI", 10, "bold"),
            bg="#1a1a1a",
            fg="#ffd43b"
        )
        self.selected_filters_label.pack(anchor=tk.W)
        
        # Execute button
        execute_frame = tk.Frame(self.filter_content_frame, bg="#1a1a1a")
        execute_frame.pack(fill=tk.X, padx=10, pady=(0, 20))
        
        execute_btn = tk.Button(
            execute_frame,
            text="🚀 Execute Query with Custom Filters",
            font=("Segoe UI", 10, "bold"),
            bg="#2a2a2a",
            fg="white",
            padx=20,
            pady=8,
            relief=tk.RAISED,
            bd=2,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=self.execute_selected_filters
        )
        execute_btn.pack(side=tk.LEFT, padx=(0, 10))

        clear_btn = tk.Button(
            execute_frame,
            text="🗑️ Clear Selected Filters",
            font=("Segoe UI", 10),
            bg="#2a2a2a",
            fg="white",
            padx=20,
            pady=8,
            relief=tk.RAISED,
            bd=2,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=self.clear_selected_filters
        )
        clear_btn.pack(side=tk.LEFT)
        
        # Dynamic filter builders
        if endpoint.filters:
            tk.Label(
                self.filter_content_frame,
                text="Build Your Custom Filters:",
                font=("Segoe UI", 11, "bold"),
                bg="#1a1a1a",
                fg="white"
            ).pack(anchor=tk.W, padx=10, pady=(10, 5))
            
            for i, filter_config in enumerate(endpoint.filters):
                self.create_dynamic_filter_builder(filter_config, i)
        else:
            tk.Label(
                self.filter_content_frame,
                text="No dynamic filters available for this endpoint.",
                font=("Segoe UI", 10),
                bg="#1a1a1a",
                fg="#cccccc"
            ).pack(anchor=tk.W, padx=10, pady=20)
    
    def create_dynamic_filter_builder(self, filter_config: Dict, index: int):
        """Create a dynamic filter builder based on filter configuration"""
        card_frame = tk.Frame(self.filter_content_frame, bg="#2a2a2a", relief=tk.RAISED, bd=1)
        card_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Header with filter name and description
        header_frame = tk.Frame(card_frame, bg="#2a2a2a")
        header_frame.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        name_label = tk.Label(
            header_frame,
            text=filter_config.get("name", "Filter"),
            font=("Segoe UI", 10, "bold"),
            bg="#2a2a2a",
            fg="#66b3ff"
        )
        name_label.pack(anchor=tk.W)
        
        desc_label = tk.Label(
            header_frame,
            text=filter_config.get("description", ""),
            font=("Segoe UI", 8),
            bg="#2a2a2a",
            fg="#cccccc"
        )
        desc_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Input area based on filter type
        input_frame = tk.Frame(card_frame, bg="#2a2a2a")
        input_frame.pack(fill=tk.X, padx=15, pady=(5, 10))
        
        filter_type = filter_config.get("type", "text")
        template = filter_config.get("template", "")
        
        # Store the input variables for this filter
        filter_vars = {}
        
        if filter_type == "boolean":
            self.create_boolean_input(input_frame, filter_config, filter_vars)
        elif filter_type == "text":
            self.create_text_input(input_frame, filter_config, filter_vars)
        elif filter_type == "email":
            self.create_email_input(input_frame, filter_config, filter_vars)
        elif filter_type == "number":
            self.create_number_input(input_frame, filter_config, filter_vars)
        elif filter_type == "datetime":
            self.create_datetime_input(input_frame, filter_config, filter_vars)
        elif filter_type == "select":
            self.create_select_input(input_frame, filter_config, filter_vars)
        elif filter_type == "multiselect":
            self.create_multiselect_input(input_frame, filter_config, filter_vars)
        elif filter_type == "compound":
            self.create_compound_input(input_frame, filter_config, filter_vars)
        elif filter_type == "static":
            self.create_static_input(input_frame, filter_config, filter_vars)
        
        # Action buttons and preview
        button_frame = tk.Frame(card_frame, bg="#2a2a2a")
        button_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        # Preview area
        preview_var = tk.StringVar()
        preview_label = tk.Label(
            button_frame,
            textvariable=preview_var,
            font=("Consolas", 8),
            bg="#1a1a1a",
            fg="#ffd43b",
            anchor="w",
            wraplength=500
        )
        preview_label.pack(fill=tk.X, pady=(0, 5))
        
        # Update preview function
        def update_preview():
            try:
                built_filter = self.build_filter_from_template(template, filter_vars, filter_config)
                preview_var.set(f"Preview: {built_filter}")
            except Exception as e:
                preview_var.set(f"Preview: Error - {str(e)}")
        
        # Bind update events
        self.bind_filter_updates(filter_vars, update_preview)
        update_preview()  # Initial preview
        
        # Action buttons
        actions_frame = tk.Frame(button_frame, bg="#2a2a2a")
        actions_frame.pack(fill=tk.X)
        
        add_btn = tk.Button(
            actions_frame,
            text="✓ Add Filter",
            font=("Segoe UI", 9),
            bg="#2a2a2a",
            fg="white",
            relief=tk.RAISED,
            bd=1,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=lambda: self.add_dynamic_filter(template, filter_vars, filter_config, add_btn)
        )
        add_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        copy_btn = tk.Button(
            actions_frame,
            text="📋 Copy",
            font=("Segoe UI", 9),
            bg="#2a2a2a",
            fg="white",
            relief=tk.RAISED,
            bd=1,
            highlightthickness=0,
            activebackground="#3a3a3a",
            activeforeground="white",
            command=lambda: self.copy_dynamic_filter(template, filter_vars, filter_config)
        )
        copy_btn.pack(side=tk.LEFT)
    
    def create_boolean_input(self, parent, config, filter_vars):
        """Create boolean input (True/False dropdown)"""
        tk.Label(parent, text="Value:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        var = tk.StringVar(value=config.get("default", "false"))
        filter_vars["value"] = var
        
        combo = ttk.Combobox(parent, textvariable=var, values=config.get("options", ["true", "false"]), 
                            state="readonly", width=10)
        combo.pack(side=tk.LEFT)
    
    def create_text_input(self, parent, config, filter_vars):
        """Create text input field"""
        tk.Label(parent, text="Text:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        var = tk.StringVar(value=config.get("default", ""))
        filter_vars["value"] = var
        
        entry = tk.Entry(parent, textvariable=var, font=("Segoe UI", 9), width=30, 
                        bg="#1a1a1a", fg="white", insertbackground="white")
        entry.pack(side=tk.LEFT)
    
    def create_email_input(self, parent, config, filter_vars):
        """Create email input field"""
        tk.Label(parent, text="Email:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        var = tk.StringVar(value=config.get("default", ""))
        filter_vars["email"] = var
        
        entry = tk.Entry(parent, textvariable=var, font=("Segoe UI", 9), width=30, 
                        bg="#1a1a1a", fg="white", insertbackground="white")
        entry.pack(side=tk.LEFT)
    
    def create_number_input(self, parent, config, filter_vars):
        """Create number input field with validation"""
        tk.Label(parent, text="Number:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        var = tk.StringVar(value=str(config.get("default", "1")))
        filter_vars["number"] = var
        
        entry = tk.Entry(parent, textvariable=var, font=("Segoe UI", 9), width=15, 
                        bg="#1a1a1a", fg="white", insertbackground="white")
        entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # Add min/max labels if specified
        min_val = config.get("min")
        max_val = config.get("max")
        if min_val is not None or max_val is not None:
            range_text = f"({min_val or 'no min'} - {max_val or 'no max'})"
            tk.Label(parent, text=range_text, font=("Segoe UI", 8), bg="#2a2a2a", fg="#888888").pack(side=tk.LEFT)
    
    def create_datetime_input(self, parent, config, filter_vars):
        """Create datetime input field"""
        tk.Label(parent, text="Date/Time:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        var = tk.StringVar(value=config.get("default", "2025-01-15T00:00:00Z"))
        filter_vars["datetime"] = var
        
        entry = tk.Entry(parent, textvariable=var, font=("Segoe UI", 9), width=25, 
                        bg="#1a1a1a", fg="white", insertbackground="white")
        entry.pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Label(parent, text="(YYYY-MM-DDTHH:MM:SSZ)", font=("Segoe UI", 8), bg="#2a2a2a", fg="#888888").pack(side=tk.LEFT)
    
    def create_select_input(self, parent, config, filter_vars):
        """Create single select dropdown"""
        tk.Label(parent, text="Option:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        var = tk.StringVar(value=config.get("default", config.get("options", [""])[0]))
        filter_vars["value"] = var
        
        combo = ttk.Combobox(parent, textvariable=var, values=config.get("options", []), 
                            state="readonly", width=15)
        combo.pack(side=tk.LEFT)
    
    def create_multiselect_input(self, parent, config, filter_vars):
        """Create multi-select checkboxes"""
        tk.Label(parent, text="Fields:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(anchor=tk.W, pady=(0, 5))
        
        checkboxes_frame = tk.Frame(parent, bg="#2a2a2a")
        checkboxes_frame.pack(fill=tk.X)
        
        options = config.get("options", [])
        defaults = config.get("default", [])
        checkbox_vars = {}
        
        for i, option in enumerate(options):
            var = tk.BooleanVar(value=option in defaults)
            checkbox_vars[option] = var
            
            cb = tk.Checkbutton(
                checkboxes_frame,
                text=option,
                variable=var,
                font=("Segoe UI", 8),
                bg="#2a2a2a",
                fg="white",
                selectcolor="#1a1a1a",
                activebackground="#2a2a2a",
                activeforeground="white"
            )
            cb.grid(row=i // 3, column=i % 3, sticky=tk.W, padx=10, pady=2)
        
        filter_vars["fields"] = checkbox_vars
    
    def create_compound_input(self, parent, config, filter_vars):
        """Create compound input with multiple fields"""
        fields_config = config.get("fields", {})
        
        for field_name, field_config in fields_config.items():
            field_frame = tk.Frame(parent, bg="#2a2a2a")
            field_frame.pack(fill=tk.X, pady=2)
            
            label_text = field_name.replace("_", " ").title()
            tk.Label(field_frame, text=f"{label_text}:", font=("Segoe UI", 9), bg="#2a2a2a", fg="white").pack(side=tk.LEFT, padx=(0, 5))
            
            if field_config.get("type") == "select":
                var = tk.StringVar(value=field_config.get("default", field_config.get("options", [""])[0]))
                filter_vars[field_name] = var
                
                combo = ttk.Combobox(field_frame, textvariable=var, values=field_config.get("options", []), 
                                    state="readonly", width=15)
                combo.pack(side=tk.LEFT)
    
    def create_static_input(self, parent, config, filter_vars):
        """Create static filter (no input needed)"""
        tk.Label(parent, text="This filter requires no input parameters.", 
                font=("Segoe UI", 9, "italic"), bg="#2a2a2a", fg="#888888").pack()
        filter_vars["static"] = True
    
    def bind_filter_updates(self, filter_vars, update_callback):
        """Bind update events to filter variables"""
        for key, var in filter_vars.items():
            if isinstance(var, (tk.StringVar, tk.BooleanVar)):
                var.trace_add("write", lambda *args: update_callback())
            elif isinstance(var, dict):  # For multiselect checkboxes
                for checkbox_var in var.values():
                    if isinstance(checkbox_var, tk.BooleanVar):
                        checkbox_var.trace_add("write", lambda *args: update_callback())
    
    def build_filter_from_template(self, template: str, filter_vars: Dict, config: Dict) -> str:
        """Build the actual filter string from template and user inputs"""
        filter_type = config.get("type", "text")
        
        if filter_type == "static":
            return template
        elif filter_type == "multiselect":
            # Build comma-separated list of selected fields
            checkbox_vars = filter_vars.get("fields", {})
            selected = [field for field, var in checkbox_vars.items() if var.get()]
            if not selected:
                selected = ["id"]  # Default to at least one field
            return template.format(fields=",".join(selected))
        elif filter_type == "compound":
            # Replace all field placeholders
            result = template
            for field_name, var in filter_vars.items():
                if isinstance(var, tk.StringVar):
                    result = result.replace(f"{{{field_name}}}", var.get())
            return result
        else:
            # Simple template replacement
            replacements = {}
            for key, var in filter_vars.items():
                if isinstance(var, (tk.StringVar, tk.BooleanVar)):
                    replacements[key] = var.get()
            return template.format(**replacements)
    
    def add_dynamic_filter(self, template: str, filter_vars: Dict, config: Dict, button: tk.Button):
        """Add the dynamic filter to selected filters"""
        try:
            built_filter = self.build_filter_from_template(template, filter_vars, config)
            
            if built_filter in self.selected_filters:
                self.selected_filters.remove(built_filter)
                button.config(text="✓ Add Filter", bg="#2a2a2a")
            else:
                self.selected_filters.append(built_filter)
                button.config(text="✗ Remove", bg="#404040")
            
            # Update selected filters display
            if self.selected_filters:
                filters_text = " & ".join(self.selected_filters)
                if len(filters_text) > 100:
                    filters_text = filters_text[:100] + "..."
                self.selected_filters_label.config(text=f"Selected Filters: {filters_text}")
            else:
                self.selected_filters_label.config(text="Selected Filters: None")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to build filter: {str(e)}")
    
    def copy_dynamic_filter(self, template: str, filter_vars: Dict, config: Dict):
        """Copy the dynamic filter URL to clipboard"""
        try:
            built_filter = self.build_filter_from_template(template, filter_vars, config)
            
            if self.selected_endpoint:
                # Use folder-specific URL if applicable
                base_url = self.selected_endpoint.url
                if "messages" in base_url.lower() and self.selected_folder:
                    folder_id = self.selected_folder.get('id')
                    base_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages"
                
                full_url = base_url + built_filter
                try:
                    pyperclip.copy(full_url)
                    messagebox.showinfo("Copied", f"Filter URL copied to clipboard!\n\n{full_url}")
                except Exception:
                    # Fallback if pyperclip is not available
                    self.root.clipboard_clear()
                    self.root.clipboard_append(full_url)
                    self.root.update()
                    messagebox.showinfo("Copied", f"Filter URL copied to clipboard!\n\n{full_url}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy filter: {str(e)}")
    
    def execute_selected_filters(self):
        """Execute the selected filters and show results"""
        if not self.selected_endpoint:
            messagebox.showwarning("Warning", "No endpoint selected.")
            return
        
        if not self.access_token:
            messagebox.showwarning("Warning", "Please authenticate with MSAL first.")
            return
        
        # Construct the query URL
        base_url = self.selected_endpoint.url
        
        # If this is a messages endpoint and a folder is selected, use the folder
        if "messages" in base_url.lower() and self.selected_folder:
            folder_id = self.selected_folder.get('id')
            base_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages"
        
        if self.selected_filters:
            # Combine filters intelligently
            combined_filters = "&".join([f.lstrip('?') for f in self.selected_filters])

            # Check if a $top parameter is already included
            has_top_param = any('$top=' in filter_str for filter_str in self.selected_filters)

            # If no $top parameter is specified, add a reasonable default to get more results
            if not has_top_param:
                combined_filters += "&$top=100"

            full_url = f"{base_url}?{combined_filters}"
        else:
            # No filters selected, add default $top to get more than the API default
            full_url = f"{base_url}?$top=100"
        
        # Show results popup
        self.show_results_popup(full_url)
    
    def show_results_popup(self, query_url: str):
        """Show popup window with query results"""
        popup = tk.Toplevel(self.root)
        popup.title("Query Results - Dynamic Filters")
        popup.geometry("1000x700")
        popup.configure(bg="#000000")
        
        # Header
        header_frame = tk.Frame(popup, bg="#1a1a1a", height=80)
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        header_frame.pack_propagate(False)
        
        header_text = "🚀 Query Results with Custom Filters"
        if self.selected_folder:
            folder_name = self.selected_folder.get('hierarchyName', 'Unknown')
            header_text += f" from '{folder_name}'"
        
        tk.Label(
            header_frame,
            text=header_text,
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#1a1a1a"
        ).pack(pady=20)
        
        # Query URL display
        url_frame = tk.Frame(popup, bg="#2a2a2a", relief=tk.RAISED, bd=1)
        url_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        
        tk.Label(
            url_frame,
            text="Query URL:",
            font=("Segoe UI", 10, "bold"),
            bg="#2a2a2a",
            fg="white"
        ).pack(anchor=tk.W, padx=10, pady=(5, 2))
        
        url_text = scrolledtext.ScrolledText(
            url_frame,
            height=3,
            font=("Consolas", 9),
            bg="#1a1a1a",
            fg="#cccccc",
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        url_text.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        url_text.config(state=tk.NORMAL)
        url_text.insert(tk.END, query_url)
        url_text.config(state=tk.DISABLED)
        
        # Results area
        results_frame = tk.LabelFrame(popup, text="Results", font=("Segoe UI", 12, "bold"), 
                                     bg="#000000", fg="white")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        results_text = scrolledtext.ScrolledText(
            results_frame,
            font=("Consolas", 9),
            bg="#1a1a1a",
            fg="#cccccc",
            wrap=tk.WORD
        )
        results_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Execute the query
        if MSAL_AVAILABLE and self.access_token:
            try:
                headers = {"Authorization": f"Bearer {self.access_token}"}
                
                # Show loading message
                results_text.insert(tk.END, "🔄 Executing Microsoft Graph query with custom filters...\n\n")
                popup.update()
                
                # Execute the query
                response = requests.get(query_url, headers=headers)
                
                if response.ok:
                    data = response.json()
                    results_text.delete(1.0, tk.END)  # Clear loading message
                    results_text.insert(tk.END, f"✅ Query executed successfully!\n")
                    results_text.insert(tk.END, f"Status Code: {response.status_code}\n\n")
                    
                    # Display results based on endpoint type
                    if "messages" in query_url:
                        self.display_messages_results(data, results_text)
                    elif "mailFolders" in query_url:
                        self.display_folders_results(data, results_text)
                    else:
                        results_text.insert(tk.END, f"Raw Response:\n{json.dumps(data, indent=2)}")
                else:
                    results_text.delete(1.0, tk.END)  # Clear loading message
                    results_text.insert(tk.END, f"❌ Query failed!\n")
                    results_text.insert(tk.END, f"Status Code: {response.status_code}\n")
                    results_text.insert(tk.END, f"Error: {response.text}\n")
                    
            except Exception as e:
                results_text.delete(1.0, tk.END)  # Clear loading message
                results_text.insert(tk.END, f"❌ Error executing query: {str(e)}\n")
        else:
            results_text.insert(tk.END, "❌ MSAL not available or not authenticated.\n")
    
    def display_messages_results(self, data: Dict, results_text):
        messages = data.get("value", [])

        folder_name = self.selected_folder.get('hierarchyName', 'All Messages') if self.selected_folder else "All Messages"
        
        results_text.insert(tk.END, f"📧 Found {len(messages)} messages in '{folder_name}':\n")
        results_text.insert(tk.END, "=" * 50 + "\n\n")

        for idx, msg in enumerate(messages):
            full_body = msg.get('body', {}).get('content', {})
            # results_text.insert(tk.END, f"Body:    {full_body}\n")
            soup = BeautifulSoup(full_body, 'html.parser')
            soup_text = soup.get_text("|").split("|")
            soup_text.append(soup.a["href"]) # type: ignore
            results_text.insert(tk.END, f"{idx}\n")
            results_text.insert(tk.END, f"{soup_text[1]}\n")
            results_text.insert(tk.END, f"{soup_text[9]}\n")
            results_text.insert(tk.END, f"{soup_text[11]}\n")
            results_text.insert(tk.END, f"{soup_text[13]}\n")
            results_text.insert(tk.END, f"{soup_text[15]}\n")
            results_text.insert(tk.END, f"{soup_text[17]}\n")
            results_text.insert(tk.END, f"{soup_text[-1]}\n")
            results_text.insert(tk.END, "\n")
    
    def display_folders_results(self, data: Dict, results_text):
        """Display mail folders results"""
        folders = data.get("value", [])
        results_text.insert(tk.END, f"📁 Found {len(folders)} mail folders:\n")
        results_text.insert(tk.END, "=" * 50 + "\n\n")
        
        for folder in folders:
            results_text.insert(tk.END, f"Folder: {folder.get('displayName', 'Unknown')}\n")
            results_text.insert(tk.END, f"  ID: {folder.get('id', 'Unknown')}\n")
            results_text.insert(tk.END, f"  Total Items: {folder.get('totalItemCount', 'Unknown')}\n")
            results_text.insert(tk.END, f"  Unread Items: {folder.get('unreadItemCount', 'Unknown')}\n")
            results_text.insert(tk.END, "-" * 30 + "\n\n")
    
    def filter_endpoints(self, *args):
        """Filter endpoints based on current filter settings"""
        search_term = self.search_var.get().lower()
        category = self.category_var.get()
        version = self.version_var.get()
        
        self.filtered_endpoints = []
        
        for endpoint in self.endpoints:
            # Search filter
            if search_term and not any(search_term in text.lower() for text in [
                endpoint.name, endpoint.description, endpoint.url
            ]):
                continue
            
            # Category filter
            if category != "All" and endpoint.category != category:
                continue
            
            # Version filter
            if version != "All" and endpoint.version != version:
                continue
            
            self.filtered_endpoints.append(endpoint)
        
        self.display_results()
    
    def display_results(self):
        """Display filtered results with dark theme"""
        # Clear previous results
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # Update count
        count = len(self.filtered_endpoints)
        self.count_label.config(text=f"Showing {count} endpoint{'s' if count != 1 else ''}")
        
        if not self.filtered_endpoints:
            no_results = tk.Label(
                self.scrollable_frame,
                text="No endpoints match your criteria.",
                font=("Segoe UI", 12),
                fg="#cccccc",
                bg="#1a1a1a"
            )
            no_results.pack(pady=50)
            return
        
        # Display endpoints
        for i, endpoint in enumerate(self.filtered_endpoints):
            self.create_endpoint_card(endpoint, i)
    
    def create_endpoint_card(self, endpoint: GraphEndpoint, index: int):
        """Create a card for displaying endpoint information"""
        # Card frame
        card_frame = tk.Frame(self.scrollable_frame, bg="#2a2a2a", relief=tk.RAISED, bd=1)
        card_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Make card clickable to show filtering options
        def on_card_click(event=None):
            self.show_filtering_options(endpoint)
        
        card_frame.bind("<Button-1>", on_card_click)
        
        # Header with method and category badges
        header_frame = tk.Frame(card_frame, bg="#2a2a2a")
        header_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        header_frame.bind("<Button-1>", on_card_click)
        
        # Method badge
        method_colors = {"GET": "#0078d4", "POST": "#107c10", "PUT": "#ff8c00", "DELETE": "#d13438"}
        method_color = method_colors.get(endpoint.method, "#666666")
        
        method_label = tk.Label(
            header_frame,
            text=endpoint.method,
            bg=method_color,
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=2
        )
        method_label.pack(side=tk.LEFT, padx=(0, 10))
        method_label.bind("<Button-1>", on_card_click)
        
        # Category badge
        category_label = tk.Label(
            header_frame,
            text=endpoint.category,
            bg="#444444",
            fg="white",
            font=("Segoe UI", 9),
            padx=8,
            pady=2
        )
        category_label.pack(side=tk.LEFT)
        category_label.bind("<Button-1>", on_card_click)
        
        # Dynamic filters count
        filter_count = len(endpoint.filters) if endpoint.filters else 0
        filter_count_label = tk.Label(
            header_frame,
            text=f"{filter_count} Dynamic Filters",
            font=("Segoe UI", 9),
            bg="#2a2a2a",
            fg="#ffd43b"
        )
        filter_count_label.pack(side=tk.RIGHT)
        filter_count_label.bind("<Button-1>", on_card_click)
        
        # Endpoint name and description
        name_label = tk.Label(
            card_frame,
            text=endpoint.name,
            font=("Segoe UI", 12, "bold"),
            bg="#2a2a2a",
            fg="white"
        )
        name_label.pack(anchor=tk.W, padx=15, pady=(5, 2))
        name_label.bind("<Button-1>", on_card_click)
        
        desc_label = tk.Label(
            card_frame,
            text=endpoint.description,
            font=("Segoe UI", 10),
            bg="#2a2a2a",
            fg="#cccccc",
            wraplength=400,
            justify=tk.LEFT
        )
        desc_label.pack(anchor=tk.W, padx=15, pady=(0, 10))
        desc_label.bind("<Button-1>", on_card_click)
        
        # Click instruction
        click_label = tk.Label(
            card_frame,
            text="→ Click to build custom filters",
            font=("Segoe UI", 8, "italic"),
            bg="#2a2a2a",
            fg="#888888"
        )
        click_label.pack(anchor=tk.E, padx=15, pady=(0, 10))
        click_label.bind("<Button-1>", on_card_click)
    
    def clear_selected_filters(self):
        """Clear all selected filters and reset filter button states"""
        if not self.selected_filters:
            messagebox.showinfo("Info", "No filters are currently selected.")
            return

        # Clear the selected filters list
        self.selected_filters.clear()

        # Update the selected filters display
        self.selected_filters_label.config(text="Selected Filters: None")

        # Reset all "Add Filter" buttons back to their original state
        # We need to refresh the filtering panel to reset button states
        if self.selected_endpoint:
            self.show_filtering_options(self.selected_endpoint)

        messagebox.showinfo("Cleared", "All selected filters have been cleared.")

    def copy_to_clipboard(self, text: str):
        """Copy text to clipboard"""
        try:
            pyperclip.copy(text)
            messagebox.showinfo("Copied", "URL copied to clipboard!")
        except Exception as e:
            # Fallback if pyperclip is not available
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()
            messagebox.showinfo("Copied", "URL copied to clipboard!")

def main():
    """Main application entry point"""
    root = tk.Tk()
    app = MicrosoftGraphExplorer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
