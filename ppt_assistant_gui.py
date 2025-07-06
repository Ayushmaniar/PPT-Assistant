import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import tkinter.font as tkfont
import re
import datetime

import importlib.util
import sys

# Dynamically import ppt_smolagent.py as ppt_smolagent
import win32com.client
import pythoncom
spec = importlib.util.spec_from_file_location("ppt_smolagent", "ppt_smolagent.py")
ppt_smolagent = importlib.util.module_from_spec(spec)
sys.modules["ppt_smolagent"] = ppt_smolagent
spec.loader.exec_module(ppt_smolagent)

class PPTAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT Assistant Chat UI")
        self.ppt_app = None
        self.presentation = None
        self.setup_ui()

    def setup_ui(self):
        # Modern color scheme with gradients and shadows
        self.bg_color = "#0f0f0f"  # Deeper black
        self.card_bg = "#1a1a1a"  # Card background
        self.user_msg_bg = "#2563eb"  # Modern blue
        self.user_msg_fg = "#ffffff"
        self.sys_msg_bg = "#1f2937"  # Dark gray
        self.sys_msg_fg = "#f9fafb"
        self.entry_bg = "#ffffff"  # White background for better contrast with black text
        self.entry_fg = "#000000"  # Black text for better visibility
        self.btn_bg = "#3b82f6"  # Modern blue
        self.btn_hover_bg = "#2563eb"  # Darker blue for hover
        self.btn_fg = "#ffffff"
        self.code_bg = "#111827"
        self.code_fg = "#e5e7eb"
        self.accent_color = "#06b6d4"  # Cyan accent
        self.border_color = "#374151"

        self.root.configure(bg=self.bg_color)
        
        # Configure window properties for modern look
        self.root.geometry("1200x800")  # Increased size for tabs
        self.root.minsize(900, 600)

        # Create main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Style the notebook
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook', background=self.bg_color)
        style.configure('TNotebook.Tab', padding=[12, 8], background=self.card_bg, foreground=self.sys_msg_fg)
        style.map('TNotebook.Tab', background=[('selected', self.accent_color)], foreground=[('selected', '#000000')])

        # Create Chat Tab
        self.chat_frame = tk.Frame(self.notebook, bg=self.bg_color)
        self.notebook.add(self.chat_frame, text="💬 Chat Assistant")
        self.setup_chat_tab()

        # Create Debug Tab
        self.debug_frame = tk.Frame(self.notebook, bg=self.bg_color)
        self.notebook.add(self.debug_frame, text="🔧 Debug Console")
        self.setup_debug_tab()

    def setup_chat_tab(self):
        """Setup the chat interface tab"""

        # Main container with padding and modern styling
        main_frame = tk.Frame(self.chat_frame, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header section
        header_frame = tk.Frame(main_frame, bg=self.card_bg, relief="flat")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Add subtle border effect
        header_border = tk.Frame(header_frame, bg=self.accent_color, height=3)
        header_border.pack(fill=tk.X, side=tk.TOP)
        
        # Check if lightning-fast optimization is available
        try:
            from lightning_slide_context_reader import LightningFastPowerPointSlideReader
            title_text = "💼 PPT Assistant ⚡ Lightning-Fast"
            title_color = "#22c55e"  # Green for optimization active
        except ImportError:
            title_text = "💼 PPT Assistant"
            title_color = self.accent_color
        
        title_label = tk.Label(header_frame, text=title_text, 
                              bg=self.card_bg, fg=title_color, 
                              font=("Segoe UI", 18, "bold"), pady=15)
        title_label.pack()

        # Chat container with modern card design
        chat_container = tk.Frame(main_frame, bg=self.card_bg, relief="flat")
        chat_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Chat header
        chat_header = tk.Frame(chat_container, bg=self.card_bg)
        chat_header.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        chat_title = tk.Label(chat_header, text="💬 Conversation", 
                             bg=self.card_bg, fg=self.sys_msg_fg, 
                             font=("Segoe UI", 12, "bold"))
        chat_title.pack(anchor="w")

        # Chat area with modern scrollbar
        self.chat_area = scrolledtext.ScrolledText(
            chat_container, 
            state='normal', 
            width=60, 
            height=15, 
            bg=self.bg_color, 
            fg=self.sys_msg_fg, 
            font=("Segoe UI", 11), 
            bd=0, 
            highlightthickness=0, 
            wrap=tk.WORD,
            padx=15,
            pady=10,
            relief="flat"
        )
        self.chat_area.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        # Make chat area read-only but selectable with copy functionality
        self.chat_area.bind("<Key>", self.handle_chat_key_event)  # Handle key events
        self.chat_area.bind("<Button-1>", lambda e: self.chat_area.focus_set())  # Allow focus for selection

        # Code display section with modern card design
        self.code_frame = tk.Frame(main_frame, bg=self.card_bg, relief="flat")
        self.code_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Code header with toggle
        code_header = tk.Frame(self.code_frame, bg=self.card_bg)
        code_header.pack(fill=tk.X, padx=15, pady=15)
        
        # Modern toggle button with icon
        self.code_toggle_btn = tk.Button(
            code_header, 
            text="▶ Generated Code", 
            command=self.toggle_code_display, 
            bg=self.sys_msg_bg, 
            fg=self.sys_msg_fg, 
            font=("Segoe UI", 10, "bold"), 
            bd=0, 
            padx=20,
            pady=8,
            cursor="hand2",
            relief="flat",
            activebackground=self.border_color,
            activeforeground=self.sys_msg_fg
        )
        self.code_toggle_btn.pack(anchor='w')
        
        # Code display area with syntax highlighting
        self.code_display = scrolledtext.ScrolledText(
            self.code_frame, 
            state='normal', 
            width=60, 
            height=8, 
            bg=self.code_bg, 
            fg=self.code_fg, 
            font=("JetBrains Mono", 10) if self.is_font_available("JetBrains Mono") else ("Consolas", 10), 
            bd=0, 
            highlightthickness=0, 
            wrap=tk.WORD,
            padx=15,
            pady=10,
            relief="flat"
        )
        self.code_display_visible = False
        
        # Make code display read-only but selectable with copy functionality
        self.code_display.bind("<Key>", self.handle_code_key_event)  # Handle key events
        self.code_display.bind("<Button-1>", lambda e: self.code_display.focus_set())  # Allow focus for selection

        # Modern input section
        input_container = tk.Frame(self.chat_frame, bg=self.card_bg, relief="flat")
        input_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Input frame with modern styling
        entry_frame = tk.Frame(input_container, bg=self.card_bg)
        entry_frame.pack(fill=tk.X, padx=15, pady=15)

        # Custom entry with placeholder effect
        self.entry = tk.Entry(
            entry_frame, 
            width=50, 
            bg=self.entry_bg, 
            fg=self.entry_fg, 
            insertbackground=self.accent_color, 
            font=("Segoe UI", 12), 
            bd=0, 
            highlightthickness=2,
            highlightcolor=self.accent_color,
            highlightbackground=self.border_color,
            relief="flat"
        )
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=8)
        self.entry.bind('<Return>', lambda event: self.send_message())
        self.entry.bind('<FocusIn>', self.on_entry_focus_in)
        self.entry.bind('<FocusOut>', self.on_entry_focus_out)

        # Modern button styling
        button_style = {
            "font": ("Segoe UI", 10, "bold"),
            "bd": 0,
            "cursor": "hand2",
            "relief": "flat",
            "padx": 20,
            "pady": 10
        }

        send_btn = tk.Button(
            entry_frame, 
            text="Send 📤", 
            command=self.send_message, 
            bg=self.btn_bg, 
            fg=self.btn_fg,
            activebackground=self.btn_hover_bg, 
            activeforeground=self.btn_fg,
            **button_style
        )
        send_btn.pack(side=tk.LEFT, padx=(0, 10))

        new_btn = tk.Button(
            entry_frame, 
            text="New 📄", 
            command=self.create_new_ppt, 
            bg=self.sys_msg_bg, 
            fg=self.sys_msg_fg,
            activebackground=self.border_color, 
            activeforeground=self.sys_msg_fg,
            **button_style
        )
        new_btn.pack(side=tk.LEFT, padx=(0, 10))

        open_btn = tk.Button(
            entry_frame, 
            text="Open 📂", 
            command=self.open_ppt, 
            bg=self.sys_msg_bg, 
            fg=self.sys_msg_fg,
            activebackground=self.border_color, 
            activeforeground=self.sys_msg_fg,
            **button_style
        )
        open_btn.pack(side=tk.LEFT)

        # Add hover effects
        self.add_hover_effects()
        
        # Set placeholder text
        self.set_entry_placeholder()

    def setup_debug_tab(self):
        """Setup the debug console tab"""
        # Main container with padding
        main_frame = tk.Frame(self.debug_frame, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header section
        header_frame = tk.Frame(main_frame, bg=self.card_bg, relief="flat")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Add subtle border effect
        header_border = tk.Frame(header_frame, bg="#dc2626", height=3)  # Red accent for debug
        header_border.pack(fill=tk.X, side=tk.TOP)
        
        title_label = tk.Label(header_frame, text="🔧 Debug Console", 
                              bg=self.card_bg, fg="#dc2626", 
                              font=("Segoe UI", 18, "bold"), pady=15)
        title_label.pack()

        # Create horizontal container for left and right panels
        content_frame = tk.Frame(main_frame, bg=self.bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Left panel for code editor and controls
        left_panel = tk.Frame(content_frame, bg=self.bg_color)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Right panel for slide context (fixed width)
        right_panel = tk.Frame(content_frame, bg=self.card_bg, width=350, relief="flat")
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_panel.pack_propagate(False)  # Maintain fixed width

        # === LEFT PANEL: Code Editor ===
        # Code editor section
        editor_container = tk.Frame(left_panel, bg=self.card_bg, relief="flat")
        editor_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Editor header
        editor_header = tk.Frame(editor_container, bg=self.card_bg)
        editor_header.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        editor_title = tk.Label(editor_header, text="📝 Code Editor", 
                               bg=self.card_bg, fg=self.sys_msg_fg, 
                               font=("Segoe UI", 12, "bold"))
        editor_title.pack(anchor="w")

        # Toolbar for editor
        toolbar_frame = tk.Frame(editor_header, bg=self.card_bg)
        toolbar_frame.pack(fill=tk.X, pady=(5, 0))

        # Example templates dropdown
        template_label = tk.Label(toolbar_frame, text="Templates:", 
                                 bg=self.card_bg, fg=self.sys_msg_fg, 
                                 font=("Segoe UI", 9))
        template_label.pack(side=tk.LEFT, padx=(0, 5))

        self.template_var = tk.StringVar(value="Select Template")
        template_dropdown = ttk.Combobox(toolbar_frame, textvariable=self.template_var, 
                                        values=[
                                            "Select Template",
                                            "Color Text Pattern Example",
                                            "Debug HTML Test",
                                            "Replace Textbox Content Example",
                                            "Modify Text in Textbox Example", 
                                            "Add Text to Textbox Example",
                                            "Format Textbox Style Example",
                                            "Add New Textbox Example",
                                            "Move and Resize Object Example",
                                            "Get Object Properties Example",
                                            "Copy Object to Slide Example",
                                            "Duplicate Object Example"
                                        ], 
                                        state="readonly", width=30)
        template_dropdown.pack(side=tk.LEFT, padx=(0, 10))
        template_dropdown.bind('<<ComboboxSelected>>', self.load_template)

        # Clear button
        clear_btn = tk.Button(toolbar_frame, text="Clear", command=self.clear_debug_editor,
                             bg=self.sys_msg_bg, fg=self.sys_msg_fg, font=("Segoe UI", 9),
                             bd=0, padx=10, pady=5, cursor="hand2", relief="flat")
        clear_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Code editor area with undo/redo support
        self.debug_editor = scrolledtext.ScrolledText(
            editor_container, 
            state='normal', 
            width=80, 
            height=25, 
            bg=self.code_bg, 
            fg=self.code_fg, 
            font=("JetBrains Mono", 11) if self.is_font_available("JetBrains Mono") else ("Consolas", 11), 
            bd=0, 
            highlightthickness=1,
            highlightcolor="#dc2626",
            wrap=tk.NONE,  # No word wrap for code
            padx=15,
            pady=10,
            relief="flat",
            insertbackground="#dc2626",  # Red cursor
            undo=True,  # Enable undo/redo
            maxundo=50  # Set maximum undo levels
        )
        self.debug_editor.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

        # Bind undo/redo keys
        self.debug_editor.bind('<Control-z>', lambda e: self.debug_editor.edit_undo())
        self.debug_editor.bind('<Control-Z>', lambda e: self.debug_editor.edit_undo())
        self.debug_editor.bind('<Control-y>', lambda e: self.debug_editor.edit_redo())
        self.debug_editor.bind('<Control-Y>', lambda e: self.debug_editor.edit_redo())

        # Add global keyboard shortcut for refreshing context (F5 or Ctrl+R)
        self.root.bind('<F5>', lambda e: self.refresh_slide_context_with_feedback())
        self.root.bind('<Control-r>', lambda e: self.refresh_slide_context_with_feedback())
        self.root.bind('<Control-R>', lambda e: self.refresh_slide_context_with_feedback())

        # Control buttons section
        controls_frame = tk.Frame(left_panel, bg=self.card_bg, relief="flat")
        controls_frame.pack(fill=tk.X, pady=(0, 15))

        button_frame = tk.Frame(controls_frame, bg=self.card_bg)
        button_frame.pack(padx=15, pady=15)

        # Execute button (prominent)
        execute_btn = tk.Button(
            button_frame, 
            text="▶ Execute Code", 
            command=self.execute_debug_code, 
            bg="#059669",  # Green for execute
            fg=self.btn_fg,
            font=("Segoe UI", 12, "bold"), 
            bd=0, 
            padx=25,
            pady=12,
            cursor="hand2",
            relief="flat",
            activebackground="#047857",
            activeforeground=self.btn_fg
        )
        execute_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Output section (much larger height)
        output_container = tk.Frame(left_panel, bg=self.card_bg, relief="flat")
        output_container.pack(fill=tk.BOTH, expand=True)
        
        # Output header
        output_header = tk.Frame(output_container, bg=self.card_bg)
        output_header.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        output_title = tk.Label(output_header, text="📤 Output", 
                               bg=self.card_bg, fg=self.sys_msg_fg, 
                               font=("Segoe UI", 12, "bold"))
        output_title.pack(anchor="w")

        # Output area (much larger)
        self.debug_output = scrolledtext.ScrolledText(
            output_container, 
            state='normal', 
            width=80, 
            height=20,  # Increased from 8 to 20
            bg="#0f172a",  # Darker for output
            fg="#e2e8f0", 
            font=("JetBrains Mono", 10) if self.is_font_available("JetBrains Mono") else ("Consolas", 10), 
            bd=0, 
            highlightthickness=0, 
            wrap=tk.WORD,
            padx=15,
            pady=10,
            relief="flat"
        )
        self.debug_output.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        # Make output read-only but allow copy
        self.debug_output.bind("<Key>", self.handle_debug_output_key_event)
        self.debug_output.bind("<Button-1>", lambda e: self.debug_output.focus_set())

        # === RIGHT PANEL: Slide Context ===
        # Context header
        context_header = tk.Frame(right_panel, bg=self.card_bg)
        context_header.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        # Add border effect for context panel
        context_border = tk.Frame(context_header, bg=self.btn_bg, height=2)
        context_border.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        
        context_title = tk.Label(context_header, text="� Current Slide Context", 
                                bg=self.card_bg, fg=self.sys_msg_fg, 
                                font=("Segoe UI", 12, "bold"))
        context_title.pack(anchor="w")

        # Refresh context button
        self.refresh_context_btn = tk.Button(
            context_header, 
            text="🔄 Refresh Context (F5)", 
            command=self.refresh_slide_context_with_feedback, 
            bg=self.btn_bg, 
            fg=self.btn_fg,
            font=("Segoe UI", 9, "bold"), 
            bd=0, 
            padx=15,
            pady=8,
            cursor="hand2",
            relief="flat",
            activebackground=self.btn_hover_bg,
            activeforeground=self.btn_fg
        )
        self.refresh_context_btn.pack(anchor="w", pady=(5, 0))

        # Add shortcut info
        shortcut_info = tk.Label(context_header, text="💡 Shortcuts: F5 or Ctrl+R", 
                                bg=self.card_bg, fg="#9ca3af", 
                                font=("Segoe UI", 8))
        shortcut_info.pack(anchor="w", pady=(2, 0))

        # Context display area (vertical wide format)
        self.context_display = scrolledtext.ScrolledText(
            right_panel, 
            state='normal', 
            width=40,
            bg="#1e293b",  # Slightly different shade
            fg="#f1f5f9", 
            font=("JetBrains Mono", 9) if self.is_font_available("JetBrains Mono") else ("Consolas", 9), 
            bd=0, 
            highlightthickness=0, 
            wrap=tk.WORD,
            padx=10,
            pady=10,
            relief="flat"
        )
        self.context_display.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        # Make context display read-only but allow copy
        self.context_display.bind("<Key>", self.handle_debug_output_key_event)
        self.context_display.bind("<Button-1>", lambda e: self.context_display.focus_set())

        # Auto-refresh context on tab switch
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)

    def is_font_available(self, font_name):
        """Check if a font is available on the system."""
        try:
            import tkinter.font as tkfont
            available_fonts = tkfont.families()
            return font_name in available_fonts
        except:
            return False

    def set_entry_placeholder(self):
        """Set placeholder text for the entry field."""
        self.placeholder_text = "Type your PowerPoint request here..."
        self.entry.insert(0, self.placeholder_text)
        self.entry.config(fg="#888888")  # Gray placeholder text against white background

    def on_entry_focus_in(self, event):
        """Handle entry field focus in - remove placeholder."""
        if self.entry.get() == self.placeholder_text:
            self.entry.delete(0, tk.END)
            self.entry.config(fg=self.entry_fg)

    def on_entry_focus_out(self, event):
        """Handle entry field focus out - add placeholder if empty."""
        if not self.entry.get():
            self.entry.insert(0, self.placeholder_text)
            self.entry.config(fg="#888888")  # Gray placeholder text

    def add_hover_effects(self):
        """Add modern hover effects to buttons."""
        def on_enter(event, widget, hover_bg, hover_fg=None):
            widget.config(bg=hover_bg)
            if hover_fg:
                widget.config(fg=hover_fg)

        def on_leave(event, widget, normal_bg, normal_fg):
            widget.config(bg=normal_bg, fg=normal_fg)

        # Find all buttons and add hover effects
        for widget in self.root.winfo_children():
            self._add_hover_to_frame(widget)

    def _add_hover_to_frame(self, frame):
        """Recursively add hover effects to buttons in frames."""
        for widget in frame.winfo_children():
            if isinstance(widget, tk.Button):
                if "Send" in widget.cget("text"):
                    widget.bind("<Enter>", lambda e, w=widget: w.config(bg=self.btn_hover_bg))
                    widget.bind("<Leave>", lambda e, w=widget: w.config(bg=self.btn_bg))
                elif "Generated Code" in widget.cget("text"):
                    widget.bind("<Enter>", lambda e, w=widget: w.config(bg=self.border_color))
                    widget.bind("<Leave>", lambda e, w=widget: w.config(bg=self.sys_msg_bg))
                else:
                    widget.bind("<Enter>", lambda e, w=widget: w.config(bg=self.border_color))
                    widget.bind("<Leave>", lambda e, w=widget: w.config(bg=self.sys_msg_bg))
            elif isinstance(widget, tk.Frame):
                self._add_hover_to_frame(widget)

    def toggle_code_display(self):
        """Toggle the visibility of the code display area with smooth animation."""
        if self.code_display_visible:
            # Hide code display
            self.code_display.pack_forget()
            self.code_toggle_btn.config(text="▶ Generated Code")
            self.code_display_visible = False
        else:
            # Show code display
            self.code_display.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
            self.code_toggle_btn.config(text="▼ Generated Code")
            self.code_display_visible = True

    def handle_chat_key_event(self, event):
        """Handle key events for chat area - allow copy operations but prevent typing."""
        # Allow copy operations (Ctrl+C, Ctrl+A for select all)
        if event.state & 0x4:  # Ctrl key is pressed
            if event.keysym in ['c', 'C', 'a', 'A']:
                return  # Allow these operations
        
        # Allow navigation keys
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Prior', 'Next']:
            return  # Allow arrow keys and navigation
            
        # Block all other key events (typing)
        return "break"

    def handle_code_key_event(self, event):
        """Handle key events for code display area - allow copy operations but prevent typing."""
        # Allow copy operations (Ctrl+C, Ctrl+A for select all)
        if event.state & 0x4:  # Ctrl key is pressed
            if event.keysym in ['c', 'C', 'a', 'A']:
                return  # Allow these operations
        
        # Allow navigation keys
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Prior', 'Next']:
            return  # Allow arrow keys and navigation
            
        # Block all other key events (typing)
        return "break"

    def handle_debug_output_key_event(self, event):
        """Handle key events for debug output area - allow copy operations but prevent typing."""
        # Allow copy operations (Ctrl+C, Ctrl+A for select all)
        if event.state & 0x4:  # Ctrl key is pressed
            if event.keysym in ['c', 'C', 'a', 'A']:
                return  # Allow these operations
        
        # Allow navigation keys
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Prior', 'Next']:
            return  # Allow arrow keys and navigation
            
        # Block all other key events (typing)
        return "break"

    def load_template(self, event=None):
        """Load a code template into the debug editor"""
        template = self.template_var.get()
        
        templates = {
            "Color Text Pattern Example": '''# Example: Make 'League of Legends' text green italic and 'Valorant' text red italic
# First, get the textbox ID from slide context (check the right panel)

# TEST: Simple color without italic first
modify_text_in_textbox(
    id=123,  # Replace with actual textbox ID from slide context
    find_pattern=r"\\bLeague of Legends\\b",  
    replacement_text="{color:#22c55e}League of Legends{/color}",  # Green color only
    regex_flags="IGNORECASE"
)

# TEST: Simple color without italic first
modify_text_in_textbox(
    id=123,  # Replace with actual textbox ID from slide context
    find_pattern=r"\\bValorant\\b",
    replacement_text="{color:#ef4444}Valorant{/color}",  # Red color only
    regex_flags="IGNORECASE"
)

# If above works, then try with italic:
# modify_text_in_textbox(
#     id=123,
#     find_pattern=r"\\bLeague of Legends\\b",
#     replacement_text="{color:#22c55e}*League of Legends*{/color}",  # Green + italic
#     regex_flags="IGNORECASE"
# )

print("Check if colors are applied correctly before adding italic formatting")''',
            
            "Debug HTML Test": '''# Debug: Test HTML parsing step by step

# Test 1: Simple replacement without any formatting
modify_text_in_textbox(
    id=123,  # Replace with actual textbox ID
    find_pattern=r"League of Legends",
    replacement_text="LEAGUE_REPLACED",  # No HTML, just plain text
    regex_flags="IGNORECASE"
)

# Test 2: Simple color formatting
modify_text_in_textbox(
    id=123,
    find_pattern=r"Valorant", 
    replacement_text="{color:#ff0000}Valorant{/color}",  # Simple red
    regex_flags="IGNORECASE"
)

# Test 3: Simple italic formatting  
modify_text_in_textbox(
    id=123,
    find_pattern=r"🤔",
    replacement_text="*thinking*",  # Simple italic
    regex_flags="IGNORECASE"
)

print("=== Debug Results ===")
print("Step 1: Did 'League of Legends' become 'LEAGUE_REPLACED'?")
print("Step 2: Did 'Valorant' turn red?") 
print("Step 3: Did emoji become italic 'thinking'?")''',
            
            "Replace Textbox Content Example": '''# Example: Completely replace all text in a textbox
replace_textbox_content(
    id=123,  # Replace with actual textbox ID from slide context
    markdown_text="<b>NEW TITLE: Battle of the Games</b> 🎮<br><br><ul><li><b>League of Legends:</b> <span style='color: purple'>Strategic MOBA</span></li><li><b>Valorant:</b> <span style='color: orange'>Tactical FPS</span></li></ul><br>Which will you choose?",
    font_size=16,
    text_align="center"
)''',
            
            "Modify Text in Textbox Example": '''# Example: Find and replace specific text while keeping everything else
# This preserves all existing text and only changes what you specify

# Change "deep strategy" to "complex strategy" with emphasis
modify_text_in_textbox(
    id=123,  # Replace with actual textbox ID
    find_pattern=r"deep strategy",
    replacement_text="**complex strategy**",
    regex_flags="IGNORECASE"
)

# Change "tactical shooting" to "precision shooting" with color
modify_text_in_textbox(
    id=123,
    find_pattern=r"tactical shooting",
    replacement_text="{color:red}**precision shooting**{/color}",
    regex_flags="IGNORECASE"
)

# Make any question text bold and blue
modify_text_in_textbox(
    id=123,
    find_pattern=r"Who will dominate\\? Choose your side!",
    replacement_text="{color:blue}**Who will dominate? Choose your side!**{/color}",
    regex_flags="IGNORECASE"
)''',
            
            "Add Text to Textbox Example": '''# Example: Add text to beginning or end of existing content

# Add text to the beginning
add_text_to_textbox(
    id=123,  # Replace with actual textbox ID
    markdown_text="🔥 <b>EPIC GAMING SHOWDOWN</b> 🔥<br><br>",
    position="start"
)

# Add text to the end
add_text_to_textbox(
    id=123,  # Replace with actual textbox ID
    markdown_text="<br><br><span style='color: gray'><i>Join the debate in the comments!</i></span>",
    position="end"
)''',
            
            "Format Textbox Style Example": '''# Example: Change visual formatting without modifying text content

# Change font and alignment
format_textbox_style(
    id=123,  # Replace with actual textbox ID
    font_name="Arial Black",  # Make it bold/dramatic
    font_size=18,  # Larger text
    text_align="center",
    line_spacing=1.5  # More space between lines
)

# Add margins for better spacing
format_textbox_style(
    id=123,
    left_margin=20,
    right_margin=20,
    top_margin=15,
    bottom_margin=15
)''',
            
            "Add New Textbox Example": '''# Example: Add a completely new textbox to the slide
add_textbox(
    slide_idx=1,  # Current slide
    markdown_text="<b>🎯 GAME STATS COMPARISON</b> 🎯<br><br><ul><li><b>League:</b> 150M+ monthly players</li><li><b>Valorant:</b> 15M+ monthly players</li></ul><br><span style='color: green'>League wins in numbers!</span>",
    left=500,  # Position on right side
    top=100,
    width=400,
    height=200,
    font_size=12,
    text_align="left"
)''',
            
            "Move and Resize Object Example": '''# Example: Reposition and resize objects on the slide

# Move an object to new position
move_object(
    id=123,  # Replace with actual object ID
    left=100,  # Move to left side
    top=50    # Move to top
)

# Resize an object
resize_object(
    id=123,  # Replace with actual object ID
    width=600,  # Make wider
    height=300  # Make taller
)

# Or do both at once
position_and_resize_object(
    id=123,  # Replace with actual object ID
    left=200,
    top=100,
    width=500,
    height=250
)''',
            
            "Get Object Properties Example": '''# Example: Inspect any object to see its properties
props = get_object_properties(id=123)  # Replace with actual object ID
print(f"Object properties: {props}")

# This will show:
# - Object type (TextBox, Picture, Shape, etc.)
# - Position (left, top)
# - Size (width, height)
# - Text content (if applicable)
# - Slide number
# - Object name and ID''',
            
            "Copy Object to Slide Example": '''# Example: Copy an object to another slide

# Copy textbox to slide 2 at same position
new_id = copy_object_to_slide(
    id=123,  # Replace with actual object ID
    target_slide_idx=2
)
print(f"Copied object to slide 2, new ID: {new_id}")

# Copy and position at specific coordinates
new_id = copy_object_to_slide(
    id=123,  # Replace with actual object ID
    target_slide_idx=3,
    new_left=300,
    new_top=200
)
print(f"Copied and positioned object on slide 3, new ID: {new_id}")''',
            
            "Duplicate Object Example": '''# Example: Duplicate an object on the same slide

# Duplicate with default offset (20 points right and down)
new_id = duplicate_object_on_same_slide(
    id=123  # Replace with actual object ID
)
print(f"Duplicated object, new ID: {new_id}")

# Duplicate with custom offset
new_id = duplicate_object_on_same_slide(
    id=123,  # Replace with actual object ID
    offset_left=50,  # Move 50 points right
    offset_top=100   # Move 100 points down
)
print(f"Duplicated with custom offset, new ID: {new_id}")

# Then you can modify the duplicate differently
modify_text_in_textbox(
    id=new_id,
    find_pattern=r"League of Legends",
    replacement_text="{color:gold}**LEAGUE OF LEGENDS**{/color}",
    regex_flags="IGNORECASE"
)'''
        }
        
        if template in templates:
            self.debug_editor.delete(1.0, tk.END)
            self.debug_editor.insert(1.0, templates[template])

    def clear_debug_editor(self):
        """Clear the debug editor"""
        self.debug_editor.delete(1.0, tk.END)
        self.debug_output.delete(1.0, tk.END)

    def refresh_slide_context(self):
        """Refresh the slide context display in the right panel"""
        try:
            self.ensure_ppt()
            if self.presentation is None:
                self.context_display.delete(1.0, tk.END)
                self.context_display.insert(tk.END, "❌ No presentation open.\n\nPlease create or open a presentation first.")
                return
            
            # Show loading indicator
            self.context_display.delete(1.0, tk.END)
            self.context_display.insert(tk.END, "🔄 Refreshing context...\n")
            self.root.update()  # Force UI update to show loading message
            
            # Get slide context using the agent's context reader with force refresh
            context = ppt_smolagent.get_current_slide_context(force_refresh=True)
            
            # Clear and update the display
            self.context_display.delete(1.0, tk.END)
            self.context_display.insert(tk.END, f"🔄 Updated: {datetime.datetime.now().strftime('%H:%M:%S')}\n")
            self.context_display.insert(tk.END, "="*40 + "\n\n")
            self.context_display.insert(tk.END, context)
            
            # Auto-scroll to top
            self.context_display.see("1.0")
            
        except Exception as e:
            self.context_display.delete(1.0, tk.END)
            self.context_display.insert(tk.END, f"❌ Error getting slide context:\n\n{str(e)}\n\n")
            self.context_display.insert(tk.END, "💡 Troubleshooting tips:\n")
            self.context_display.insert(tk.END, "• Make sure PowerPoint is open\n")
            self.context_display.insert(tk.END, "• Ensure a presentation is loaded\n")
            self.context_display.insert(tk.END, "• Try clicking on a slide in PowerPoint first\n")
            self.context_display.insert(tk.END, "• Check if there are any unsaved changes")

    def refresh_slide_context_with_feedback(self):
        """Refresh slide context with visual feedback on the button"""
        # Change button text to show loading state
        original_text = self.refresh_context_btn.cget("text")
        original_bg = self.refresh_context_btn.cget("bg")
        
        self.refresh_context_btn.config(text="⏳ Refreshing...", bg="#f59e0b")  # Orange color for loading
        self.refresh_context_btn.update()
        
        try:
            # Call the actual refresh method
            self.refresh_slide_context()
            
            # Show success state briefly
            self.refresh_context_btn.config(text="✅ Refreshed!", bg="#10b981")  # Green for success
            self.root.after(1000, lambda: self.refresh_context_btn.config(text=original_text, bg=original_bg))
            
        except Exception as e:
            # Show error state briefly
            self.refresh_context_btn.config(text="❌ Error", bg="#ef4444")  # Red for error
            self.root.after(2000, lambda: self.refresh_context_btn.config(text=original_text, bg=original_bg))

    def on_tab_changed(self, event):
        """Handle tab change events"""
        selected_tab = self.notebook.select()
        tab_text = self.notebook.tab(selected_tab, "text")
        
        # Auto-refresh context when switching to debug tab
        if "Debug Console" in tab_text:
            # Small delay to ensure tab is fully loaded
            self.root.after(100, self.refresh_slide_context)

    def get_slide_context(self):
        """Get current slide context and display in output (legacy method for backward compatibility)"""
        try:
            self.ensure_ppt()
            if self.presentation is None:
                self.debug_output.delete(1.0, tk.END)
                self.debug_output.insert(tk.END, "❌ No presentation open. Please create or open a presentation first.\n")
                return
            
            # Get slide context using the agent's context reader
            context = ppt_smolagent.get_current_slide_context()
            
            self.debug_output.delete(1.0, tk.END)
            self.debug_output.insert(tk.END, "📋 Current Slide Context:\n")
            self.debug_output.insert(tk.END, "="*50 + "\n")
            self.debug_output.insert(tk.END, context)
            self.debug_output.insert(tk.END, "\n" + "="*50 + "\n")
            
            # Also update the context panel
            self.refresh_slide_context()
            
        except Exception as e:
            self.debug_output.delete(1.0, tk.END)
            self.debug_output.insert(tk.END, f"❌ Error getting slide context: {str(e)}\n")

    def execute_debug_code(self):
        """Execute the code in the debug editor"""
        code = self.debug_editor.get(1.0, tk.END).strip()
        
        if not code:
            self.debug_output.delete(1.0, tk.END)
            self.debug_output.insert(tk.END, "⚠️ No code to execute. Please enter some code first.\n")
            return
        
        try:
            self.ensure_ppt()
            if self.presentation is None:
                self.debug_output.delete(1.0, tk.END)
                self.debug_output.insert(tk.END, "❌ No presentation open. Please create or open a presentation first.\n")
                return
            
            self.debug_output.delete(1.0, tk.END)
            self.debug_output.insert(tk.END, "🚀 Executing code...\n")
            self.debug_output.insert(tk.END, "="*50 + "\n")
            self.root.update()  # Force UI update
            
            # Import the tools into the execution namespace
            import ppt_smolagent
            
            # Create execution namespace with all the tools
            exec_namespace = {
                # New improved tools
                'add_textbox': ppt_smolagent.add_textbox,
                'replace_textbox_content': ppt_smolagent.replace_textbox_content,
                'modify_text_in_textbox': ppt_smolagent.modify_text_in_textbox,
                'add_text_to_textbox': ppt_smolagent.add_text_to_textbox,
                'format_textbox_style': ppt_smolagent.format_textbox_style,
                'move_object': ppt_smolagent.move_object,
                'resize_object': ppt_smolagent.resize_object,
                'position_and_resize_object': ppt_smolagent.position_and_resize_object,
                'get_object_properties': ppt_smolagent.get_object_properties,
                'copy_object_to_slide': ppt_smolagent.copy_object_to_slide,
                'duplicate_object_on_same_slide': ppt_smolagent.duplicate_object_on_same_slide,
                'delete_object': ppt_smolagent.delete_object,
                # Legacy tools for backward compatibility (if they still exist)
                'update_textbox': getattr(ppt_smolagent, 'update_textbox', None),
                'format_text_pattern': getattr(ppt_smolagent, 'format_text_pattern', None),
                'duplicate_object': getattr(ppt_smolagent, 'duplicate_object', None),
                # Utility functions
                'print': lambda *args: self.debug_print(*args)
            }
            
            # Execute the code
            try:
                exec(code, exec_namespace)
                self.debug_output.insert(tk.END, "\n✅ Code executed successfully!\n")
            except Exception as e:
                self.debug_output.insert(tk.END, f"\n❌ Execution error: {str(e)}\n")
                import traceback
                self.debug_output.insert(tk.END, f"Traceback:\n{traceback.format_exc()}\n")
                
        except Exception as e:
            self.debug_output.insert(tk.END, f"❌ Setup error: {str(e)}\n")
    
    def debug_print(self, *args):
        """Custom print function for debug console"""
        message = " ".join(str(arg) for arg in args)
        self.debug_output.insert(tk.END, message + "\n")
        self.debug_output.see(tk.END)

    def handle_code_key_event_debug(self, event):
        """Handle key events for code display area - allow copy operations but prevent typing."""
        # Allow copy operations (Ctrl+C, Ctrl+A for select all)
        if event.state & 0x4:  # Ctrl key is pressed
            if event.keysym in ['c', 'C', 'a', 'A']:
                return  # Allow these operations
        
        # Allow navigation keys
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Prior', 'Next']:
            return  # Allow arrow keys and navigation
            
        # Block all other key events (typing)
        return "break"

    def strip_ansi_codes(self, text):
        """Remove ANSI color codes and formatting from text."""
        import re
        # Pattern to match ANSI escape codes
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        # Also remove common color codes that might appear
        color_codes = re.compile(r'\[[0-9;]*m')
        
        # Remove ANSI codes
        text = ansi_escape.sub('', text)
        text = color_codes.sub('', text)
        
        # Clean up extra whitespace and newlines
        lines = text.split('\n')
        cleaned_lines = []
        for line in lines:
            line = line.strip()
            # Skip lines that are just formatting artifacts
            if line and not line.startswith('[') and 'Duration' not in line and 'tokens:' not in line:
                cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)

    def update_code_display(self, code_text):
        """Update the code display area with enhanced syntax highlighting."""
        self.code_display.delete(1.0, tk.END)
        
        # Strip ANSI codes from the code text
        cleaned_code = self.strip_ansi_codes(code_text)
        
        # Add modern header with better formatting
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        header = f"# 🔧 Generated Code ({timestamp})\n# Click '▼ Generated Code' to collapse\n\n"
        
        self.code_display.insert(tk.END, header)
        self.code_display.insert(tk.END, cleaned_code)
        
        # Enhanced syntax highlighting with modern colors
        self.code_display.tag_configure("keyword", foreground="#569cd6", font=("JetBrains Mono", 10, "bold") if self.is_font_available("JetBrains Mono") else ("Consolas", 10, "bold"))
        self.code_display.tag_configure("string", foreground="#ce9178")
        self.code_display.tag_configure("comment", foreground="#6a9955", font=("JetBrains Mono", 10, "italic") if self.is_font_available("JetBrains Mono") else ("Consolas", 10, "italic"))
        self.code_display.tag_configure("function", foreground="#dcdcaa")
        self.code_display.tag_configure("number", foreground="#b5cea8")
        
        # Apply enhanced syntax highlighting
        content = self.code_display.get(1.0, tk.END)
        lines = content.split('\n')
        
        for i, line in enumerate(lines):
            line_start = f"{i+1}.0"
            line_end = f"{i+1}.end"
            
            # Highlight Python keywords
            keywords = ['def', 'import', 'from', 'if', 'else', 'elif', 'for', 'while', 'try', 'except', 'with', 'as', 'return', 'class', 'True', 'False', 'None']
            for keyword in keywords:
                import re
                pattern = r'\b' + re.escape(keyword) + r'\b'
                for match in re.finditer(pattern, line):
                    start_idx = f"{i+1}.{match.start()}"
                    end_idx = f"{i+1}.{match.end()}"
                    self.code_display.tag_add("keyword", start_idx, end_idx)
            
            # Highlight strings
            string_patterns = [r'"[^"]*"', r"'[^']*'"]
            for pattern in string_patterns:
                for match in re.finditer(pattern, line):
                    start_idx = f"{i+1}.{match.start()}"
                    end_idx = f"{i+1}.{match.end()}"
                    self.code_display.tag_add("string", start_idx, end_idx)
            
            # Highlight comments
            if '#' in line:
                comment_start = line.find('#')
                if comment_start >= 0:
                    start_idx = f"{i+1}.{comment_start}"
                    self.code_display.tag_add("comment", start_idx, line_end)
            
            # Highlight numbers
            number_pattern = r'\b\d+\.?\d*\b'
            for match in re.finditer(number_pattern, line):
                start_idx = f"{i+1}.{match.start()}"
                end_idx = f"{i+1}.{match.end()}"
                self.code_display.tag_add("number", start_idx, end_idx)

    def log(self, message):
        # Modern chat bubbles with better spacing and typography
        if message.startswith("[You]"):
            msg_text = message.replace("[You] ", "")
            tag = "user_msg"
            # Clean modern bubble without extra symbols
            bubble = f"  {msg_text}  "
            self.chat_area.insert(tk.END, "\n", tag) if self.chat_area.index(tk.END) != "1.0" else None
            self.chat_area.insert(tk.END, f"{bubble}\n", tag)
        elif message.startswith("[System]"):
            msg_text = message.replace("[System] ", "")
            tag = "sys_msg"
            # Clean modern bubble without extra symbols
            bubble = f"  {msg_text}  "
            self.chat_area.insert(tk.END, "\n", tag) if self.chat_area.index(tk.END) != "1.0" else None
            self.chat_area.insert(tk.END, f"{bubble}\n", tag)
        else:
            self.chat_area.insert(tk.END, message + '\n')
        
        # Modern chat bubble styling
        self.chat_area.tag_configure(
            "user_msg",
            justify='right',
            background=self.user_msg_bg,
            foreground=self.user_msg_fg,
            lmargin1=150, lmargin2=150, rmargin=20,
            spacing1=8, spacing3=8,
            font=("Segoe UI", 11),
            relief="flat",
            borderwidth=0
        )
        self.chat_area.tag_configure(
            "sys_msg",
            justify='left',
            background=self.sys_msg_bg,
            foreground=self.sys_msg_fg,
            lmargin1=20, lmargin2=20, rmargin=150,
            spacing1=8, spacing3=8,
            font=("Segoe UI", 11),
            relief="flat",
            borderwidth=0
        )
        self.chat_area.see(tk.END)

    def ensure_ppt(self):
        pythoncom.CoInitialize()
        if self.ppt_app is None:
            try:
                self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                self.ppt_app.Visible = True
        # Always try to use the active presentation if available
        if self.ppt_app.Presentations.Count > 0:
            self.presentation = self.ppt_app.ActivePresentation
        
        # NEW: Ensure a slide is selected (fallback to last slide)
        self.select_default_slide()

    def select_default_slide(self):
        """Ensure there is always an active slide selection.
        If a new presentation has no slides, create a blank one and select it.
        If slides exist but none is selected, select the last slide by default.
        """
        if not self.presentation:
            return

        # If the presentation has no slides, add a blank slide and select it
        if self.presentation.Slides.Count == 0:
            # 12 corresponds to ppLayoutBlank
            self.presentation.Slides.Add(1, 12)
            try:
                self.ppt_app.ActiveWindow.View.GotoSlide(1)
            except Exception:
                pass
            return

        # Check if a slide is currently selected/active
        try:
            _ = self.ppt_app.ActiveWindow.View.Slide.SlideIndex  # Accessing raises if no slide
            return  # Slide is already selected, no further action needed
        except Exception:
            pass  # No active slide, fall through to select last slide

        # Select the last slide as a sensible default
        last_idx = self.presentation.Slides.Count
        try:
            self.ppt_app.ActiveWindow.View.GotoSlide(last_idx)
        except Exception:
            # In some views (e.g., slide sorter), GotoSlide may not work. Fallback to selection API.
            try:
                slide_range = self.presentation.Slides(last_idx)
                slide_range.Select()
            except Exception:
                pass

    def create_new_ppt(self):
        self.ensure_ppt()
        self.presentation = self.ppt_app.Presentations.Add()
        # Ensure the new presentation has a first slide selected
        self.select_default_slide()
        self.log("[System] New PowerPoint presentation created.")

    def open_ppt(self):
        self.ensure_ppt()
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx;*.ppt")])
        if file_path:
            self.presentation = self.ppt_app.Presentations.Open(file_path)
            # Ensure a slide is selected in the newly opened presentation
            self.select_default_slide()
            self.log(f"[System] Opened presentation: {file_path}")

    def send_message(self):
        msg = self.entry.get().strip()
        # Don't send if it's empty or just placeholder text
        if not msg or msg == self.placeholder_text:
            return
        self.log(f"[You] {msg}")
        self.entry.delete(0, tk.END)
        # Reset placeholder
        self.set_entry_placeholder()
        # Always try to attach to the active presentation before handling command
        self.ensure_ppt()
        self.handle_command(msg)

    def handle_command(self, msg):
        if self.presentation is None:
            self.log("[System] 📋 Please create or open a presentation first.")
            return
        # Use the smolagent to process the message and execute the tool
        try:
            # Show modern processing message
            self.log("[System] 🔄 Processing your request...")
            self.root.update()  # Force UI update
            
            result = ppt_smolagent.run_agent_with_code_capture(msg)
            
            # Display the final answer with emoji
            self.log(f"[System] ✅ {result['answer']}")
            
            # Update the code display
            self.update_code_display(result['generated_code'])
            
            # Show code toggle button with modern indicator
            if result['generated_code'] and "Error occurred" not in result['generated_code']:
                self.code_toggle_btn.config(text="▶ Generated Code �")  # Green indicator for new code
                # If code area is currently visible, update the button text accordingly
                if self.code_display_visible:
                    self.code_toggle_btn.config(text="▼ Generated Code")
            else:
                self.code_toggle_btn.config(text="▶ Generated Code")
            
        except Exception as e:
            self.log(f"[System] ❌ Error: {e}")
            self.update_code_display(f"# ❌ Error occurred during execution:\n# {str(e)}\n\n# 🔍 Please check:\n# - PowerPoint is running\n# - Valid presentation is open\n# - Network connection for AI model")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTAssistant(root)
    root.mainloop()
