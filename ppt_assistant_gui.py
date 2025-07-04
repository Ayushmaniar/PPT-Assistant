import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

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
        self.root.geometry("900x700")
        self.root.minsize(700, 500)

        # Main container with padding and modern styling
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header section
        header_frame = tk.Frame(main_frame, bg=self.card_bg, relief="flat")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Add subtle border effect
        header_border = tk.Frame(header_frame, bg=self.accent_color, height=3)
        header_border.pack(fill=tk.X, side=tk.TOP)
        
        title_label = tk.Label(header_frame, text="💼 PPT Assistant", 
                              bg=self.card_bg, fg=self.accent_color, 
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
        input_container = tk.Frame(self.root, bg=self.card_bg, relief="flat")
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
