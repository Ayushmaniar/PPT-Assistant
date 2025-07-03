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
        # Set dark mode colors
        self.bg_color = "#1e1e1e"
        self.user_msg_bg = "#075e54"
        self.user_msg_fg = "#e0f2f1"
        self.sys_msg_bg = "#262d31"
        self.sys_msg_fg = "#ece5dd"
        self.entry_bg = "#2a2f32"
        self.entry_fg = "#ece5dd"
        self.btn_bg = "#25d366"
        self.btn_fg = "#1e1e1e"
        self.code_bg = "#1a1a1a"
        self.code_fg = "#f0f0f0"

        self.root.configure(bg=self.bg_color)

        # Main container
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Chat area
        self.chat_area = scrolledtext.ScrolledText(main_frame, state='normal', width=60, height=15, bg=self.bg_color, fg=self.sys_msg_fg, font=("Segoe UI", 11), bd=0, highlightthickness=0, wrap=tk.WORD)
        self.chat_area.pack(fill=tk.BOTH, expand=True)
        
        # Make chat area read-only but selectable with copy functionality
        self.chat_area.bind("<Key>", self.handle_chat_key_event)  # Handle key events
        self.chat_area.bind("<Button-1>", lambda e: self.chat_area.focus_set())  # Allow focus for selection

        # Code display section (initially hidden)
        self.code_frame = tk.Frame(main_frame, bg=self.bg_color)
        self.code_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Code toggle button
        self.code_toggle_btn = tk.Button(self.code_frame, text="▶ Show Generated Code", command=self.toggle_code_display, bg=self.btn_bg, fg=self.btn_fg, font=("Segoe UI", 9), bd=0, activebackground="#128c7e", activeforeground=self.btn_fg, cursor="hand2")
        self.code_toggle_btn.pack(anchor='w')
        
        # Code display area (hidden by default)
        self.code_display = scrolledtext.ScrolledText(self.code_frame, state='normal', width=60, height=8, bg=self.code_bg, fg=self.code_fg, font=("Consolas", 10), bd=1, highlightthickness=1, wrap=tk.WORD)
        self.code_display_visible = False
        
        # Make code display read-only but selectable with copy functionality
        self.code_display.bind("<Key>", self.handle_code_key_event)  # Handle key events
        self.code_display.bind("<Button-1>", lambda e: self.code_display.focus_set())  # Allow focus for selection

        # Entry frame
        entry_frame = tk.Frame(self.root, bg=self.bg_color)
        entry_frame.pack(fill=tk.X, padx=10, pady=(0,10))

        self.entry = tk.Entry(entry_frame, width=50, bg=self.entry_bg, fg=self.entry_fg, insertbackground=self.entry_fg, font=("Segoe UI", 11), bd=0, highlightthickness=0)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        self.entry.bind('<Return>', lambda event: self.send_message())

        send_btn = tk.Button(entry_frame, text="Send", command=self.send_message, bg=self.btn_bg, fg=self.btn_fg, font=("Segoe UI", 10, "bold"), bd=0, activebackground="#128c7e", activeforeground=self.btn_fg, cursor="hand2")
        send_btn.pack(side=tk.LEFT, padx=(0,5))

        new_btn = tk.Button(entry_frame, text="New PPT", command=self.create_new_ppt, bg=self.btn_bg, fg=self.btn_fg, font=("Segoe UI", 10), bd=0, activebackground="#128c7e", activeforeground=self.btn_fg, cursor="hand2")
        new_btn.pack(side=tk.LEFT, padx=(0,5))

        open_btn = tk.Button(entry_frame, text="Open PPT", command=self.open_ppt, bg=self.btn_bg, fg=self.btn_fg, font=("Segoe UI", 10), bd=0, activebackground="#128c7e", activeforeground=self.btn_fg, cursor="hand2")
        open_btn.pack(side=tk.LEFT)

    def toggle_code_display(self):
        """Toggle the visibility of the code display area."""
        if self.code_display_visible:
            # Hide code display
            self.code_display.pack_forget()
            self.code_toggle_btn.config(text="▶ Show Generated Code")
            self.code_display_visible = False
        else:
            # Show code display
            self.code_display.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
            self.code_toggle_btn.config(text="▼ Hide Generated Code")
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
        """Update the code display area with new code."""
        # Code display is already in 'normal' state but read-only via key binding
        self.code_display.delete(1.0, tk.END)
        
        # Strip ANSI codes from the code text
        cleaned_code = self.strip_ansi_codes(code_text)
        
        # Add header with timestamp
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        header = f"# Generated Code - {timestamp}\n# Click 'Hide Generated Code' to collapse this view\n\n"
        
        self.code_display.insert(tk.END, header)
        self.code_display.insert(tk.END, cleaned_code)
        
        # Simple syntax highlighting for Python keywords
        self.code_display.tag_configure("keyword", foreground="#569cd6")  # Blue for keywords
        self.code_display.tag_configure("string", foreground="#ce9178")   # Orange for strings
        self.code_display.tag_configure("comment", foreground="#6a9955")  # Green for comments
        
        # Apply basic syntax highlighting
        content = self.code_display.get(1.0, tk.END)
        lines = content.split('\n')
        
        for i, line in enumerate(lines):
            line_start = f"{i+1}.0"
            line_end = f"{i+1}.end"
            
            # Highlight Python keywords
            keywords = ['def', 'import', 'from', 'if', 'else', 'for', 'while', 'try', 'except', 'with', 'as', 'return', 'class']
            for keyword in keywords:
                if f' {keyword} ' in line or line.startswith(keyword + ' '):
                    # Simple highlighting - this is basic but functional
                    pass
            
            # Highlight comments
            if '#' in line:
                comment_start = line.find('#')
                if comment_start >= 0:
                    start_idx = f"{i+1}.{comment_start}"
                    self.code_display.tag_add("comment", start_idx, line_end)

    def log(self, message):
        # Chat area is already in 'normal' state but read-only via key binding
        # WhatsApp-like chat bubbles using Unicode and padding
        if message.startswith("[You]"):
            msg_text = message.replace("[You] ", "")
            tag = "user_msg"
            # Add padding and a right-pointing tail
            bubble = f"\u2003\u2003{msg_text}   \u25B6"
            self.chat_area.insert(tk.END, "\n", tag) if self.chat_area.index(tk.END) != "1.0" else None
            self.chat_area.insert(tk.END, f"{bubble}\n", tag)
        elif message.startswith("[System]"):
            msg_text = message.replace("[System] ", "")
            tag = "sys_msg"
            # Add padding and a left-pointing tail
            bubble = f"\u25C0   {msg_text}\u2003\u2003"
            self.chat_area.insert(tk.END, "\n", tag) if self.chat_area.index(tk.END) != "1.0" else None
            self.chat_area.insert(tk.END, f"{bubble}\n", tag)
        else:
            self.chat_area.insert(tk.END, message + '\n')
        # Tag configs for chat bubbles
        self.chat_area.tag_configure(
            "user_msg",
            justify='right',
            background=self.user_msg_bg,
            foreground=self.user_msg_fg,
            lmargin1=200, lmargin2=200, rmargin=10,
            spacing1=5, spacing3=5,
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            borderwidth=10
        )
        self.chat_area.tag_configure(
            "sys_msg",
            justify='left',
            background=self.sys_msg_bg,
            foreground=self.sys_msg_fg,
            lmargin1=10, lmargin2=10, rmargin=200,
            spacing1=5, spacing3=5,
            font=("Segoe UI", 11),
            relief="flat",
            borderwidth=10
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
        if not msg:
            return
        self.log(f"[You] {msg}")
        self.entry.delete(0, tk.END)
        # Always try to attach to the active presentation before handling command
        self.ensure_ppt()
        self.handle_command(msg)

    def handle_command(self, msg):
        if self.presentation is None:
            self.log("[System] Please create or open a presentation first.")
            return
        # Use the smolagent to process the message and execute the tool
        try:
            # Show processing message
            self.log("[System] Processing your request...")
            self.root.update()  # Force UI update
            
            result = ppt_smolagent.run_agent_with_code_capture(msg)
            
            # Display the final answer
            self.log(f"[System] {result['answer']}")
            
            # Update the code display
            self.update_code_display(result['generated_code'])
            
            # Show code toggle button with indicator if code was generated
            if result['generated_code'] and "Error occurred" not in result['generated_code']:
                self.code_toggle_btn.config(text="▶ Show Generated Code 🔴")  # Red dot indicates new code
                # If code area is currently visible, update the button text accordingly
                if self.code_display_visible:
                    self.code_toggle_btn.config(text="▼ Hide Generated Code")
            else:
                self.code_toggle_btn.config(text="▶ Show Generated Code")
            
        except Exception as e:
            self.log(f"[System] Error: {e}")
            self.update_code_display(f"# Error occurred during execution:\n# {str(e)}\n\n# Please check:\n# - PowerPoint is running\n# - Valid presentation is open\n# - Network connection for AI model")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTAssistant(root)
    root.mainloop()
