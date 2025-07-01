import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import win32com.client
import pythoncom

class PPTAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT Assistant Chat UI")
        self.ppt_app = None
        self.presentation = None
        self.setup_ui()

    def setup_ui(self):
        self.chat_area = scrolledtext.ScrolledText(self.root, state='disabled', width=60, height=20)
        self.chat_area.pack(padx=10, pady=10)

        self.entry = tk.Entry(self.root, width=50)
        self.entry.pack(side=tk.LEFT, padx=(10,0), pady=(0,10))
        self.entry.bind('<Return>', lambda event: self.send_message())

        send_btn = tk.Button(self.root, text="Send", command=self.send_message)
        send_btn.pack(side=tk.LEFT, padx=(5,0), pady=(0,10))

        new_btn = tk.Button(self.root, text="New PPT", command=self.create_new_ppt)
        new_btn.pack(side=tk.LEFT, padx=(5,0), pady=(0,10))

        open_btn = tk.Button(self.root, text="Open PPT", command=self.open_ppt)
        open_btn.pack(side=tk.LEFT, padx=(5,10), pady=(0,10))

    def log(self, message):
        self.chat_area.config(state='normal')
        self.chat_area.insert(tk.END, message + '\n')
        self.chat_area.config(state='disabled')
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

    def create_new_ppt(self):
        self.ensure_ppt()
        self.presentation = self.ppt_app.Presentations.Add()
        self.log("[System] New PowerPoint presentation created.")

    def open_ppt(self):
        self.ensure_ppt()
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx;*.ppt")])
        if file_path:
            self.presentation = self.ppt_app.Presentations.Open(file_path)
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
        # Basic command parsing for demo
        if msg.lower().startswith("add textbox"):
            # Add a slide if none exist
            if self.presentation.Slides.Count == 0:
                slide = self.presentation.Slides.Add(1, 12)  # 12 = ppLayoutBlank
            else:
                slide = self.presentation.Slides(1)
            box = slide.Shapes.AddTextbox(1, 100, 100, 400, 50)
            box.TextFrame.TextRange.Text = "Sample Text"
            self.log("[System] Added a textbox to slide 1.")
        else:
            self.log("[System] Command not recognized. Try 'add textbox'.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTAssistant(root)
    root.mainloop()
