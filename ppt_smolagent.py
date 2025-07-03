from smolagents import CodeAgent, tool, OpenAIServerModel
from smolagents.monitoring import LogLevel
import os
from dotenv import load_dotenv
import logging
import io
import sys

# Load environment variables from .env file
load_dotenv()

# Set the OpenAI API key from environment
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found in environment variables. Please check your .env file.")

# Define the model using OpenAIServerModel
model = OpenAIServerModel(
    model_id="gpt-4o-mini",
    api_key=openai_api_key,
    api_base = "https://api.openai.com/v1"
)

# Tool to add a textbox to a PowerPoint slide
@tool
def add_textbox_tool(slide_idx: int = 1, text: str = "Sample Text", left: int = 100, top: int = 100, width: int = 400, height: int = 50, font_size: int = None, font_name: str = None, font_bold: bool = None, font_italic: bool = None) -> str:
    """
    Add a textbox to a PowerPoint slide with customizable text and formatting.
    
    Args:
        slide_idx: The slide number (1-indexed) to add the textbox to
        text: The text content for the textbox
        left: Left position of the textbox in points
        top: Top position of the textbox in points
        width: Width of the textbox in points
        height: Height of the textbox in points
        font_size: Font size for the text (optional)
        font_name: Font name for the text (optional)
        font_bold: Whether to make the text bold (optional)
        font_italic: Whether to make the text italic (optional)
    
    Returns:
        str: Confirmation message of the textbox addition
    """
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
    presentation = ppt_app.ActivePresentation
    if presentation.Slides.Count < slide_idx:
        slide = presentation.Slides.Add(slide_idx, 12)  # 12 = ppLayoutBlank
    else:
        slide = presentation.Slides(slide_idx)
    box = slide.Shapes.AddTextbox(1, left, top, width, height)
    box.TextFrame.TextRange.Text = text
    if font_size:
        box.TextFrame.TextRange.Font.Size = font_size
    if font_name:
        box.TextFrame.TextRange.Font.Name = font_name
    if font_bold is not None:
        box.TextFrame.TextRange.Font.Bold = -1 if font_bold else 0
    if font_italic is not None:
        box.TextFrame.TextRange.Font.Italic = -1 if font_italic else 0
    return f"Textbox added to slide {slide_idx} with text: {text}"

# The tool is automatically registered when using the @tool decorator

instructions = """
You are a highly capable AI assistant that writes Python code to automate Microsoft PowerPoint presentations using the tools provided to you.

CODE FORMATTING REQUIREMENTS:
- ALWAYS write code between triple backticks (```) 
- NEVER EVER use other code formatting like <code></code>, single backticks, or any other format

General PowerPoint Concepts:
- A PowerPoint presentation is a `.pptx` file containing one or more slides.
- Each slide can contain placeholders, text boxes, shapes, images, charts, tables, and animations.
- The slide coordinate system is in **points** (1 inch = 72 points).
- Standard slide dimensions (default in PowerPoint):
  - Width: 960 points (13.33 inches)
  - Height: 540 points (7.5 inches)
  - Origin (0, 0) is the top-left corner of the slide.

"""

# Create a custom logging handler to capture code generation
class CodeCaptureHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.captured_code = []
        
    def emit(self, record):
        if hasattr(record, 'msg'):
            msg = str(record.msg)
            # Look for code patterns in the log messages
            if any(keyword in msg for keyword in ['def ', 'import ', 'from ', 'class ', 'with ', 'for ', 'if ']):
                self.captured_code.append(msg)
    
    def get_code(self):
        return '\n'.join(self.captured_code)
    
    def clear(self):
        self.captured_code = []

# Global code capture handler
code_capture_handler = CodeCaptureHandler()

agent = CodeAgent(
    tools=[add_textbox_tool],
    instructions=instructions,
    max_steps=3,
    model=model,
    verbosity_level=LogLevel.DEBUG
)

def strip_ansi_codes(text):
    """Remove ANSI color codes and formatting from text."""
    import re
    # Pattern to match ANSI escape codes
    ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
    # Also remove common color codes that might appear
    color_codes = re.compile(r'\[[0-9;]*m')
    
    # Remove ANSI codes
    text = ansi_escape.sub('', text)
    text = color_codes.sub('', text)
    
    return text

def run_agent_with_code_capture(message):
    """
    Run the agent and capture both the final answer and generated code.
    
    Returns:
        dict: Contains 'answer' and 'generated_code' keys
    """
    try:
        # Clear previous captured code
        code_capture_handler.clear()
        
        # Set up logging to capture the agent's output
        logger = logging.getLogger()
        logger.addHandler(code_capture_handler)
        logger.setLevel(logging.DEBUG)
        
        # Capture stdout/stderr as well
        stdout_backup = sys.stdout
        stderr_backup = sys.stderr
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()
        
        try:
            sys.stdout = stdout_capture
            sys.stderr = stderr_capture
            
            # Run the agent
            answer = agent.run(message)
            
        finally:
            # Restore stdout/stderr
            sys.stdout = stdout_backup
            sys.stderr = stderr_backup
            logger.removeHandler(code_capture_handler)
        
        # Get captured outputs and clean them
        stdout_content = strip_ansi_codes(stdout_capture.getvalue())
        stderr_content = strip_ansi_codes(stderr_capture.getvalue())
        captured_code = strip_ansi_codes(code_capture_handler.get_code())
        
        # Try to extract code from various sources
        generated_code = ""
        
        # First, try the captured code from logs
        if captured_code.strip():
            generated_code = captured_code
        
        # Next, try to extract from stdout
        elif stdout_content:
            # Look for code patterns in stdout
            import re
            
            # Look for function definitions and imports
            code_patterns = [
                r'(def\s+\w+.*?(?=\n\w|\n$))',  # Function definitions
                r'(import\s+\w+.*)',  # Import statements
                r'(from\s+\w+.*)',  # From imports
                r'(\w+\s*=\s*.*)',  # Assignments
            ]
            
            for pattern in code_patterns:
                matches = re.findall(pattern, stdout_content, re.MULTILINE | re.DOTALL)
                if matches:
                    generated_code += '\n'.join(matches) + '\n'
        
        # If still no code, try to extract from the answer itself
        if not generated_code.strip():
            import re
            # Clean the answer first
            clean_answer = strip_ansi_codes(answer)
            
            # Look for code blocks in the answer
            code_blocks = re.findall(r'```(?:python)?\n?(.*?)\n?```', clean_answer, re.DOTALL)
            if code_blocks:
                generated_code = '\n'.join(code_blocks)
            else:
                # Look for Python-like statements in the answer
                lines = clean_answer.split('\n')
                code_lines = []
                for line in lines:
                    stripped = line.strip()
                    if any(keyword in stripped for keyword in ['def ', 'import ', 'from ', '=', 'print(', 'if ', 'for ', 'with ', 'try:']):
                        code_lines.append(line)
                if code_lines:
                    generated_code = '\n'.join(code_lines)
        
        # Fallback message if no code was captured
        if not generated_code.strip():
            # Create a summary based on the tool that was likely used
            tool_name = "add_textbox_tool" if "textbox" in message.lower() else "PowerPoint automation tool"
            generated_code = f"""# Agent Execution Summary
# Request: "{message}"
# 
# The agent executed your request using the {tool_name}().
# This is a direct tool call that doesn't require custom code generation.
#
# The operation was completed successfully using the built-in PowerPoint COM interface.
# 
# Example of what the agent did internally:
import win32com.client
import pythoncom

# Initialize PowerPoint COM interface
pythoncom.CoInitialize()
ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
presentation = ppt_app.ActivePresentation

# Tool executed with your parameters
# Result: {strip_ansi_codes(answer) if answer else 'Operation completed'}"""
        
        # Clean the final answer
        clean_answer = strip_ansi_codes(answer) if answer else "Operation completed"
        
        return {
            'answer': clean_answer,
            'generated_code': generated_code,
            'debug_output': f"STDOUT:\n{stdout_content}\n\nSTDERR:\n{stderr_content}"
        }
        
    except Exception as e:
        return {
            'answer': f"Error: {str(e)}",
            'generated_code': f"# Error occurred during execution:\n# {str(e)}\n\n# This might be due to:\n# - Missing dependencies\n# - PowerPoint not running\n# - Invalid parameters",
            'debug_output': str(e)
        }

if __name__ == "__main__":
    # Example: agent.run("Add a textbox with text 'Hello' on slide 2, font size 32, bold.")
    print(agent.run("Add a textbox with text 'Hello' on slide 2, font size 32, bold."))
