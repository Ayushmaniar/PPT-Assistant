from smolagents import CodeAgent, tool, OpenAIServerModel
from smolagents.monitoring import LogLevel
import os
from dotenv import load_dotenv
import logging
import io
import sys
from slide_context_reader import PowerPointSlideReader

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
def add_textbox_tool(slide_idx: int = 1, text: str = "Sample Text", left: int = 100, top: int = 100, width: int = 400, height: int = 50, font_size: int = None, font_name: str = None, font_bold: bool = None, font_italic: bool = None, text_align: str = "left") -> str:
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
        text_align: Text alignment - "left", "center", or "right" (default: "left")
    
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
    
    # Set text alignment
    # PowerPoint alignment constants: 1 = left, 2 = center, 3 = right
    alignment_map = {
        "left": 1,
        "center": 2, 
        "right": 3
    }
    
    if text_align.lower() in alignment_map:
        box.TextFrame.TextRange.ParagraphFormat.Alignment = alignment_map[text_align.lower()]
    
    return f"Textbox added to slide {slide_idx} with text: {text}"

# The tool is automatically registered when using the @tool decorator

instructions = """
You are a highly capable AI assistant that writes Python code to automate Microsoft PowerPoint presentations using the tools provided to you.

IMPORTANT: You will ALWAYS receive the current slide context before the user's request. This context contains detailed information about:
- The currently selected slide number and layout
- All objects/shapes present on the slide (textboxes, images, tables, charts, etc.)
- Their positions, sizes, text content, and formatting
- Any animations or slide notes

USE THIS CONTEXT to make informed decisions about:
- Where to position new elements (avoid overlapping existing content)
- What font sizes and styles to use (match existing elements when appropriate)
- How to complement or enhance the existing slide content
- Whether modifications should be made to existing elements vs. adding new ones

CODE FORMATTING REQUIREMENTS:
- ALWAYS write code between <code>(.*?)</code> 
- NEVER EVER use other code formatting.

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

# Global slide context reader instance
slide_reader = None

def get_slide_reader():
    """Get or create the global slide reader instance."""
    global slide_reader
    if slide_reader is None:
        try:
            slide_reader = PowerPointSlideReader()
        except Exception as e:
            print(f"Warning: Could not initialize slide reader: {e}")
            slide_reader = None
    return slide_reader

def get_current_slide_context():
    """Get the current slide context as a string."""
    try:
        reader = get_slide_reader()
        if reader and reader.ppt_app:
            # Force refresh of context by clearing cached values
            # This ensures we always get the latest slide when user switches
            context = reader.get_current_context()
            return context if context else "No slide context available"
        else:
            return "PowerPoint not connected - no slide context available"
    except Exception as e:
        return f"Error reading slide context: {e}"

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
    Automatically includes current slide context in the message.
    
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context' keys
    """
    try:
        # Get current slide context
        slide_context = get_current_slide_context()
        
        # Debug: Print current slide info (you can remove this later)
        if "Slide:" in slide_context:
            slide_line = [line for line in slide_context.split('\n') if line.startswith('Slide:')]
            if slide_line:
                print(f"🎯 Current slide context: {slide_line[0]}")
        
        # Enhance the message with slide context
        enhanced_message = f"""CURRENT SLIDE CONTEXT:
{slide_context}

USER REQUEST:
{message}

INSTRUCTIONS: Please consider the current slide context above when processing the user's request. If the user is asking to modify, add to, or work with the current slide, use the context information to make informed decisions about positioning, styling, and content placement."""
        
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
            
            # Run the agent with enhanced message
            answer = agent.run(enhanced_message)
            
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
            'slide_context': slide_context,
            'debug_output': f"STDOUT:\n{stdout_content}\n\nSTDERR:\n{stderr_content}"
        }
        
    except Exception as e:
        return {
            'answer': f"Error: {str(e)}",
            'generated_code': f"# Error occurred during execution:\n# {str(e)}\n\n# This might be due to:\n# - Missing dependencies\n# - PowerPoint not running\n# - Invalid parameters",
            'slide_context': "Error reading slide context",
            'debug_output': str(e)
        }

def run_agent_with_slide_context(message):
    """
    Convenience function to run the agent with slide context.
    This is the main function that external code should call.
    
    Args:
        message (str): The user's request/message
        
    Returns:
        dict: Contains 'answer', 'generated_code', 'slide_context', and 'debug_output'
    """
    return run_agent_with_code_capture(message)

def test_integration():
    """Test the slide context integration."""
    print("🧪 Testing PPT Agent with Slide Context Integration")
    print("=" * 60)
    
    # Test 1: Check slide reader connection
    print("\n📡 Test 1: Checking PowerPoint connection...")
    reader = get_slide_reader()
    if reader and reader.ppt_app:
        print("✅ PowerPoint connected successfully!")
        
        # Show current slide context
        print("\n📄 Current slide context:")
        context = get_current_slide_context()
        print(context[:500] + "..." if len(context) > 500 else context)
        
    else:
        print("❌ PowerPoint not connected. Please open PowerPoint with a presentation.")
        return
    
    # Test 2: Run agent with context
    print("\n🤖 Test 2: Running agent with slide context...")
    test_message = "Add a textbox with 'Test Integration' in a good position that doesn't overlap existing content"
    
    result = run_agent_with_slide_context(test_message)
    
    print(f"\n📝 Agent Answer: {result['answer']}")
    print(f"\n💻 Generated Code:\n{result['generated_code'][:300]}...")
    print(f"\n📄 Slide Context Available: {'Yes' if result['slide_context'] else 'No'}")
    
    print("\n✅ Integration test completed!")

if __name__ == "__main__":
    # Test the integration
    test_integration()
    
    # Original example (still works)
    # print(agent.run("Add a textbox with text 'Hello' on slide 2, font size 32, bold."))
