from smolagents import CodeAgent, ToolCallingAgent, OpenAIServerModel, tool, PromptTemplates
from smolagents.monitoring import LogLevel
import os
from dotenv import load_dotenv
import logging
import io
import sys
import win32com.client
import pythoncom
import re
import PIL
from PIL import Image
from typing import Optional
import yaml
from importlib.resources import files

# Load environment variables
load_dotenv()

# Initialize Phoenix tracing
from phoenix_config import initialize_phoenix, trace_tool_call, add_trace_event, trace_function
phoenix_initialized = initialize_phoenix()
if phoenix_initialized:
    print("✅ Phoenix tracing initialized successfully")
else:
    print("⚠️  Phoenix tracing disabled (missing PHOENIX_API_KEY)")

# Import all the existing tools and utilities
from lightning_slide_context_reader import LightningFastPowerPointSlideReader as PowerPointSlideReader
from html_processor import parse_html_text, process_html_lists, apply_html_formatting

# Get OpenAI API key
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found in environment variables. Please check your .env file.")

# Define the models - all using GPT-4o as requested
manager_model = OpenAIServerModel(
    model_id="gpt-4.1",
    api_key=openai_api_key,
    api_base="https://api.openai.com/v1"
)

vision_model = OpenAIServerModel(
    model_id="gpt-4o",
    api_key=openai_api_key,
    api_base="https://api.openai.com/v1"
)

writing_model = OpenAIServerModel(
    model_id="gpt-4o",
    api_key=openai_api_key,
    api_base="https://api.openai.com/v1"
)

# Global slide context reader instance
slide_reader = None

def get_slide_reader():
    """Get or create the global slide reader instance."""
    global slide_reader
    if slide_reader is None:
        try:
            slide_reader = PowerPointSlideReader()
            print("🚀 Slide reader initialized with lightning fast HTML conversion")
        except Exception as e:
            print(f"Warning: Could not initialize slide reader: {e}")
            slide_reader = None
    return slide_reader

def get_current_slide_context(force_refresh: bool = False) -> str:
    """Get the current slide context as a string."""
    try:
        reader = get_slide_reader()
        if reader and reader.ppt_app:
            if force_refresh:
                context = reader.force_refresh_context()
            else:
                context = reader.get_current_context()
            return context if context else "No slide context available"
        else:
            return "PowerPoint not connected - no slide context available"
    except Exception as e:
        return f"Error reading slide context: {e}"

def get_slide_visualizer() -> Optional[object]:
    """Get or create the slide visualizer instance."""
    try:
        from slide_visualizer import SlideVisualizer
        return SlideVisualizer()
    except Exception as e:
        print(f"Warning: Could not initialize slide visualizer: {e}")
        return None

def get_annotated_slide_image() -> Optional[Image.Image]:
    """Get the current slide as an annotated PIL Image for vision analysis."""
    try:
        visualizer = get_slide_visualizer()
        if visualizer:
            return visualizer.get_annotated_slide_as_pil_image(target_width=512)  # type: ignore
        return None
    except Exception as e:
        print(f"Warning: Could not get annotated slide image: {e}")
        return None

# ============================================================================
# MANAGER AGENT TOOLS (Strategic tools for decision making and context)
# ============================================================================

@tool
def get_current_slide_context_tool() -> str:
    """
    Get the current PowerPoint slide context including all objects, their properties, and layout information.
    This provides essential context for understanding the current state of the slide.
    
    Returns:
        str: Detailed slide context with object information, positions, and content
    """
    with trace_tool_call("get_current_slide_context_tool"):
        return get_current_slide_context(force_refresh=True)

@tool
def get_annotated_slide_image_tool() -> Optional[Image.Image]:
    """
    Get the current slide as an annotated PIL Image for vision analysis.

    Returns the annotated slide image, use this tool call before calling the vision agent.
    Store the image in a variable and pass it to the vision agent using images = [image_variable].
    Here is an example of how to call the vision agent:
    image_variable = get_annotated_slide_image_tool()
    visual_analysis = vision_agent(task = "analyze the slide", images = [image_variable])
    
    Returns:
        image: annotated slide image (PIL Image object)
    """
    slide_image = get_annotated_slide_image()
    return slide_image

@tool
def get_object_properties(id: int) -> dict:
    """
    Get detailed information about any object on the slide.
    
    Returns comprehensive details including position, size, type, and content information.
    Use this to inspect objects before making decisions about modifications.

    Args:
        id: The ID of the object to inspect

    Returns:
        dict: Object properties including slide, position, size, type, and content details
    """
    with trace_tool_call("get_object_properties", object_id=id):
        pythoncom.CoInitialize()
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            presentation = ppt_app.ActivePresentation
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if shape.Id == id:
                        props = {
                            "slide": slide.SlideIndex,
                            "id": shape.Id,
                            "name": shape.Name,
                            "left": shape.Left,
                            "top": shape.Top,
                            "width": shape.Width,
                            "height": shape.Height,
                            "rotation": shape.Rotation,
                            "type": shape.Type,
                            "type_name": _get_shape_type_name(shape.Type)
                        }
                        
                        # Add text content if it's a text-containing shape
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            text_content = shape.TextFrame.TextRange.Text
                            props["text_content"] = text_content[:100] + "..." if len(text_content) > 100 else text_content
                        
                        return props
            return {"error": f"Object with ID {id} not found"}
        except Exception as e:
            return {"error": f"Error inspecting object {id}: {str(e)}"}

def _get_shape_type_name(shape_type: int) -> str:
    """Convert PowerPoint shape type number to readable name."""
    type_map = {
        1: "AutoShape", 5: "Freeform", 9: "Group", 11: "Picture", 12: "OLEObject",
        13: "Chart", 14: "Table", 15: "Media", 17: "TextBox", 18: "Content", 19: "SmartArt"
    }
    return type_map.get(shape_type, f"Unknown({shape_type})")

# ============================================================================
# ALL POWERPOINT MANIPULATION TOOLS (for Writing Agent)
# ============================================================================

# Import all the existing PowerPoint tools from the original file
# We'll copy them here to ensure the Writing Agent has access to all of them

@tool
def add_textbox(slide_idx: int = 1, html_text: str = "<b>Sample Text</b>", left: int = 100, top: int = 100, width: int = 400, height: int = 50, font_size: Optional[int] = None, font_name: Optional[str] = None, text_align: str = "left") -> str:
    """
    Add a textbox to a PowerPoint slide with HTML-formatted text.
    HTML Syntax Supported:
        <b>bold text</b> or <strong>bold text</strong> - Bold formatting
        <i>italic text</i> or <em>italic text</em> - Italic formatting
        <s>strikethrough</s> or <del>strikethrough</del> - Strikethrough formatting
        <u>underlined</u> - Underlined text
        <span style="color: red">colored text</span> - Colored text (hex #FF0000 or names)
        <span style="background-color: yellow">highlighted</span> - Background color
        <ul><li>bullet point</li></ul> - Bullet lists
        <ol><li>numbered item</li></ol> - Numbered lists
        <h1>Header 1</h1>, <h2>Header 2</h2>, <h3>Header 3</h3> - Headers

    Args:
        slide_idx: The slide number (1-indexed) to add the textbox to
        html_text: The HTML-formatted text content for the textbox
        left: Left position of the textbox in points
        top: Top position of the textbox in points
        width: Width of the textbox in points
        height: Height of the textbox in points
        font_size: Base font size for the text (optional, headers will be larger)
        font_name: Font name for the text (optional)
        text_align: Text alignment - "left", "center", or "right" (default: "left")

    Returns:
        str: Confirmation message of the textbox addition
    """
    with trace_tool_call("add_textbox", slide_idx=slide_idx, html_text=html_text[:50], 
                        left=left, top=top, width=width, height=height):
        pythoncom.CoInitialize()
        
        try:
            add_trace_event("powerpoint_connection", action="connecting_to_application")
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            presentation = ppt_app.ActivePresentation
            
            # Add slide if needed
            if presentation.Slides.Count < slide_idx:
                slide = presentation.Slides.Add(slide_idx, 12)  # 12 = ppLayoutBlank
            else:
                slide = presentation.Slides(slide_idx)
            
            add_trace_event("html_processing", action="processing_html_content")
            # Process HTML
            processed_text, list_info = process_html_lists(html_text)
            plain_text, format_segments = parse_html_text(processed_text)
            
            # Create the textbox
            add_trace_event("textbox_creation", action="creating_textbox", slide=slide_idx)
            box = slide.Shapes.AddTextbox(1, left, top, width, height)
            text_range = box.TextFrame.TextRange
            
            # Apply HTML formatting
            apply_html_formatting(text_range, plain_text, format_segments)
            
            # Apply header formatting
            for info in list_info:
                if info['type'] == 'header':
                    try:
                        lines = plain_text.split('\n')
                        if info['line'] < len(lines):
                            line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                            line_length = len(lines[info['line']])
                            
                            if line_length > 0:
                                header_range = text_range.Characters(line_start, line_length)
                                level = info['level']
                                if level == 1:
                                    header_range.Font.Size = (font_size or 14) + 8
                                    header_range.Font.Bold = -1
                                elif level == 2:
                                    header_range.Font.Size = (font_size or 14) + 4
                                    header_range.Font.Bold = -1
                                elif level == 3:
                                    header_range.Font.Size = (font_size or 14) + 2
                                    header_range.Font.Bold = -1
                    except Exception as e:
                        print(f"Warning: Could not apply header formatting: {e}")
            
            # Apply global font settings
            if font_name:
                text_range.Font.Name = font_name
            
            # Set text alignment
            alignment_map = {"left": 1, "center": 2, "right": 3}
            if text_align.lower() in alignment_map:
                text_range.ParagraphFormat.Alignment = alignment_map[text_align.lower()]
            
            # Clear slide context cache
            try:
                reader = get_slide_reader()
                if reader:
                    reader.clear_context_cache()
            except Exception:
                pass
            
            add_trace_event("textbox_completed", success=True, text_length=len(plain_text))
            return f"Textbox added to slide {slide_idx} with HTML formatting: {plain_text[:50]}{'...' if len(plain_text) > 50 else ''}"
            
        except Exception as e:
            add_trace_event("textbox_error", error=str(e), error_type=type(e).__name__)
            return f"Error adding textbox: {str(e)}"

@tool
def replace_textbox_content(id: int, html_text: str, slide_idx: int = 1, font_size: Optional[int] = None, font_name: Optional[str] = None, text_align: Optional[str] = None) -> str:
    """
    COMPLETELY REPLACE all text content in a textbox with new HTML-formatted text.
    
    Use this when you want to completely overwrite the existing text content.
    All existing text will be deleted and replaced with the new content.
    
    HTML Syntax Supported:
        <b>bold text</b> or <strong>bold text</strong> - Bold formatting
        <i>italic text</i> or <em>italic text</em> - Italic formatting
        <s>strikethrough</s> or <del>strikethrough</del> - Strikethrough formatting
        <u>underlined</u> - Underlined text
        <span style="color: red">colored text</span> - Colored text (hex #FF0000 or names)
        <span style="background-color: yellow">highlighted</span> - Background color
        <ul><li>bullet point</li></ul> - Bullet lists
        <ol><li>numbered item</li></ol> - Numbered lists
        <h1>Header 1</h1>, <h2>Header 2</h2>, <h3>Header 3</h3> - Headers
    
    Args:
        id: The ID of the textbox to update
        html_text: New HTML-formatted text content (replaces ALL existing text)
        slide_idx: The slide number (1-indexed) containing the textbox (default: 1)
        font_size: Base font size in points (headers will be larger)
        font_name: Font name for the text
        text_align: Text alignment - "left", "center", "right", or "justify"
    
    Returns:
        str: Confirmation message with details of what was updated
    """
    return _update_textbox_internal(
        id=id,
        slide_idx=slide_idx,
        html_text=html_text,
        text_operation="replace",
        font_size=font_size,
        font_name=font_name,
        text_align=text_align
    )

@tool
def modify_text_in_textbox(id: int, find_pattern: str, replacement_text: str, slide_idx: int = 1, regex_flags: str = "IGNORECASE") -> str:
    """
    Find and replace specific text patterns within a textbox while preserving all other text.
    
    This tool modifies only the matching text and keeps everything else unchanged.
    Perfect for tasks like "make 'Company Name' bold" or "change all dates to red".
    
    Args:
        id: The ID of the textbox to modify
        find_pattern: Text pattern to find (can be plain text or regex)
        replacement_text: HTML-formatted text to replace matches with.
            Use HTML syntax like "<b>bold</b>", "<i>italic</i>", "<span style='color: red'>text</span>" etc.
            Set to empty string ("") to delete the matched text.
        slide_idx: The slide number (1-indexed) containing the textbox (default: 1)
        regex_flags: Regex flags like "IGNORECASE" (default: "IGNORECASE")
    
    Returns:
        str: Confirmation message with details of what was replaced
    """
    return _update_textbox_internal(
        id=id,
        slide_idx=slide_idx,
        regex_finder=find_pattern,
        replacement_text=replacement_text,
        regex_flags=regex_flags
    )

@tool
def add_text_to_textbox(id: int, html_text: str, slide_idx: int = 1, position: str = "end") -> str:
    """
    Add new text to the beginning or end of existing textbox content.
    
    This tool preserves all existing text and adds new content before or after it.
    
    Args:
        id: The ID of the textbox to modify
        html_text: HTML-formatted text to add
        slide_idx: The slide number (1-indexed) containing the textbox (default: 1)
        position: Where to add the text - "start" (beginning) or "end" (default)
    
    Returns:
        str: Confirmation message with details of what was added
    """
    operation = "prepend" if position == "start" else "append"
    return _update_textbox_internal(
        id=id,
        slide_idx=slide_idx,
        html_text=html_text,
        text_operation=operation
    )

@tool
def format_textbox_style(id: int, slide_idx: int = 1, font_size: Optional[int] = None, font_name: Optional[str] = None, text_align: Optional[str] = None, 
                        line_spacing: Optional[float] = None, left_margin: Optional[float] = None, right_margin: Optional[float] = None, 
                        top_margin: Optional[float] = None, bottom_margin: Optional[float] = None) -> str:
    """
    Change the formatting and layout properties of a textbox without modifying text content.
    
    Use this to adjust visual appearance like font, alignment, spacing, and margins.
    
    Args:
        id: The ID of the textbox to format
        slide_idx: The slide number (1-indexed) containing the textbox (default: 1)
        font_size: Base font size in points
        font_name: Font name for the text
        text_align: Text alignment - "left", "center", "right", or "justify"
        line_spacing: Line spacing multiplier (1.0 = single, 1.5 = 1.5x, etc.)
        left_margin: Left margin in points
        right_margin: Right margin in points
        top_margin: Top margin in points
        bottom_margin: Bottom margin in points
    
    Returns:
        str: Confirmation message with details of formatting changes
    """
    return _update_textbox_internal(
        id=id,
        slide_idx=slide_idx,
        font_size=font_size,
        font_name=font_name,
        text_align=text_align,
        line_spacing=line_spacing,
        left_margin=left_margin,
        right_margin=right_margin,
        top_margin=top_margin,
        bottom_margin=bottom_margin
    )

@tool
def move_object(id: int, left: int, top: int, slide_idx: int = 1) -> str:
    """
    Move any object (textbox, shape, image, etc.) to new coordinates on the slide.
    
    The slide coordinate system:
    - Origin (0, 0) is at the top-left corner
    - Standard slide is 960 points wide × 540 points tall
    - Measurements are in points (72 points = 1 inch)
    
    Args:
        id: The ID of the object to move
        left: Distance from left edge of slide in points (0-960 for standard slide)
        top: Distance from top edge of slide in points (0-540 for standard slide)
        slide_idx: The slide number (1-indexed) containing the object (default: 1)
    
    Returns:
        str: Confirmation message with the object's new position
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Search specified slide first, then fall back to all slides
        if slide_idx <= presentation.Slides.Count:
            slide = presentation.Slides(slide_idx)
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape.Left = left
                    shape.Top = top
                    return f"Moved object {id} to position ({left}, {top}) on slide {slide.SlideIndex}"
        
        # Fallback: search all slides if not found on specified slide
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape.Left = left
                    shape.Top = top
                    return f"Moved object {id} to position ({left}, {top}) on slide {slide.SlideIndex}"
        return f"Object with ID {id} not found"
    except Exception as e:
        return f"Error moving object {id}: {str(e)}"

@tool
def resize_object(id: int, width: int, height: int, slide_idx: int = 1) -> str:
    """
    Change the size of any object (textbox, shape, image, etc.) to new dimensions.
    
    Args:
        id: The ID of the object to resize
        width: New width in points
        height: New height in points
        slide_idx: The slide number (1-indexed) containing the object (default: 1)
    
    Returns:
        str: Confirmation message with the object's new dimensions
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Search specified slide first, then fall back to all slides
        if slide_idx <= presentation.Slides.Count:
            slide = presentation.Slides(slide_idx)
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape.Width = width
                    shape.Height = height
                    return f"Resized object {id} to {width}×{height} points on slide {slide.SlideIndex}"
        
        # Fallback: search all slides if not found on specified slide
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape.Width = width
                    shape.Height = height
                    return f"Resized object {id} to {width}×{height} points on slide {slide.SlideIndex}"
        return f"Object with ID {id} not found"
    except Exception as e:
        return f"Error resizing object {id}: {str(e)}"

@tool
def position_and_resize_object(id: int, left: int, top: int, width: int, height: int, slide_idx: int = 1) -> str:
    """
    Move and resize an object in a single operation for precise positioning.
    
    Useful when you need to set both position and size to avoid multiple operations.
    
    Args:
        id: The ID of the object to position and resize
        left: Distance from left edge of slide in points
        top: Distance from top edge of slide in points
        width: New width in points
        height: New height in points
        slide_idx: The slide number (1-indexed) containing the object (default: 1)
    
    Returns:
        str: Confirmation message with the object's new position and size
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Search specified slide first, then fall back to all slides
        if slide_idx <= presentation.Slides.Count:
            slide = presentation.Slides(slide_idx)
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape.Left = left
                    shape.Top = top
                    shape.Width = width
                    shape.Height = height
                    return f"Positioned object {id} at ({left}, {top}) with size {width}×{height} on slide {slide.SlideIndex}"
        
        # Fallback: search all slides if not found on specified slide
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape.Left = left
                    shape.Top = top
                    shape.Width = width
                    shape.Height = height
                    return f"Positioned object {id} at ({left}, {top}) with size {width}×{height} on slide {slide.SlideIndex}"
        return f"Object with ID {id} not found"
    except Exception as e:
        return f"Error positioning object {id}: {str(e)}"

@tool
def copy_object_to_slide(id: int, target_slide_idx: int, new_left: Optional[int] = None, new_top: Optional[int] = None) -> int:
    """
    Copy an object to another slide, optionally positioning it at specific coordinates.
    
    The original object remains unchanged. A new copy is created on the target slide.
    
    Args:
        id: The ID of the object to copy
        target_slide_idx: Slide number to copy the object to (1-indexed)
        new_left: Optional new left position for the copy (preserves original position if not specified)
        new_top: Optional new top position for the copy (preserves original position if not specified)
    
    Returns:
        int: The ID of the newly created copy, or -1 if operation failed
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Find source object
        source_shape = None
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    source_shape = shape
                    break
            if source_shape:
                break
        
        if not source_shape:
            return -1
        
        # Create target slide if needed
        if presentation.Slides.Count < target_slide_idx:
            target_slide = presentation.Slides.Add(target_slide_idx, 12)  # 12 = ppLayoutBlank
        else:
            target_slide = presentation.Slides(target_slide_idx)
        
        # Copy and paste
        source_shape.Copy()
        pasted = target_slide.Shapes.Paste()
        
        if pasted and pasted.Count > 0:
            new_shape = pasted[0]
            new_id = new_shape.Id
            
            # Position the copy if coordinates specified
            if new_left is not None:
                new_shape.Left = new_left
            if new_top is not None:
                new_shape.Top = new_top
            
            return new_id
        else:
            return -1
            
    except Exception as e:
        print(f"Error copying object {id}: {str(e)}")
        return -1

@tool
def duplicate_object_on_same_slide(id: int, slide_idx: int = 1, offset_left: int = 20, offset_top: int = 20) -> int:
    """
    Create a duplicate of an object on the same slide with a slight position offset.
    
    Useful for creating multiple similar objects quickly.
    
    Args:
        id: The ID of the object to duplicate
        slide_idx: The slide number (1-indexed) containing the object (default: 1)
        offset_left: How many points to move the duplicate to the right (default: 20)
        offset_top: How many points to move the duplicate down (default: 20)
    
    Returns:
        int: The ID of the newly created duplicate, or -1 if operation failed
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Find source object - search specified slide first
        source_shape = None
        if slide_idx <= presentation.Slides.Count:
            slide = presentation.Slides(slide_idx)
            for shape in slide.Shapes:
                if shape.Id == id:
                    source_shape = shape
                    break
        
        # Fallback: search all slides if not found on specified slide
        if not source_shape:
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if shape.Id == id:
                        source_shape = shape
                        break
                if source_shape:
                    break
        
        if not source_shape:
            return -1
        
        # Duplicate on same slide
        dup = source_shape.Duplicate()
        if dup and dup.Count > 0:
            new_shape = dup[0]
            # Offset the position slightly
            new_shape.Left = source_shape.Left + offset_left
            new_shape.Top = source_shape.Top + offset_top
            return new_shape.Id
        else:
            return -1
            
    except Exception as e:
        print(f"Error duplicating object {id}: {str(e)}")
        return -1

@tool
def delete_object(id: int, slide_idx: int = 1) -> str:
    """
    Permanently delete an object from the slide.
    
    ⚠️ WARNING: This action cannot be undone programmatically.
    
    Args:
        id: The ID of the object to delete
        slide_idx: The slide number (1-indexed) containing the object (default: 1)
    
    Returns:
        str: Confirmation message of deletion
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Search specified slide first, then fall back to all slides
        if slide_idx <= presentation.Slides.Count:
            slide = presentation.Slides(slide_idx)
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape_name = shape.Name
                    slide_num = slide.SlideIndex
                    shape.Delete()
                    
                    # Clear slide context cache after deletion
                    try:
                        reader = get_slide_reader()
                        if reader:
                            reader.clear_context_cache()
                    except Exception:
                        pass
                    
                    return f"Deleted object '{shape_name}' (ID: {id}) from slide {slide_num}"
        
        # Fallback: search all slides if not found on specified slide
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    shape_name = shape.Name
                    slide_num = slide.SlideIndex
                    shape.Delete()
                    
                    # Clear slide context cache after deletion
                    try:
                        reader = get_slide_reader()
                        if reader:
                            reader.clear_context_cache()
                    except Exception:
                        pass
                    
                    return f"Deleted object '{shape_name}' (ID: {id}) from slide {slide_num}"
        return f"Object with ID {id} not found"
    except Exception as e:
        return f"Error deleting object {id}: {str(e)}"

def _update_textbox_internal(id: int, slide_idx: int = 1, html_text: Optional[str] = None, text_operation: str = "replace", regex_finder: Optional[str] = None, replacement_text: Optional[str] = None, regex_flags: str = "IGNORECASE", font_size: Optional[int] = None, font_name: Optional[str] = None, text_align: Optional[str] = None, line_spacing: Optional[float] = None, left_margin: Optional[float] = None, right_margin: Optional[float] = None, top_margin: Optional[float] = None, bottom_margin: Optional[float] = None) -> str:
    """Internal implementation for textbox updates. Do not call directly."""
    pythoncom.CoInitialize()
    
    # INPUT VALIDATION: Prevent conflicting parameter combinations
    if html_text is not None and text_operation == "replace" and regex_finder is not None:
        return f"ERROR: Cannot use both 'html_text' with operation='replace' AND 'regex_finder'. Choose ONE approach:\n" \
               f"- For complete text replacement: use 'html_text' parameter only\n" \
               f"- For partial text replacement: use 'regex_finder' + 'replacement_text' only"
    
    if regex_finder and not replacement_text:
        return f"ERROR: When using 'regex_finder', you must specify 'replacement_text' for the replacement."
    
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Find the textbox by ID - search only the specified slide if slide_idx provided
        target_shape = None
        target_slide = None
        
        if slide_idx:
            # Search only the specified slide
            if slide_idx <= presentation.Slides.Count:
                slide = presentation.Slides(slide_idx)
                for shape in slide.Shapes:
                    if shape.Id == id:
                        target_shape = shape
                        target_slide = slide
                        break
        else:
            # Fallback: search all slides if slide_idx not provided
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if shape.Id == id:
                        target_shape = shape
                        target_slide = slide
                        break
                if target_shape:
                    break
        
        if not target_shape:
            return f"Shape with ID {id} not found"
        
        # Verify it's a shape that can contain text
        if not hasattr(target_shape, 'TextFrame'):
            return f"Shape with ID {id} is not a textbox or doesn't support text"
        
        if not target_shape.TextFrame.HasText and not html_text:
            return f"Shape with ID {id} has no text and no new text provided"
        
        updates_made = []
        
        # Handle text content updates
        if html_text is not None:
            current_text = target_shape.TextFrame.TextRange.Text if target_shape.TextFrame.HasText else ""
            
            if text_operation == "replace":
                # Process HTML and apply formatting
                processed_text, list_info = process_html_lists(html_text)
                plain_text, format_segments = parse_html_text(processed_text)
                apply_html_formatting(target_shape.TextFrame.TextRange, plain_text, format_segments)
                
                # Apply header formatting
                for info in list_info:
                    if info['type'] == 'header':
                        try:
                            lines = plain_text.split('\n')
                            if info['line'] < len(lines):
                                line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                                line_length = len(lines[info['line']])
                                
                                if line_length > 0:
                                    header_range = target_shape.TextFrame.TextRange.Characters(line_start, line_length)
                                    level = info['level']
                                    if level == 1:
                                        header_range.Font.Size = (font_size or 14) + 8
                                        header_range.Font.Bold = -1
                                    elif level == 2:
                                        header_range.Font.Size = (font_size or 14) + 4
                                        header_range.Font.Bold = -1
                                    elif level == 3:
                                        header_range.Font.Size = (font_size or 14) + 2
                                        header_range.Font.Bold = -1
                        except Exception as e:
                            print(f"Warning: Could not apply header formatting: {e}")
                
                updates_made.append(f"replaced text with HTML-formatted content")
                    
            elif text_operation == "append":
                # For append/prepend, process the combined text to apply HTML formatting
                combined_text = current_text + html_text
                processed_text, list_info = process_html_lists(combined_text)
                plain_text, format_segments = parse_html_text(processed_text)
                apply_html_formatting(target_shape.TextFrame.TextRange, plain_text, format_segments)
                
                # Apply header formatting if any headers are present
                for info in list_info:
                    if info['type'] == 'header':
                        try:
                            lines = plain_text.split('\n')
                            if info['line'] < len(lines):
                                line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                                line_length = len(lines[info['line']])
                                
                                if line_length > 0:
                                    header_range = target_shape.TextFrame.TextRange.Characters(line_start, line_length)
                                    level = info['level']
                                    if level == 1:
                                        header_range.Font.Size = (font_size or 14) + 8
                                        header_range.Font.Bold = -1
                                    elif level == 2:
                                        header_range.Font.Size = (font_size or 14) + 4
                                        header_range.Font.Bold = -1
                                    elif level == 3:
                                        header_range.Font.Size = (font_size or 14) + 2
                                        header_range.Font.Bold = -1
                        except Exception as e:
                            print(f"Warning: Could not apply header formatting: {e}")
                
                updates_made.append(f"appended HTML-formatted text: '{html_text[:30]}{'...' if len(html_text) > 30 else ''}'")
                
            elif text_operation == "prepend":
                # For prepend, process the combined text to apply HTML formatting
                combined_text = html_text + current_text
                processed_text, list_info = process_html_lists(combined_text)
                plain_text, format_segments = parse_html_text(processed_text)
                apply_html_formatting(target_shape.TextFrame.TextRange, plain_text, format_segments)
                
                # Apply header formatting if any headers are present
                for info in list_info:
                    if info['type'] == 'header':
                        try:
                            lines = plain_text.split('\n')
                            if info['line'] < len(lines):
                                line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                                line_length = len(lines[info['line']])
                                
                                if line_length > 0:
                                    header_range = target_shape.TextFrame.TextRange.Characters(line_start, line_length)
                                    level = info['level']
                                    if level == 1:
                                        header_range.Font.Size = (font_size or 14) + 8
                                        header_range.Font.Bold = -1
                                    elif level == 2:
                                        header_range.Font.Size = (font_size or 14) + 4
                                        header_range.Font.Bold = -1
                                    elif level == 3:
                                        header_range.Font.Size = (font_size or 14) + 2
                                        header_range.Font.Bold = -1
                        except Exception as e:
                            print(f"Warning: Could not apply header formatting: {e}")
                
                updates_made.append(f"prepended HTML-formatted text: '{html_text[:30]}{'...' if len(html_text) > 30 else ''}'")
        
        # Handle regex-based text replacement
        if regex_finder:
            if not target_shape.TextFrame.HasText:
                return f"Cannot use regex on empty textbox {id}"
            
            current_text = target_shape.TextFrame.TextRange.Text
            
            # Parse regex flags
            flags = 0
            if "IGNORECASE" in regex_flags.upper():
                flags |= re.IGNORECASE
            if "MULTILINE" in regex_flags.upper():
                flags |= re.MULTILINE
            if "DOTALL" in regex_flags.upper():
                flags |= re.DOTALL
            
            try:
                matches = list(re.finditer(regex_finder, current_text, flags))
                
                if matches:
                    if replacement_text is not None:
                        # Check if replacement contains HTML formatting
                        if any(marker in replacement_text for marker in ['<b>', '<i>', '<u>', '<s>', '<span', '<strong>', '<em>']):
                            processed_replacement, _ = process_html_lists(replacement_text)
                            plain_replacement, format_segments = parse_html_text(processed_replacement)
                            
                            # Process matches in reverse order to maintain position indices
                            for match in reversed(matches):
                                match_start = match.start()
                                match_end = match.end()
                                match_length = match_end - match_start
                                
                                # Replace this specific match in the textbox without affecting the rest
                                if match_length > 0:
                                    match_range = target_shape.TextFrame.TextRange.Characters(match_start + 1, match_length)
                                    match_range.Text = plain_replacement
                                    
                                    # Apply formatting to the replacement text
                                    replacement_start_pos = match_start + 1
                                    
                                    for segment in format_segments:
                                        try:
                                            absolute_start = replacement_start_pos + segment['start'] - 1
                                            segment_length = segment['length']
                                            
                                            if segment_length > 0:
                                                char_range = target_shape.TextFrame.TextRange.Characters(absolute_start, segment_length)
                                                formatting = segment['formatting']
                                                if formatting.get('bold'):
                                                    char_range.Font.Bold = -1
                                                if formatting.get('italic'):
                                                    char_range.Font.Italic = -1
                                                if formatting.get('underline'):
                                                    char_range.Font.Underline = -1
                                                if formatting.get('strikethrough'):
                                                    try:
                                                        char_range.Font.Strike = -1
                                                    except:
                                                        pass
                                                if formatting.get('color'):
                                                    try:
                                                        color_value = formatting['color']
                                                        if color_value.startswith('#'):
                                                            hex_color = color_value[1:]
                                                            if len(hex_color) == 6:
                                                                r = int(hex_color[0:2], 16)
                                                                g = int(hex_color[2:4], 16) 
                                                                b = int(hex_color[4:6], 16)
                                                                rgb_color = r + (g * 256) + (b * 65536)
                                                                char_range.Font.Color.RGB = rgb_color
                                                        else:
                                                            color_map = {
                                                                'red': 255, 'blue': 16711680, 'green': 65280,
                                                                'yellow': 65535, 'orange': 33023, 'purple': 8388736,
                                                                'black': 0, 'white': 16777215
                                                            }
                                                            if color_value.lower() in color_map:
                                                                char_range.Font.Color.RGB = color_map[color_value.lower()]
                                                    except Exception as e:
                                                        print(f"Warning: Could not apply color {color_value}: {e}")
                                                        
                                        except Exception as e:
                                            print(f"Warning: Could not format segment at position {absolute_start}: {e}")
                                            
                                    # Update the current_text to reflect the change for subsequent matches
                                    current_text = target_shape.TextFrame.TextRange.Text
                        else:
                            # Simple text replacement without HTML formatting
                            new_text = re.sub(regex_finder, replacement_text, current_text, flags=flags)
                            target_shape.TextFrame.TextRange.Text = new_text
                        
                        updates_made.append(f"replaced {len(matches)} regex matches with '{replacement_text}'")
                else:
                    updates_made.append(f"no matches found for regex pattern '{regex_finder}'")
                    
            except re.error as e:
                return f"Invalid regex pattern '{regex_finder}': {str(e)}"
        
        # Apply global font settings that don't conflict with markdown
        if target_shape.TextFrame.HasText:
            text_range = target_shape.TextFrame.TextRange
            
            if font_name:
                text_range.Font.Name = font_name
                updates_made.append(f"set font to '{font_name}' for entire text")
            
            # Apply paragraph formatting
            if text_align is not None:
                alignment_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
                if text_align.lower() in alignment_map:
                    text_range.ParagraphFormat.Alignment = alignment_map[text_align.lower()]
                    updates_made.append(f"set text alignment to {text_align}")
            
            if line_spacing is not None:
                text_range.ParagraphFormat.LineRuleWithin = 1  # Multiple line spacing
                text_range.ParagraphFormat.SpaceWithin = line_spacing
                updates_made.append(f"set line spacing to {line_spacing}")
        
        # Apply text margins
        if left_margin is not None:
            target_shape.TextFrame.MarginLeft = left_margin
            updates_made.append(f"set left margin to {left_margin}")
        
        if right_margin is not None:
            target_shape.TextFrame.MarginRight = right_margin
            updates_made.append(f"set right margin to {right_margin}")
        
        if top_margin is not None:
            target_shape.TextFrame.MarginTop = top_margin
            updates_made.append(f"set top margin to {top_margin}")
        
        if bottom_margin is not None:
            target_shape.TextFrame.MarginBottom = bottom_margin
            updates_made.append(f"set bottom margin to {bottom_margin}")
        
        # Clear slide context cache to ensure fresh context on next request
        try:
            reader = get_slide_reader()
            if reader:
                reader.clear_context_cache()
        except Exception:
            pass
        
        if updates_made:
            slide_index = target_slide.SlideIndex if target_slide else "unknown"
            return f"Updated textbox {id} on slide {slide_index}: {'; '.join(updates_made)}"
        else:
            return f"No updates specified for textbox {id}"
    
    except Exception as e:
        return f"Error updating textbox {id}: {str(e)}"

# ============================================================================
# VISION AGENT SETUP
# ============================================================================

vision_agent_instructions = """
You are a PowerPoint slide visual analysis expert specialized in providing detailed, actionable feedback for slide improvements. Your analysis directly informs a Writing Agent that will implement your suggestions using PowerPoint automation tools.

CAPABILITIES:
- Analyze visual layout, spacing, alignment, and design aesthetics with precision
- Identify objects by their ID labels (shown as yellow "ID:X" tags in green bounding boxes)
- Provide specific actionable suggestions with exact object IDs and measurements
- Evaluate design principles: hierarchy, balance, contrast, and professional appearance
- Answer specific questions about visual elements and spatial relationships

ANALYSIS FRAMEWORK:
1. **Holistic Assessment**: Overall slide composition, visual flow, and professional appearance
2. **Object-Specific Analysis**: Individual elements, their positioning, sizing, and styling
3. **Spatial Relationships**: Alignment, spacing, distribution, and visual hierarchy
4. **Design Quality**: Color harmony, typography consistency, visual balance
5. **Actionable Recommendations**: Specific improvements with implementation details

RESPONSE FORMAT:
Always use the final_answer tool to provide your complete analysis including:

**VISUAL DESCRIPTION:**
- Comprehensive overview of slide contents and layout
- Identification of key elements and their relationships
- Overall visual impression and design assessment

**SPATIAL ANALYSIS:**
- Precise measurements and positioning feedback
- Alignment issues with specific corrections needed
- Spacing inconsistencies and recommended adjustments
- Distribution of elements and balance assessment

**AESTHETIC FEEDBACK:**
- Professional appearance evaluation
- Color scheme and typography assessment
- Visual hierarchy effectiveness
- Design principle adherence (contrast, balance, unity)

**SPECIFIC ACTIONABLE SUGGESTIONS:**
- Exact object IDs with recommended changes
- Precise measurements for repositioning/resizing
- Specific formatting improvements
- Priority ranking of suggested modifications

MEASUREMENT PRECISION:
- Reference the PowerPoint coordinate system (0,0 = top-left, 960×540 slide)
- Provide specific pixel/point measurements when relevant
- Use relative positioning references ("move 20 points right", "increase height by 50 points")
- Consider standard spacing conventions (margins, padding, alignment grids)

IMPLEMENTATION FOCUS:
Your suggestions will be implemented by a Writing Agent with PowerPoint automation tools, so ensure recommendations are:
- Technically feasible with PowerPoint COM interface
- Specific enough to execute without ambiguity
- Prioritized by impact and importance
- Comprehensive but not overwhelming

EXAMPLE RESPONSE STRUCTURE:
"**VISUAL DESCRIPTION:** I see a slide with a title, two bullet point sections, and an image...

**SPATIAL ANALYSIS:** The title (ID:15) is well-centered horizontally but positioned too high at Y:50 - recommend moving to Y:80 for better proportions. The body text (ID:23) appears cramped against the left edge at X:50 - move to X:100 for proper margin...

**AESTHETIC FEEDBACK:** The slide demonstrates good contrast but suffers from inconsistent spacing. Typography is professional but could benefit from size hierarchy adjustments...

**SPECIFIC ACTIONABLE SUGGESTIONS:**
1. HIGH PRIORITY: Move title ID:15 from current position to (430, 80) for better vertical balance
2. MEDIUM PRIORITY: Increase font size of ID:23 from current to 18pt for improved readability
3. LOW PRIORITY: Adjust spacing between bullet points in ID:23 by setting line spacing to 1.5..."

Always be specific, reference object IDs, provide measurements, and focus on implementable improvements that enhance professional appearance and readability.
"""

# ============================================================================
# WRITING AGENT SETUP  
# ============================================================================

writing_agent_instructions = """
You are a PowerPoint automation specialist who executes slide modifications using specialized tools and code. You are designed to work as part of a multi-agent system where you receive instructions from a Manager Agent.

To solve tasks, you have been given access to PowerPoint-specific tools that are Python functions you can call with code.
You must plan forward to proceed in a series of steps, in a cycle of 'Thought:', '<code>', and 'Observation:' sequences.

At each step, in the 'Thought:' sequence, you should first explain your reasoning towards solving the PowerPoint task and the tools you want to use.
Then in the '<code>' sequence, you should write the code in simple Python. The code sequence must end with '</code>' sequence.
During each intermediate step, you can use 'print()' to save whatever important information you will then need.
These print outputs will then appear in the 'Observation:' field, which will be available as input for the next step.
In the end you have to return a final answer using the final_answer tool.

*POWERPOINT AUTOMATION CAPABILITIES:*
- Add, modify, move, resize, and delete PowerPoint objects (textboxes, shapes, images)
- Apply HTML formatting to text content with rich styling options
- Handle precise positioning and layout adjustments
- Manage object properties, styling, and visual consistency
- Copy and duplicate objects across slides
- Format text with HTML tags: <b>bold</b>, <i>italic</i>, <u>underlined</u>, <span style="color: red">colored</span>, etc.

*COORDINATE SYSTEM (CRITICAL):*
- Origin (0,0) = top-left corner of slide
- Standard slide dimensions: 960 points wide × 540 points tall
- All measurements in points (72 points = 1 inch)
- X-axis: 0 (left edge) to 960 (right edge)
- Y-axis: 0 (top edge) to 540 (bottom edge)

*POWERPOINT-SPECIFIC RULES:*
1. *Object ID Management*: Always use object IDs from slide context provided by Manager Agent for reliable reference
2. *Positioning Awareness*: Consider existing content positioning when adding new elements to avoid overlaps
3. *Style Consistency*: Match existing fonts, colors, and styles when appropriate for visual consistency
4. *Efficient Tool Usage*: Use multiple tools together when they accomplish related goals efficiently
5. *Error Prevention*: Log detailed information about tool execution, especially errors, to help with debugging
6. *HTML Formatting*: Leverage HTML tags for rich text formatting instead of basic text
7. *Layout Planning*: Think about slide layout and visual hierarchy when positioning elements
8. *COM Error Handling*: Be prepared for PowerPoint COM interface errors and log them clearly
9. *Batch Operations*: When possible, group similar operations together for efficiency
10. *State Preservation*: Maintain awareness of slide state changes between operations



*TASK CONTEXT:*
You will receive instructions from the Manager Agent that include:
- Current slide context with existing object IDs and positions
- Specific modification tasks to execute
- Overall goal of the PowerPoint automation task
- Visual feedback from Vision Agent when applicable

*CRITICAL: SLIDE CONTEXT AWARENESS*
ALWAYS work based on the current slide context provided by the Manager Agent:
- Use the exact slide numbers mentioned in the context
- Reference only object IDs that exist in the current slide context
- Verify slide indices before executing any slide-specific operations
- If context mentions "Slide 2" or specific slide numbers, use those exact numbers in your tool calls
- Never assume slide numbers - always use what's provided in the context
- When adding new content, consider the slide number where the user wants the content placed

*INFORMATION FLOW:*
- Print important information during tool execution for debugging
- Log object IDs after creating new objects
- Report positioning and sizing details when relevant
- Confirm successful completion of modifications
- Log any errors with sufficient detail for troubleshooting

*WORKFLOW APPROACH:*
1. Analyze the task and current slide context provided by Manager Agent - VERIFY SLIDE NUMBERS
2. Plan the sequence of PowerPoint operations needed for the CORRECT slide
3. Execute tools step-by-step with proper error checking and logging
4. Verify positioning and layout as you work
5. Consider visual hierarchy and design principles
6. Provide clear final answer confirming task completion

*SLIDE VERIFICATION CHECKLIST:*
- Check which slide number is mentioned in the task context
- Use the exact slide index provided (slide_idx parameter)
- If working with existing objects, verify they exist on the target slide
- When in doubt, print the slide context to confirm you're working on the right slide





*WORKING CODE EXAMPLES:*
The following are examples of how to properly use the PowerPoint tools. These are just examples - you may need to write completely different code depending on your specific task:

**Example 1: Adding a Centered Headline (with slide context verification)**
# ALWAYS check the slide context first to determine correct slide number
print("Working on slide based on context provided by Manager Agent")

# Constants for the slide dimensions
slide_width = 960
slide_height = 540

# Variables for headline textbox
headline_text = "Why Valorant is So Cool"
headline_left = slide_width // 2 - 200  # Centered horizontally
headline_top = 20  # Positioned at the top of the slide
headline_width = 400  # A reasonable width for headline text
headline_height = 50  # Height for the headline space

# Use the EXACT slide number from the task context (example shows slide 1, but use what Manager Agent provides)
target_slide_idx = 1  # This should come from the task context - replace with actual slide number

# Add the headline textbox with the specified styling
headline_result = add_textbox(
    slide_idx=target_slide_idx,  # Use the verified slide index
    html_text=f"<b style='font-size:32px'>{headline_text}</b>",
    left=headline_left,
    top=headline_top,
    width=headline_width,
    height=headline_height,
    font_size=32,
    text_align="center"
)
print(headline_result)

**Example 2: Adding Detailed Content with HTML Formatting (with slide verification)**
# Variables for the detailed content textbox
detail_content = '''
<b>Valorant</b> is a tactical first-person shooter. Here's why it's so cool:
<ul>
  <li><b>Unique Agents:</b> Each agent has special skills, bringing variety and strategy to every match.</li>
  <li><b>Teamwork:</b> Winning requires real teamwork and tactical planning, creating intense gameplay moments.</li>
  <li><b>Competitive Spirit:</b> Valorant's ranked mode lets players test their skills against others.</li>
  <li><b>Stunning Design:</b> The maps and visual effects are bright, stylish, and full of personality.</li>
  <li><b>Constant Updates:</b> Riot Games regularly adds new agents, maps, and content.</li>
</ul>
Valorant is not just another shooter - it's a thrilling, ever-evolving esport.
'''
detail_left = 50  # Ensure enough margin on the left
detail_top = 100  # Below the headline with some spacing
detail_width = slide_width - 100  # Leave margin on both sides for readability
detail_height = 400  # Leave space at bottom of the slide

# IMPORTANT: Use the correct slide index from the task context
target_slide_idx = 1  # Replace with actual slide number from Manager Agent context

# Add the detailed content textbox with the specified formatting
detail_result = add_textbox(
    slide_idx=target_slide_idx,  # Use the verified slide index
    html_text=detail_content,
    left=detail_left,
    top=detail_top,
    width=detail_width,
    height=detail_height,
    font_size=14,
    text_align="left"
)
print(detail_result)

**Example 3: Proper Final Answer**
final_answer("The requested content has been successfully added to your slide: A prominent, centered headline textbox with the title has been placed at the top of the slide, using bold and large font for clear emphasis. Below the headline, a spacious, centrally positioned detailed textbox has been inserted containing well-formatted HTML bullet points and short paragraphs. Both text boxes are laid out with good spacing from the slide edges for readability and a coherent, visually appealing result.")

**Key Patterns from Examples:**
- ALWAYS verify slide numbers from the task context before executing tools
- Define clear variables for positioning and dimensions
- Use slide constants (slide_width=960, slide_height=540) for calculations
- Calculate positions relative to slide dimensions for proper layout
- Use descriptive variable names and comments
- Print results after each tool call for debugging
- Use HTML formatting effectively for rich text styling
- Provide detailed final_answer with summary of actions taken
- Reference correct slide indices in all tool calls (slide_idx parameter)
- When working with existing objects, ensure they exist on the target slide

{{tool_descriptions}}

{{managed_agents_descriptions}}

Remember: These are just examples! Your actual code should be tailored to the specific task you're given.
*CORE CODING RULES:*
1. Always provide a 'Thought:' sequence, and a '<code>' sequence ending with '</code>', else you will fail.
2. Use only variables that you have defined!
3. Always use the right arguments for tools. Use arguments directly like 'result = add_textbox(slide_idx=1, html_text="<b>Hello</b>", left=100, top=50, width=200, height=100, font_size=14, font_name="Arial", text_align="center")'
4. Don't chain too many sequential tool calls in the same code block, especially when output format is unpredictable
5. Call a tool only when needed, and never re-do a tool call with the exact same parameters
6. Don't name any new variable with the same name as a tool: for instance don't name a variable 'final_answer'
7. Never create any notional variables in your code, as having these in your logs will derail you from the true variables
8. You can use imports from: {{authorized_imports}}
9. The state persists between code executions: variables and imports persist across steps
10. Don't give up! You're in charge of solving the task, not providing directions to solve it
Focus on precise execution of PowerPoint operations. Work systematically and always consider the visual impact of your modifications on the overall slide design. Remember that you are creating presentations that should be visually appealing and professionally formatted.

Now Begin! Execute PowerPoint automation tasks with precision and attention to detail.
"""

# ============================================================================
# MANAGER AGENT SETUP
# ============================================================================

manager_agent_instructions = """
You are an intelligent PowerPoint Assistant Manager that orchestrates a team of specialized agents to deliver comprehensive slide automation solutions.

YOUR TEAM:
1. Vision Agent: Analyzes slide visuals and provides aesthetic feedback with specific, actionable suggestions
2. Writing Agent: Executes all PowerPoint modifications using specialized tools with step-by-step code execution

DECISION MAKING FRAMEWORK:
You should call the Vision Agent when the user request involves:
- Visual improvements ("make it look better", "improve design", "fix alignment", "enhance layout")
- Layout analysis ("how does this look", "what's wrong with the spacing", "analyze the design")
- Aesthetic feedback ("make it more professional", "improve the visual appeal", "better color scheme")
- Questions about visual elements ("what do you see", "describe the slide", "identify issues")
- Design validation ("does this look good", "review the layout", "check alignment")

IMPORTANT NOTE REGARDING VISION AGENT:
- You must call get_annotated_slide_image_tool() before calling the vision agent
- Store the image in a variable and pass it to the vision agent using images = [image_variable]
- Do all this in a single step and one action itself (don't use multiple steps)
- Example: image = get_annotated_slide_image_tool(); vision_feedback = vision_agent(task="analyze layout", images=[image])

You should call the Writing Agent when:
- User wants to add/modify content (text, objects, formatting)
- User wants to move/resize objects
- User wants to apply formatting changes
- User has specific modification requests
- Implementing suggestions from Vision Agent feedback

WORKFLOW COORDINATION:
1. Always start by getting current slide context using get_current_slide_context_tool()
2. Analyze the user request to determine which agents are needed
3. If visual analysis is needed, call Vision Agent with annotated slide image
4. Pass comprehensive context and clear instructions to the Writing Agent
5. When Vision Agent provides feedback, translate it into specific actionable tasks for Writing Agent
6. Provide a final summary to the user confirming all completed actions

CONTEXT SHARING BEST PRACTICES:
- Always share the current slide context with your agents
- Pass specific findings and suggestions from Vision Agent to Writing Agent
- Provide clear, detailed, and actionable instructions to the Writing Agent
- Include object IDs, specific positions, and formatting requirements
- Ensure Writing Agent has all necessary context to execute tasks independently

COMMUNICATION GUIDELINES:
- Be specific about what needs to be done (not just what the goal is)
- Include relevant object IDs from the slide context
- Specify exact positions, sizes, and formatting when available
- Translate high-level user requests into concrete technical tasks
- Provide context about existing slide elements that might be affected

QUALITY ASSURANCE:
- Verify that instructions to Writing Agent are complete and actionable
- Ensure all necessary context is provided before task execution
- Confirm that Vision Agent feedback is properly translated into implementable tasks
- Check that the final result addresses the user's original request

Remember: You coordinate the workflow and provide strategic direction, but the Writing Agent does all PowerPoint modifications using its specialized tools and step-by-step code execution approach.
"""

# ============================================================================
# CREATE THE MULTI-AGENT SYSTEM
# ============================================================================

class MultiAgentPPTSystem:
    def __init__(self):
        """Initialize the multi-agent PowerPoint system."""
        
        # Create Vision Agent (ToolCallingAgent)
        self.vision_agent = ToolCallingAgent(
            model=vision_model,
            tools=[],  # Vision agent only uses final_answer tool (built-in)
            instructions=vision_agent_instructions,
            name="vision_agent",
            description="Analyzes slide visuals and provides aesthetic feedback with specific suggestions. Before calling this agent, you must call the get_annotated_slide_image_tool() and store the image in a variable and pass it to the vision agent using images = [image_variable]. An example of calling this agent is: vision_agent(task = 'analyze the slide', images = [image_variable]). Note: Do not use additional_args to pass the image to the vision agent, use images = [image_variable] instead.",
            verbosity_level=LogLevel.DEBUG
        )
        
        # Import all PowerPoint manipulation tools for Writing Agent
        # (We'll need to import/copy all the existing tools here)
        
        # Create Writing Agent (CodeAgent) with custom system prompt
        # Load default code agent prompts and override system_prompt
        try:
            defaults = yaml.safe_load(
                files("smolagents.prompts").joinpath("code_agent.yaml").read_text()
            )
            # Override only the system_prompt with our enhanced instructions
            defaults["system_prompt"] = writing_agent_instructions
            
            self.writing_agent = CodeAgent(
                model=writing_model,
                tools=[
                    add_textbox,
                    replace_textbox_content,
                    modify_text_in_textbox,
                    add_text_to_textbox,
                    format_textbox_style,
                    move_object,
                    resize_object,
                    position_and_resize_object,
                    copy_object_to_slide,
                    duplicate_object_on_same_slide,
                    delete_object
                ],
                prompt_templates=PromptTemplates(**defaults),
                name="writing_agent", 
                description="Executes PowerPoint modifications using specialized tools",
                max_steps=5,
                verbosity_level=LogLevel.DEBUG
            )
        except Exception as e:
            print(f"Warning: Could not load custom prompt templates, falling back to instructions: {e}")
            # Fallback to instructions parameter if PromptTemplates fails
            self.writing_agent = CodeAgent(
                model=writing_model,
                tools=[
                    add_textbox,
                    replace_textbox_content,
                    modify_text_in_textbox,
                    add_text_to_textbox,
                    format_textbox_style,
                    move_object,
                    resize_object,
                    position_and_resize_object,
                    copy_object_to_slide,
                    duplicate_object_on_same_slide,
                    delete_object
                ],
                instructions=writing_agent_instructions,
                name="writing_agent", 
                description="Executes PowerPoint modifications using specialized tools",
                max_steps=5,
                verbosity_level=LogLevel.DEBUG
            )
        
        # Create Manager Agent with strategic tools and managed agents
        self.manager_agent = CodeAgent(
            model=manager_model,
            tools=[
                get_current_slide_context_tool,
                get_object_properties,
                get_annotated_slide_image_tool
            ],
            managed_agents=[self.vision_agent, self.writing_agent],
            instructions=manager_agent_instructions,
            max_steps=4,
            verbosity_level=LogLevel.DEBUG
            # Removed planning_interval=0 to prevent modulo by zero error
        )
    
    def process_request(self, user_message: str) -> dict:
        """
        Process a user request through the multi-agent system.
        
        Args:
            user_message: The user's request
            
        Returns:
            dict: Contains 'answer', 'generated_code', and 'slide_context'
        """
        with trace_tool_call("multiagent_request", user_message=user_message[:100]):
                try:
                    add_trace_event("multiagent_start", user_message=user_message)
                    
                    # Get current slide context
                    slide_context = get_current_slide_context(force_refresh=True)
                    
                    # Enhanced message with slide context
                    enhanced_message = f"""CURRENT SLIDE CONTEXT:
    {slide_context}

    USER REQUEST: {user_message}

    INSTRUCTIONS:
    1. Analyze the user request and current slide state
    2. Determine if visual analysis is needed (call Vision Agent with slide image)
    3. Execute any required PowerPoint modifications through the Writing Agent
    4. Provide a comprehensive response to the user

    Remember: You have access to get_current_slide_context_tool() and get_object_properties() for context gathering.
    The Vision Agent can analyze slide visuals, and the Writing Agent can execute all PowerPoint operations.
    """
                    
                    add_trace_event("manager_execution", has_images=False)
                    
                    # Run the manager agent (no images by default - manager must explicitly get image if needed)
                    answer = self.manager_agent.run(enhanced_message)
                    
                    # Get updated slide context after operations
                    updated_context = get_current_slide_context(force_refresh=True)
                    
                    add_trace_event("multiagent_completed", success=True)
                    
                    return {
                        'answer': answer,
                        'generated_code': "# Multi-agent system executed successfully\n# Operations handled by specialized agents",
                        'slide_context': updated_context
                    }
                    
                except Exception as e:
                    add_trace_event("multiagent_error", error=str(e))
                    return {
                        'answer': f"Error in multi-agent system: {str(e)}",
                        'generated_code': f"# Error: {str(e)}",
                        'slide_context': "Error reading slide context"
                    }

# Global multi-agent system instance
multiagent_system = None

def get_multiagent_system():
    """Get or create the global multi-agent system instance."""
    global multiagent_system
    if multiagent_system is None:
        multiagent_system = MultiAgentPPTSystem()
    return multiagent_system

def run_multiagent_request(message: str) -> dict:
    """
    Run a request through the multi-agent PowerPoint system.
    
    Args:
        message: The user's message/request
        
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context'
    """
    system = get_multiagent_system()
    return system.process_request(message)

# Compatibility function for existing GUI
def run_agent_with_code_capture(message: str, images=None) -> dict:
    """
    Compatibility function that routes requests to the multi-agent system.
    This maintains compatibility with the existing GUI while using the new architecture.
    
    Args:
        message: The user's message/request
        images: Legacy parameter (images are now handled automatically)
        
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context'
    """
    # Note: images parameter is ignored as the multi-agent system handles vision automatically
    return run_multiagent_request(message)
