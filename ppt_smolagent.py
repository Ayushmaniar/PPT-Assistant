from smolagents import CodeAgent, tool, OpenAIServerModel
from smolagents.monitoring import LogLevel
import os
import re
from dotenv import load_dotenv
import logging
import io
import sys
import win32com.client
import pythoncom

from lightning_slide_context_reader import LightningFastPowerPointSlideReader as PowerPointSlideReader

# Load environment variables from .env file
load_dotenv()

# Initialize Phoenix tracing
from phoenix_config import initialize_phoenix, trace_tool_call, add_trace_event, trace_function
phoenix_initialized = initialize_phoenix()
if phoenix_initialized:
    print("✅ Phoenix tracing initialized successfully")
else:
    print("⚠️  Phoenix tracing disabled (missing PHOENIX_API_KEY)")

# Set the OpenAI API key from environment
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found in environment variables. Please check your .env file.")

# Define the model using OpenAIServerModel
model = OpenAIServerModel(
    model_id="gpt-4.1",
    api_key=openai_api_key,
    api_base = "https://api.openai.com/v1"
)

# Import HTML processing functions
from html_processor import parse_html_text, process_html_lists, apply_html_formatting

# Tool to add a textbox to a PowerPoint slide
@tool
def add_textbox(slide_idx: int = 1, html_text: str = "<b>Sample Text</b>", left: int = 100, top: int = 100, width: int = 400, height: int = 50, font_size: int = None, font_name: str = None, text_align: str = "left") -> str:
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
    # Trace the tool call
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
            # Process HTML (always enabled now)
            # First process lists and headers
            processed_text, list_info = process_html_lists(html_text)
            
            # Then process inline formatting
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
                        # Calculate line position in the text
                        lines = plain_text.split('\n')
                        if info['line'] < len(lines):
                            line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                            line_length = len(lines[info['line']])
                            
                            if line_length > 0:
                                header_range = text_range.Characters(line_start, line_length)
                                
                                # Apply header formatting based on level
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
            
            # Apply global font settings (font_name and base font_size for non-headers)
            if font_name:
                text_range.Font.Name = font_name
            
            # Set text alignment
            alignment_map = {
                "left": 1,
                "center": 2, 
                "right": 3
            }
            
            if text_align.lower() in alignment_map:
                text_range.ParagraphFormat.Alignment = alignment_map[text_align.lower()]
            
            # Clear slide context cache to ensure fresh context on next request
            try:
                from slide_context_reader import PowerPointSlideReader
                reader = get_slide_reader()
                if reader:
                    reader.clear_context_cache()
            except Exception as e:
                pass  # Silently continue if cache clearing fails
            
            add_trace_event("textbox_completed", success=True, text_length=len(plain_text))
            return f"Textbox added to slide {slide_idx} with HTML formatting: {plain_text[:50]}{'...' if len(plain_text) > 50 else ''}"
            
        except Exception as e:
            add_trace_event("textbox_error", error=str(e), error_type=type(e).__name__)
            return f"Error adding textbox: {str(e)}"

@tool
def replace_textbox_content(id: int, html_text: str, font_size: int = None, font_name: str = None, text_align: str = None) -> str:
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
        font_size: Base font size in points (headers will be larger)
        font_name: Font name for the text
        text_align: Text alignment - "left", "center", "right", or "justify"
    
    Returns:
        str: Confirmation message with details of what was updated
    """
    return _update_textbox_internal(
        id=id,
        html_text=html_text,
        text_operation="replace",
        font_size=font_size,
        font_name=font_name,
        text_align=text_align
    )

@tool
def modify_text_in_textbox(id: int, find_pattern: str, replacement_text: str, regex_flags: str = "IGNORECASE") -> str:
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
        regex_flags: Regex flags like "IGNORECASE" (default: "IGNORECASE")
    
    Returns:
        str: Confirmation message with details of what was replaced
    """
    return _update_textbox_internal(
        id=id,
        regex_finder=find_pattern,
        replacement_text=replacement_text,
        regex_flags=regex_flags
    )

@tool
def add_text_to_textbox(id: int, html_text: str, position: str = "end") -> str:
    """
    Add new text to the beginning or end of existing textbox content.
    
    This tool preserves all existing text and adds new content before or after it.
    
    Args:
        id: The ID of the textbox to modify
        html_text: HTML-formatted text to add
        position: Where to add the text - "start" (beginning) or "end" (default)
    
    Returns:
        str: Confirmation message with details of what was added
    """
    operation = "prepend" if position == "start" else "append"
    return _update_textbox_internal(
        id=id,
        html_text=html_text,
        text_operation=operation
    )

@tool
def format_textbox_style(id: int, font_size: int = None, font_name: str = None, text_align: str = None, 
                        line_spacing: float = None, left_margin: float = None, right_margin: float = None, 
                        top_margin: float = None, bottom_margin: float = None) -> str:
    """
    Change the formatting and layout properties of a textbox without modifying text content.
    
    Use this to adjust visual appearance like font, alignment, spacing, and margins.
    
    Args:
        id: The ID of the textbox to format
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
        font_size=font_size,
        font_name=font_name,
        text_align=text_align,
        line_spacing=line_spacing,
        left_margin=left_margin,
        right_margin=right_margin,
        top_margin=top_margin,
        bottom_margin=bottom_margin
    )

def _update_textbox_internal(id: int, html_text: str = None, text_operation: str = "replace", regex_finder: str = None, replacement_text: str = None, regex_flags: str = "IGNORECASE", font_size: int = None, font_name: str = None, text_align: str = None, line_spacing: float = None, left_margin: float = None, right_margin: float = None, top_margin: float = None, bottom_margin: float = None) -> str:
    """
    Internal implementation for textbox updates. Do not call directly.
    """
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
        
        # Find the textbox by ID
        target_shape = None
        target_slide = None
        
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
                            # Calculate line position in the text
                            lines = plain_text.split('\n')
                            if info['line'] < len(lines):
                                line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                                line_length = len(lines[info['line']])
                                
                                if line_length > 0:
                                    header_range = target_shape.TextFrame.TextRange.Characters(line_start, line_length)
                                    
                                    # Apply header formatting based on level
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
                # For append/prepend, we need to process the combined text to apply HTML formatting
                combined_text = current_text + html_text
                
                # Process the combined HTML text
                processed_text, list_info = process_html_lists(combined_text)
                plain_text, format_segments = parse_html_text(processed_text)
                apply_html_formatting(target_shape.TextFrame.TextRange, plain_text, format_segments)
                
                # Apply header formatting if any headers are present
                for info in list_info:
                    if info['type'] == 'header':
                        try:
                            # Calculate line position in the text
                            lines = plain_text.split('\n')
                            if info['line'] < len(lines):
                                line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                                line_length = len(lines[info['line']])
                                
                                if line_length > 0:
                                    header_range = target_shape.TextFrame.TextRange.Characters(line_start, line_length)
                                    
                                    # Apply header formatting based on level
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
                # For prepend, we need to process the combined text to apply HTML formatting
                combined_text = html_text + current_text
                
                # Process the combined HTML text
                processed_text, list_info = process_html_lists(combined_text)
                plain_text, format_segments = parse_html_text(processed_text)
                apply_html_formatting(target_shape.TextFrame.TextRange, plain_text, format_segments)
                
                # Apply header formatting if any headers are present
                for info in list_info:
                    if info['type'] == 'header':
                        try:
                            # Calculate line position in the text
                            lines = plain_text.split('\n')
                            if info['line'] < len(lines):
                                line_start = sum(len(lines[i]) + 1 for i in range(info['line'])) + 1
                                line_length = len(lines[info['line']])
                                
                                if line_length > 0:
                                    header_range = target_shape.TextFrame.TextRange.Characters(line_start, line_length)
                                    
                                    # Apply header formatting based on level
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
                # Find all matches in the original text
                matches = list(re.finditer(regex_finder, current_text, flags))
                
                if matches:
                    if replacement_text is not None:
                        # Check if replacement contains HTML formatting
                        if any(marker in replacement_text for marker in ['<b>', '<i>', '<u>', '<s>', '<span', '<strong>', '<em>']):
                            # Process HTML in replacement text to get clean text and formatting
                            processed_replacement, _ = process_html_lists(replacement_text)
                            plain_replacement, format_segments = parse_html_text(processed_replacement)
                            
                            # CRITICAL FIX: Instead of replacing all text at once, replace each match individually
                            # This preserves existing formatting that was applied by previous calls
                            
                            # Process matches in reverse order to maintain position indices
                            for match in reversed(matches):
                                match_start = match.start()
                                match_end = match.end()
                                match_length = match_end - match_start
                                
                                # Replace this specific match in the textbox without affecting the rest
                                if match_length > 0:
                                    # Get the character range for this match (1-based indexing in PowerPoint)
                                    match_range = target_shape.TextFrame.TextRange.Characters(match_start + 1, match_length)
                                    
                                    # Replace the text in this range only
                                    match_range.Text = plain_replacement
                                    
                                    # Now apply formatting to the replacement text
                                    replacement_start_pos = match_start + 1  # 1-based for PowerPoint
                                    
                                    for segment in format_segments:
                                        try:
                                            # Calculate absolute position within the replacement
                                            # segment['start'] is 1-based relative to replacement start
                                            absolute_start = replacement_start_pos + segment['start'] - 1
                                            segment_length = segment['length']
                                            
                                            if segment_length > 0:
                                                # Get the character range for this formatting segment
                                                char_range = target_shape.TextFrame.TextRange.Characters(absolute_start, segment_length)
                                                
                                                # Apply the specific formatting from this segment
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
                                    # This is needed because we're processing in reverse order
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
            
            # Apply paragraph formatting (these don't conflict with markdown)
            if text_align is not None:
                alignment_map = {
                    "left": 1,
                    "center": 2,
                    "right": 3,
                    "justify": 4
                }
                if text_align.lower() in alignment_map:
                    text_range.ParagraphFormat.Alignment = alignment_map[text_align.lower()]
                    updates_made.append(f"set text alignment to {text_align}")
            
            if line_spacing is not None:
                text_range.ParagraphFormat.LineRuleWithin = 1  # Multiple line spacing
                text_range.ParagraphFormat.SpaceWithin = line_spacing
                updates_made.append(f"set line spacing to {line_spacing}")
        
        # Apply text margins (only to entire textbox)
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
            from slide_context_reader import PowerPointSlideReader
            reader = get_slide_reader()
            if reader:
                reader.clear_context_cache()
        except Exception as e:
            pass  # Silently continue if cache clearing fails
        
        if updates_made:
            return f"Updated textbox {id} on slide {target_slide.SlideIndex}: {'; '.join(updates_made)}"
        else:
            return f"No updates specified for textbox {id}"
    
    except Exception as e:
        return f"Error updating textbox {id}: {str(e)}"

# Universal object manipulation tools

@tool
def move_object(id: int, left: int, top: int) -> str:
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
    
    Returns:
        str: Confirmation message with the object's new position
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
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
def resize_object(id: int, width: int, height: int) -> str:
    """
    Change the size of any object (textbox, shape, image, etc.) to new dimensions.
    
    Args:
        id: The ID of the object to resize
        width: New width in points
        height: New height in points
    
    Returns:
        str: Confirmation message with the object's new dimensions
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
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
def position_and_resize_object(id: int, left: int, top: int, width: int, height: int) -> str:
    """
    Move and resize an object in a single operation for precise positioning.
    
    Useful when you need to set both position and size to avoid multiple operations.
    
    Args:
        id: The ID of the object to position and resize
        left: Distance from left edge of slide in points
        top: Distance from top edge of slide in points
        width: New width in points
        height: New height in points
    
    Returns:
        str: Confirmation message with the object's new position and size
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
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
def get_object_properties(id: int) -> dict:
    """
    Get detailed information about any object on the slide.
    
    Returns comprehensive details including position, size, type, and content information.
    Use this to inspect objects before modifying them.

    Args:
        id: The ID of the object to inspect

    Returns:
        dict: Object properties including slide, position, size, type, and content details
    """
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
                        props["text_content"] = shape.TextFrame.TextRange.Text[:100] + "..." if len(shape.TextFrame.TextRange.Text) > 100 else shape.TextFrame.TextRange.Text
                    
                    return props
        return {"error": f"Object with ID {id} not found"}
    except Exception as e:
        return {"error": f"Error inspecting object {id}: {str(e)}"}

def _get_shape_type_name(shape_type: int) -> str:
    """Convert PowerPoint shape type number to readable name."""
    type_map = {
        1: "AutoShape",
        5: "Freeform", 
        9: "Group",
        11: "Picture",
        12: "OLEObject",
        13: "Chart",
        14: "Table",
        15: "Media",
        17: "TextBox",
        18: "Content",
        19: "SmartArt"
    }
    return type_map.get(shape_type, f"Unknown({shape_type})")

@tool
def copy_object_to_slide(id: int, target_slide_idx: int, new_left: int = None, new_top: int = None) -> int:
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
def duplicate_object_on_same_slide(id: int, offset_left: int = 20, offset_top: int = 20) -> int:
    """
    Create a duplicate of an object on the same slide with a slight position offset.
    
    Useful for creating multiple similar objects quickly.
    
    Args:
        id: The ID of the object to duplicate
        offset_left: How many points to move the duplicate to the right (default: 20)
        offset_top: How many points to move the duplicate down (default: 20)
    
    Returns:
        int: The ID of the newly created duplicate, or -1 if operation failed
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
def delete_object(id: int) -> str:
    """
    Permanently delete an object from the slide.
    
    ⚠️ WARNING: This action cannot be undone programmatically.
    
    Args:
        id: The ID of the object to delete
    
    Returns:
        str: Confirmation message of deletion
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
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

# The tool is automatically registered when using the @tool decorator

instructions = """
You are a highly capable AI assistant that automates Microsoft PowerPoint presentations using specialized tools.

IMPORTANT: You will ALWAYS receive current slide context before user requests. This context contains:
- Current slide number and layout
- All objects/shapes with their positions, sizes, text content, and formatting
- Object IDs (permanent identifiers for reliable reference)
- Animations and slide notes

USE THIS CONTEXT in your THOUGHT process to make informed decisions about positioning, styling, and content placement.

📝 TEXT EDITING TOOLS - Choose the RIGHT tool for the task:

- NOTE : NEVER USE ANY EMOTICONS OR EMOJIS.

1. **replace_textbox_content(id, html_text)** 
   - COMPLETELY REPLACES all text in a textbox
   - Use when: User wants to change entire content
   - Example: "Change the title to 'New Title'"

2. **modify_text_in_textbox(id, find_pattern, replacement_text)**
   - FINDS and REPLACES specific words/phrases only
   - PRESERVES all other existing text
   - Use when: User wants to modify specific words
   - Example: "Make 'Company Name' bold" or "Change 'red' to 'blue'"

3. **add_text_to_textbox(id, html_text, position)**
   - ADDS text to beginning ("start") or end ("end") of existing content
   - Use when: User wants to append/prepend text
   - Example: "Add 'Confidential' to the end"

4. **format_textbox_style(id, font_size, font_name, text_align, etc.)**
   - Changes visual formatting WITHOUT modifying text content
   - Use when: User wants to change appearance only
   - Example: "Make the text center-aligned" or "Change font to Arial"

🎨 HTML FORMATTING SYNTAX:
- <b>bold</b> or <strong>bold</strong> - Bold text
- <i>italic</i> or <em>italic</em> - Italic text  
- <s>strikethrough</s> or <del>strikethrough</del> - Strikethrough text
- <u>underlined</u> - Underlined text
- <span style="color: red">colored</span> - Colored text (hex #FF0000 or color names)
- <span style="background-color: yellow">highlighted</span> - Background color
- <h1>Header 1</h1>, <h2>Header 2</h2>, <h3>Header 3</h3> - Headers (auto-sized)
- <ul><li>bullet point</li></ul> - Bullet lists
- <ol><li>numbered item</li></ol> - Numbered lists

📐 POSITIONING TOOLS:

1. **move_object(id, left, top)** - Move object to new position
2. **resize_object(id, width, height)** - Change object size  
3. **position_and_resize_object(id, left, top, width, height)** - Move and resize in one operation

📋 OBJECT MANAGEMENT:

1. **get_object_properties(id)** - Inspect object details before modifying
2. **copy_object_to_slide(id, target_slide, new_left, new_top)** - Copy to another slide
3. **duplicate_object_on_same_slide(id, offset_left, offset_top)** - Duplicate with offset
4. **delete_object(id)** - Permanently remove object

📏 SLIDE COORDINATE SYSTEM:
- Origin (0,0) = top-left corner
- Standard slide: 960 points wide × 540 points tall  
- Measurements in points (72 points = 1 inch)

⚠️ CRITICAL RULES:
- ALWAYS use object IDs from slide context for reliable reference
- Choose the most specific tool for each task
- Consider existing content positioning when adding new elements
- Match existing fonts/styles when appropriate for consistency
- **ALWAYS end with final_answer() to report completion status and summary**
- **LEVERAGE MULTI-TOOL ACTIONS**: Use multiple tools together when they accomplish related goals efficiently

🏁 TASK COMPLETION:
When you have successfully completed all requested tasks:
1. Provide a final THOUGHT summarizing what was accomplished
2. Use final_answer() tool to report the completion status
3. Include a clear summary of what was changed/added/modified
4. Example: final_answer("Successfully added title and formatted text. All requested changes have been completed.")

Remember: Only modify slides when the user specifically requests changes. For informational questions about slide content, respond without using tools but still follow the THOUGHT→OBSERVATION pattern.
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

@tool
def get_current_slide_context_html(force_refresh: bool = False) -> str:
    """
    Get the current PowerPoint slide context with ALL text content in HTML format.
    
    *** IMPORTANT: This tool returns text content in HTML format ***
    
    The returned context includes HTML tags for formatting:
    - <b>bold text</b> for bold formatting
    - <i>italic text</i> for italic formatting
    - <u>underlined text</u> for underlined formatting  
    - <s>strikethrough text</s> for strikethrough formatting
    - <span style="color: #RRGGBB">colored text</span> for colored text
    
    Use this context to understand the current slide content and formatting.
    When modifying text, you can use these same HTML tags in your tool calls.
    
    Args:
        force_refresh: If True, force a fresh read of the slide (ignore cache)
    
    Returns:
        str: The current slide context with HTML-formatted text content
    """
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

def get_slide_reader():
    """Get or create the global slide reader instance."""
    global slide_reader
    if slide_reader is None:
        try:
            slide_reader = PowerPointSlideReader()
            print("🚀 Slide reader initialized with original HTML conversion")
        except Exception as e:
            print(f"Warning: Could not initialize slide reader: {e}")
            slide_reader = None
    return slide_reader

def get_current_slide_context(force_refresh=False):
    """Get the current slide context as a string."""
    try:
        reader = get_slide_reader()
        if reader and reader.ppt_app:
            # Force refresh of context by clearing cached values
            # This ensures we always get the latest slide when user switches
            if force_refresh:
                context = reader.force_refresh_context()
            else:
                context = reader.get_current_context()
            return context if context else "No slide context available"
        else:
            return "PowerPoint not connected - no slide context available"
    except Exception as e:
        return f"Error reading slide context: {e}"

def get_fresh_slide_context():
    """Get a completely fresh slide context, ignoring any cache."""
    return get_current_slide_context(force_refresh=True)

def clear_slide_context_cache():
    """Clear the slide context cache to force refresh on next access."""
    try:
        reader = get_slide_reader()
        if reader:
            reader.clear_context_cache()
            print("🗑️ Slide context cache cleared")
    except Exception as e:
        print(f"⚠️ Warning: Could not clear context cache: {e}")

agent = CodeAgent(
    tools=[
        add_textbox,
        replace_textbox_content,
        modify_text_in_textbox,
        add_text_to_textbox,
        format_textbox_style,
        move_object,
        resize_object,
        position_and_resize_object,
        get_object_properties,
        copy_object_to_slide,
        duplicate_object_on_same_slide,
        delete_object
    ],
    instructions=instructions,
    max_steps=2,
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
    # Trace the entire agent interaction
    with trace_tool_call("agent_interaction", user_message=message[:100]):
        try:
            add_trace_event("agent_start", user_message=message)
            
            # Get current slide context
            add_trace_event("context_retrieval", action="getting_slide_context")
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
                add_trace_event("agent_execution", action="running_smolagent", enhanced_message_length=len(enhanced_message))
                answer = agent.run(enhanced_message)
                add_trace_event("agent_response", answer_length=len(answer) if answer else 0)
                
            finally:
                # Restore stdout/stderr
                sys.stdout = stdout_backup
                sys.stderr = stderr_backup
                logger.removeHandler(code_capture_handler)
            
            # Get captured outputs and clean them
            stdout_content = strip_ansi_codes(stdout_capture.getvalue())
            stderr_content = strip_ansi_codes(stderr_capture.getvalue())
            captured_code = strip_ansi_codes(code_capture_handler.get_code())
            
            # IMPORTANT: Force refresh the slide context after agent execution
            # This ensures that any objects added/deleted by the agent are reflected in the context
            try:
                add_trace_event("context_refresh", action="refreshing_slide_context")
                reader = get_slide_reader()
                if reader and reader.ppt_app:
                    # Force refresh the context to reflect any changes made by the agent
                    updated_context = reader.force_refresh_context()
                    print("✅ Slide context refreshed after agent execution")
                else:
                    updated_context = slide_context
            except Exception as e:
                print(f"⚠️ Warning: Could not refresh context after execution: {e}")
                updated_context = slide_context
            
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
                if "textbox" in message.lower() or "add" in message.lower():
                    tool_name = "add_textbox_tool"
                else:
                    tool_name = "PowerPoint automation tool"
                    
                generated_code = f"""# Agent Execution Summary
# Request: "{message}"
# 
# The agent executed your request using the {tool_name}().
# This is a direct tool call that doesn't require custom code generation.
#
# The operation was completed successfully using the built-in PowerPoint COM interface.
# 
# Available tools:
# - add_textbox_tool: Create textboxes with formatting options
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
            
            add_trace_event("agent_completed", 
                success=True, 
                answer_length=len(clean_answer),
                code_generated=bool(generated_code.strip()),
                context_updated=bool(updated_context != slide_context)
            )
            
            return {
                'answer': clean_answer,
                'generated_code': generated_code,
                'slide_context': updated_context,
                'debug_output': f"STDOUT:\n{stdout_content}\n\nSTDERR:\n{stderr_content}"
            }
            
        except Exception as e:
            add_trace_event("agent_error", error=str(e), error_type=type(e).__name__)
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


