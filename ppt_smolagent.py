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
    model_id="gpt-4o",
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

# Tool replace_textbox_content removed (functionality covered by update_textbox with operation="replace")

# Tool modify_text_in_textbox removed (functionality covered by update_textbox with operation="find_replace")

# Tool add_text_to_textbox removed (functionality covered by update_textbox with operation="append"/"prepend")

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
def update_textbox(id: int, html_text: str = None, operation: str = "replace", find_pattern: str = None, replacement_text: str = None, regex_flags: str = "IGNORECASE", font_size: int = None, font_name: str = None, text_align: str = None) -> str:
    """
    Versatile tool to update text in a textbox with multiple operation modes.
    
    Operation modes:
    - "replace": Replace ALL text with new HTML-formatted content (default)
    - "append": Add new text at the END of existing text
    - "prepend": Add new text at the BEGINNING of existing text
    - "find_replace": Find and replace specific text patterns using regex
    
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
        html_text: HTML-formatted text content (used for replace/append/prepend)
        operation: How to apply the text - "replace", "append", "prepend", or "find_replace"
        find_pattern: Text pattern to find (only used with operation="find_replace")
        replacement_text: HTML text to replace matches with (only used with operation="find_replace")
        regex_flags: Regex flags like "IGNORECASE" (default: "IGNORECASE")
        font_size: Base font size in points (headers will be larger)
        font_name: Font name for the text
        text_align: Text alignment - "left", "center", "right", or "justify"
    
    Returns:
        str: Confirmation message with details of what was updated
    """
    # Ensure operation is lowercase
    operation = operation.lower()
    
    # Map operations to internal text_operation values
    text_operation_map = {
        "replace": "replace",
        "append": "append",
        "prepend": "prepend",
        "find_replace": "regex"
    }
    
    # Set text_operation to the mapped value, defaulting to "replace"
    text_operation = text_operation_map.get(operation, "replace")
    
    # Call the internal implementation based on operation type
    if operation == "find_replace":
        if not find_pattern:
            return "ERROR: When using operation='find_replace', you must provide the find_pattern parameter"
        if not replacement_text:
            return "ERROR: When using operation='find_replace', you must provide the replacement_text parameter"
            
        return _update_textbox_internal(
            id=id,
            regex_finder=find_pattern,
            replacement_text=replacement_text,
            regex_flags=regex_flags,
            font_size=font_size,
            font_name=font_name,
            text_align=text_align
        )
    else:
        # For replace, append, prepend operations
        if not html_text:
            return f"ERROR: When using operation='{operation}', you must provide the html_text parameter"
            
        return _update_textbox_internal(
            id=id,
            html_text=html_text,
            text_operation=text_operation,
            font_size=font_size,
            font_name=font_name,
            text_align=text_align
        )

@tool
def position_object(id: int, left: int = None, top: int = None, width: int = None, height: int = None) -> str:
    """
    Position and/or resize any object on the slide with flexible options.
    
    This is a versatile tool that can:
    1. Move an object (provide only left/top)
    2. Resize an object (provide only width/height)
    3. Both move and resize (provide all parameters)
    
    The slide coordinate system:
    - Origin (0, 0) is at the top-left corner
    - Standard slide is 960 points wide × 540 points tall
    - Measurements are in points (72 points = 1 inch)
    
    Args:
        id: The ID of the object to modify
        left: Optional - New distance from left edge of slide in points
        top: Optional - New distance from top edge of slide in points
        width: Optional - New width in points
        height: Optional - New height in points
    
    Returns:
        str: Confirmation message with the object's new position and/or size
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    changes = []
                    
                    # Apply position changes if specified
                    if left is not None:
                        shape.Left = left
                        changes.append(f"left to {left}")
                    if top is not None:
                        shape.Top = top
                        changes.append(f"top to {top}")
                        
                    # Apply size changes if specified
                    if width is not None:
                        shape.Width = width
                        changes.append(f"width to {width}")
                    if height is not None:
                        shape.Height = height
                        changes.append(f"height to {height}")
                    
                    # Build appropriate response based on what changed
                    if not changes:
                        return f"No changes specified for object {id}"
                    else:
                        action = "Positioned" if (left is not None or top is not None) else ""
                        action = "Resized" if action == "" and (width is not None or height is not None) else action
                        action = "Positioned and resized" if (left is not None or top is not None) and (width is not None or height is not None) else action
                        return f"{action} object {id} on slide {slide.SlideIndex}: set " + ", ".join(changes)
                        
        return f"Object with ID {id} not found"
    except Exception as e:
        return f"Error modifying object {id}: {str(e)}"

@tool
def duplicate_object(id: int, target_slide_idx: int = None, left: int = None, top: int = None, offset_left: int = 20, offset_top: int = 20) -> int:
    """
    Create a duplicate of an object on the same or different slide.
    
    This versatile tool can:
    1. Duplicate on same slide with slight offset (default behavior)
    2. Copy to another slide at original position (specify target_slide_idx only)
    3. Copy to another slide at new position (specify target_slide_idx and left/top)
    4. Duplicate on same slide at specific position (specify left/top)
    
    Args:
        id: The ID of the object to duplicate
        target_slide_idx: Optional slide number to copy to (if None, uses current slide)
        left: Optional specific left position for the copy
        top: Optional specific top position for the copy
        offset_left: How many points to offset if left not specified (default: 20)
        offset_top: How many points to offset if top not specified (default: 20)
    
    Returns:
        int: The ID of the newly created duplicate, or -1 if operation failed
    """
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # Find source object and its slide
        source_shape = None
        source_slide = None
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.Id == id:
                    source_shape = shape
                    source_slide = slide
                    break
            if source_shape:
                break
        
        if not source_shape:
            return -1
            
        # Determine if we're duplicating on same slide or copying to a different slide
        if target_slide_idx is None or target_slide_idx == source_slide.SlideIndex:
            # Same-slide duplication
            dup = source_shape.Duplicate()
            if dup and dup.Count > 0:
                new_shape = dup[0]
                
                # Set position with either specific coordinates or offset
                if left is not None:
                    new_shape.Left = left
                else:
                    new_shape.Left = source_shape.Left + offset_left
                    
                if top is not None:
                    new_shape.Top = top
                else:
                    new_shape.Top = source_shape.Top + offset_top
                    
                return new_shape.Id
            else:
                return -1
        else:
            # Cross-slide copying
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
                if left is not None:
                    new_shape.Left = left
                if top is not None:
                    new_shape.Top = top
                
                return new_id
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

# The tool is automatically registered when using the @tool decorator

instructions = """
You are a highly capable AI assistant that automates Microsoft PowerPoint presentations using specialized tools.

IMPORTANT: You will ALWAYS receive current slide context before user requests. This context contains:
- Current slide number and layout
- All objects/shapes with their positions, sizes, text content, and formatting
- Object IDs (permanent identifiers for reliable reference)
- Animations and slide notes

🔍 VISION CAPABILITIES:
When an image of the slide is provided, you can:
- See object bounding boxes (green rectangles) with ID labels (yellow "ID:29" tags)
- Understand spatial relationships and visual layout
- Make better positioning decisions for new objects
- Identify visual patterns and alignment

The image provides visual context to enhance your tool usage decisions.

USE THIS CONTEXT in YOUR THOUGHT process to make informed decisions about positioning, styling, and content placement.

📝 TEXT EDITING TOOLS - Choose the RIGHT tool for the task:

- NOTE : NEVER USE ANY EMOTICONS OR EMOJIS.

📏 SLIDE COORDINATE SYSTEM:
- Origin (0,0) = top-left corner
- Standard slide: 960 points wide × 540 points tall  
- Measurements in points (72 points = 1 inch)

⚠️ CRITICAL RULES:
- ALWAYS use object IDs from slide context for reliable reference
- When an image is provided, use visual cues to enhance your understanding
- Choose the most specific tool for each task
- Consider existing content positioning when adding new elements
- Match existing fonts/styles when appropriate for consistency
- **LEVERAGE MULTI-TOOL ACTIONS**: Use multiple tools together when they accomplish related goals efficiently

Remember: Only modify slides when the user specifically requests changes.
Remember: If the user asks a query, then you need to reply by using the final_answer tool in the code, ALWAYS answer queries using final_answer tool.
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

agent = CodeAgent(
    tools=[
        add_textbox,
        update_textbox,
        format_textbox_style,
        position_object,
        duplicate_object,
        get_object_properties,
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

def _legacy_run_agent_with_code_capture(message, images=None):
    """
    Legacy single-agent implementation for backwards compatibility.
    Automatically includes current slide context in the message.
    
    Args:
        message (str): The user's message/request
        images (list[PIL.Image.Image], optional): List of PIL Image objects to pass to the agent
    
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context' keys
    """
    # Set up logging to capture the agent's output
    code_capture_handler.clear()
    logger = logging.getLogger()
    logger.addHandler(code_capture_handler)
    
    # Get current slide context
    slide_context = get_current_slide_context()
    
    # Enhance the message with slide context
    enhanced_message = f"""CURRENT SLIDE CONTEXT:
{slide_context}

USER REQUEST:
{message}
"""
    
    # Run the agent
    answer = agent.run(enhanced_message, images=images)
    
    # Clean the answer
    clean_answer = strip_ansi_codes(answer) if answer else "Operation completed"
    
    # Get the captured code
    captured_code = strip_ansi_codes(code_capture_handler.get_code())
    
    # Remove the handler to avoid duplicated logs
    logger.removeHandler(code_capture_handler)
    
    return {
        'answer': clean_answer,
        'generated_code': captured_code,
        'slide_context': slide_context
    }

def run_agent_with_code_capture(message, images=None):
    """
    Run the agent and capture both the final answer and generated code.
    This function now routes to the new multi-agent system for enhanced capabilities.
    
    Args:
        message (str): The user's message/request
        images (list[PIL.Image.Image], optional): Legacy parameter (handled automatically by multi-agent system)
    
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context' keys
    """
    # Import the new multi-agent system
    try:
        from multiagent_ppt_system import run_multiagent_request
        print("🤖 Using new multi-agent PowerPoint system...")
        return run_multiagent_request(message)
    except ImportError:
        print("⚠️ Multi-agent system not available, falling back to legacy single-agent...")
        # Fallback to legacy system if multi-agent is not available
        return _legacy_run_agent_with_code_capture(message, images)

# Add a simple CLI interface
if __name__ == "__main__":
    import sys
    
    # If command line arguments are provided, use them as the query
    if len(sys.argv) > 1:
        # Join all arguments to form the query
        query = " ".join(sys.argv[1:])
        print(f"Running agent with query: {query}")
        result = run_agent_with_code_capture(query)
        print("\nANSWER:")
        print(result['answer'])
        
        # Optionally show generated code with --verbose flag
        if "--verbose" in sys.argv or "-v" in sys.argv:
            print("\nGENERATED CODE:")
            print(result['generated_code'])
    else:
        print("Usage: uv run python ppt_smolagent.py \"your query here\" [--verbose|-v]")
        print("\nExample: uv run python ppt_smolagent.py \"Add a textbox that says Hello World\"")
        print("Example with code output: uv run python ppt_smolagent.py \"Add a textbox\" --verbose")


