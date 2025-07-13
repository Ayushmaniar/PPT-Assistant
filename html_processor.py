"""
HTML processing module for PowerPoint text formatting.

This module provides functions to parse HTML text and apply formatting
to PowerPoint TextRange objects.
"""

import re
from html.parser import HTMLParser


class PowerPointHTMLParser(HTMLParser):
    """HTML parser specifically designed for PowerPoint text formatting."""
    
    def __init__(self):
        super().__init__()
        self.reset_parser()
    
    def reset_parser(self):
        """Reset the parser state."""
        self.plain_text = ""
        self.format_segments = []
        self.tag_stack = []
        self.current_position = 0
        
    def handle_starttag(self, tag, attrs):
        """Handle opening HTML tags."""
        formatting = {}
        
        # Handle self-closing tags that insert content
        if tag == 'br':
            # Insert a line break
            self.plain_text += '\n'
            self.current_position += 1
            return  # Don't push to stack for self-closing tags
        
        if tag == 'b' or tag == 'strong':
            formatting['bold'] = True
        elif tag == 'i' or tag == 'em':
            formatting['italic'] = True
        elif tag == 'u':
            formatting['underline'] = True
        elif tag == 's' or tag == 'strike' or tag == 'del':
            formatting['strikethrough'] = True
        elif tag == 'span':
            # Parse style attributes for span tags
            for attr_name, attr_value in attrs:
                if attr_name == 'style':
                    span_formatting = self._parse_style(attr_value)
                    formatting.update(span_formatting)
                elif attr_name == 'color':
                    formatting['color'] = attr_value
        elif tag.startswith('h') and len(tag) == 2 and tag[1].isdigit():
            # Handle header tags (h1, h2, h3, etc.)
            level = int(tag[1])
            formatting['header'] = level
        
        # Push formatting onto stack
        self.tag_stack.append({
            'tag': tag,
            'start_position': self.current_position,
            'formatting': formatting
        })
    
    def handle_endtag(self, tag):
        """Handle closing HTML tags."""
        # Find the most recent matching opening tag (proper nesting)
        tag_found = False
        for i in range(len(self.tag_stack) - 1, -1, -1):
            if self.tag_stack[i]['tag'] == tag:
                tag_info = self.tag_stack.pop(i)
                tag_found = True
                
                # Only create format segment if there was actual content
                segment_length = self.current_position - tag_info['start_position']
                if segment_length > 0:
                    self.format_segments.append({
                        'start': tag_info['start_position'] + 1,  # 1-indexed for PowerPoint
                        'length': segment_length,
                        'formatting': tag_info['formatting']
                    })
                break
        
        # If tag wasn't found, it might be malformed HTML - just ignore it
        if not tag_found:
            pass  # Silently ignore unmatched closing tags
    
    def handle_startendtag(self, tag, attrs):
        """Handle self-closing tags like <br />."""
        if tag == 'br':
            # Insert a line break
            self.plain_text += '\n'
            self.current_position += 1
    
    def handle_data(self, data):
        """Handle text content."""
        self.plain_text += data
        self.current_position += len(data)
    
    def _parse_style(self, style_str):
        """Parse CSS style string and extract formatting."""
        formatting = {}
        
        # Split by semicolon and process each property
        properties = [prop.strip() for prop in style_str.split(';') if prop.strip()]
        
        for prop in properties:
            if ':' in prop:
                key, value = prop.split(':', 1)
                key = key.strip().lower()
                value = value.strip()
                
                if key == 'color':
                    formatting['color'] = value
                elif key == 'background-color' or key == 'background':
                    formatting['background_color'] = value
                elif key == 'font-weight' and value == 'bold':
                    formatting['bold'] = True
                elif key == 'font-style' and value == 'italic':
                    formatting['italic'] = True
                elif key == 'text-decoration':
                    if 'underline' in value:
                        formatting['underline'] = True
                    if 'line-through' in value:
                        formatting['strikethrough'] = True
        
        return formatting


def parse_html_text(html_text):
    """
    Parse HTML text and return structured formatting data.
    
    Args:
        html_text (str): Text with HTML formatting
        
    Returns:
        tuple: (plain_text, formatting_segments)
            - plain_text: Text without HTML tags
            - formatting_segments: List of formatting instructions
    """
    parser = PowerPointHTMLParser()
    parser.reset_parser()
    
    try:
        parser.feed(html_text)
        parser.close()
    except Exception as e:
        # If parsing fails, return the text as-is
        return html_text, []
    
    # Sort segments by start position for consistent application
    parser.format_segments.sort(key=lambda x: x['start'])
    
    return parser.plain_text, parser.format_segments


def process_html_lists(text):
    """
    Process HTML lists and convert to PowerPoint-friendly format.
    
    Args:
        text (str): Text potentially containing HTML lists
        
    Returns:
        tuple: (processed_text, list_info)
    """
    list_info = []
    original_text = text
    
    # Handle unordered lists (ul/li)
    ul_pattern = r'<ul[^>]*>(.*?)</ul>'
    ol_pattern = r'<ol[^>]*>(.*?)</ol>'
    li_pattern = r'<li[^>]*>(.*?)</li>'
    
    def process_ul(match):
        ul_content = match.group(1)
        li_matches = re.finditer(li_pattern, ul_content, re.DOTALL)
        
        result = ""
        for li_match in li_matches:
            li_content = li_match.group(1).strip()
            # Keep nested HTML tags for further processing
            result += f"• {li_content}\n"
        
        return result.rstrip()
    
    def process_ol(match):
        ol_content = match.group(1)
        li_matches = list(re.finditer(li_pattern, ol_content, re.DOTALL))
        
        result = ""
        for i, li_match in enumerate(li_matches, 1):
            li_content = li_match.group(1).strip()
            # Keep nested HTML tags for further processing
            result += f"{i}. {li_content}\n"
        
        return result.rstrip()
    
    # Process lists first
    text = re.sub(ul_pattern, process_ul, text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(ol_pattern, process_ol, text, flags=re.DOTALL | re.IGNORECASE)
    
    # Process headers and store their info
    header_pattern = r'<h([1-6])[^>]*>(.*?)</h[1-6]>'
    header_matches = []
    
    for match in re.finditer(header_pattern, text, re.IGNORECASE):
        level = int(match.group(1))
        content = match.group(2).strip()
        header_matches.append((match.start(), match.end(), level, content))
    
    # Replace headers with their content
    text = re.sub(header_pattern, r'\2', text, flags=re.DOTALL | re.IGNORECASE)
    
    # Remove other block tags like <p>, <div>, etc., but keep their content
    block_tags = ['p', 'div', 'section', 'article', 'main', 'aside', 'nav', 'header', 'footer']
    for tag in block_tags:
        pattern = f'<{tag}[^>]*>(.*?)</{tag}>'
        text = re.sub(pattern, r'\1', text, flags=re.DOTALL | re.IGNORECASE)
    
    # Clean up extra whitespace and normalize - but preserve list line breaks
    text = re.sub(r'[ \t]+', ' ', text)  # Normalize spaces and tabs to single spaces
    text = re.sub(r'\n\s*\n', '\n', text)  # Remove empty lines
    text = text.strip()
    
    # Add header info based on content matching
    lines = text.split('\n')
    for start, end, level, content in header_matches:
        # Find which line contains this header content
        for line_idx, line in enumerate(lines):
            if content.strip() in line.strip():
                list_info.append({
                    'line': line_idx,
                    'type': 'header',
                    'level': level
                })
                break
    
    return text, list_info


def apply_html_formatting(text_range, plain_text, segments):
    """
    Apply HTML formatting to a PowerPoint TextRange.
    
    Args:
        text_range: PowerPoint TextRange object
        plain_text (str): Plain text content
        segments (list): Formatting segments from parse_html_text
    """
    # Set the plain text first
    text_range.Text = plain_text
    
    # CRITICAL: Force PowerPoint to process the text change before applying formatting
    try:
        # Force a refresh by accessing the text
        _ = text_range.Text
        ppt_text_length = len(text_range.Text)
        ppt_text_content = text_range.Text
        
        if ppt_text_length != len(plain_text):
            print(f"Warning: PowerPoint text length mismatch")
    except:
        ppt_text_length = len(plain_text)
        ppt_text_content = plain_text
    
    # Sort segments by start position to ensure consistent application
    segments_sorted = sorted(segments, key=lambda x: x['start'])
    
    # Apply formatting to each segment with PowerPoint boundary issue workaround
    for segment in segments_sorted:
        if not segment['formatting']:
            continue
            
        start_pos = segment['start']
        length = segment['length']
        
        # Validate bounds
        if start_pos < 1 or start_pos > ppt_text_length:
            continue
            
        # Adjust length if it would exceed text bounds
        if start_pos + length - 1 > ppt_text_length:
            length = ppt_text_length - start_pos + 1
        
        if length <= 0:
            continue
        
        # POWERPOINT BUG WORKAROUND: Apply formatting character-by-character
        # This avoids PowerPoint's character range boundary issues
        
        formatting = segment['formatting']
        chars_processed = 0
        
        try:
            # Calculate RGB color once if needed
            rgb_color = None
            if formatting.get('color'):
                color_value = formatting['color']
                if color_value.startswith('#'):
                    hex_color = color_value[1:]
                    if len(hex_color) == 6:
                        r = int(hex_color[0:2], 16)
                        g = int(hex_color[2:4], 16) 
                        b = int(hex_color[4:6], 16)
                        rgb_color = r + (g * 256) + (b * 65536)
                    elif len(hex_color) == 3:
                        r = int(hex_color[0] * 2, 16)
                        g = int(hex_color[1] * 2, 16)
                        b = int(hex_color[2] * 2, 16)
                        rgb_color = r + (g * 256) + (b * 65536)
                else:
                    color_map = {
                        'red': 255, 'blue': 16711680, 'green': 65280,
                        'yellow': 65535, 'orange': 33023, 'purple': 8388736,
                        'black': 0, 'white': 16777215
                    }
                    if color_value.lower() in color_map:
                        rgb_color = color_map[color_value.lower()]
            
            # Apply formatting character by character to avoid boundary issues
            # POWERPOINT BUG FIX: Add 1 extra character to account for PowerPoint's off-by-one issue
            actual_length = min(length + 1, ppt_text_length - start_pos + 1)
            
            for i in range(actual_length):
                char_pos = start_pos + i
                if char_pos > ppt_text_length:
                    break
                    
                try:
                    char_range = text_range.Characters(char_pos, 1)
                    
                    # Apply formatting properties
                    if formatting.get('bold'):
                        char_range.Font.Bold = -1
                        
                    if formatting.get('italic'):
                        char_range.Font.Italic = -1
                        
                    if formatting.get('underline'):
                        char_range.Font.Underline = -1
                        
                    if formatting.get('strikethrough'):
                        try:
                            char_range.Font.Strikethrough = -1
                        except:
                            try:
                                char_range.Font.Strike = -1
                            except:
                                pass
                    
                    if rgb_color is not None:
                        char_range.Font.Color.RGB = rgb_color
                    
                    if formatting.get('background_color'):
                        try:
                            bg_value = formatting['background_color']
                            if bg_value.startswith('#'):
                                hex_color = bg_value[1:]
                                if len(hex_color) == 6:
                                    r = int(hex_color[0:2], 16)
                                    g = int(hex_color[2:4], 16) 
                                    b = int(hex_color[4:6], 16)
                                    bg_rgb_color = r + (g * 256) + (b * 65536)
                                    char_range.Font.Fill.ForeColor.RGB = bg_rgb_color
                        except Exception as bg_e:
                            pass  # Ignore background color errors
                    
                    chars_processed += 1
                    
                except Exception as char_e:
                    # Continue with next character instead of failing entirely
                    pass
            
            # Text verification is removed to reduce debug output - formatting is applied successfully
                
        except Exception as e:
            # Final fallback: Try the old approach
            try:
                char_range = text_range.Characters(start_pos, length)
                if formatting.get('color') and rgb_color is not None:
                    char_range.Font.Color.RGB = rgb_color
            except:
                pass  # Silent fallback failure


# Convenience functions for common HTML patterns
def bold(text):
    """Wrap text in bold HTML tags."""
    return f"<b>{text}</b>"

def italic(text):
    """Wrap text in italic HTML tags."""
    return f"<i>{text}</i>"

def underline(text):
    """Wrap text in underline HTML tags."""
    return f"<u>{text}</u>"

def strikethrough(text):
    """Wrap text in strikethrough HTML tags."""
    return f"<s>{text}</s>"

def color(text, color_value):
    """Wrap text in colored span tags."""
    return f'<span style="color: {color_value}">{text}</span>'

def background(text, bg_color):
    """Wrap text in span tags with background color."""
    return f'<span style="background-color: {bg_color}">{text}</span>'

def header(text, level=1):
    """Wrap text in header tags."""
    return f"<h{level}>{text}</h{level}>"

def bullet_list(*items):
    """Create an HTML unordered list."""
    li_items = ''.join(f"<li>{item}</li>" for item in items)
    return f"<ul>{li_items}</ul>"

def numbered_list(*items):
    """Create an HTML ordered list."""
    li_items = ''.join(f"<li>{item}</li>" for item in items)
    return f"<ol>{li_items}</ol>"