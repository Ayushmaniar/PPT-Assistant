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
        # Find matching opening tag
        for i in range(len(self.tag_stack) - 1, -1, -1):
            if self.tag_stack[i]['tag'] == tag:
                tag_info = self.tag_stack.pop(i)
                
                # Only create format segment if there was content
                if self.current_position > tag_info['start_position']:
                    self.format_segments.append({
                        'start': tag_info['start_position'] + 1,  # 1-indexed for PowerPoint
                        'length': self.current_position - tag_info['start_position'],
                        'formatting': tag_info['formatting']
                    })
                break
    
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
    
    # Apply formatting to each segment
    for segment in segments:
        if not segment['formatting']:
            continue
            
        try:
            start_pos = segment['start']
            length = segment['length']
            
            # Ensure we don't exceed text bounds
            if start_pos > len(plain_text) or start_pos + length - 1 > len(plain_text):
                continue
            
            # Get the character range for this segment
            char_range = text_range.Characters(start_pos, length)
            
            # Apply formatting
            formatting = segment['formatting']
            
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
                    # Try alternative property names if Strikethrough doesn't work
                    try:
                        char_range.Font.Strike = -1
                    except:
                        pass  # Strikethrough not supported in all versions
                    
            if formatting.get('color'):
                try:
                    color_value = formatting['color']
                    if color_value.startswith('#'):
                        # Convert hex to RGB - PowerPoint uses BGR format
                        hex_color = color_value[1:]
                        if len(hex_color) == 6:
                            # Extract R, G, B components
                            r = int(hex_color[0:2], 16)
                            g = int(hex_color[2:4], 16) 
                            b = int(hex_color[4:6], 16)
                            # PowerPoint uses BGR format: B + (G * 256) + (R * 65536)
                            bgr_color = b + (g * 256) + (r * 65536)
                            char_range.Font.Color.RGB = bgr_color
                    else:
                        # Named colors (basic support)
                        color_map = {
                            'red': 255, 'blue': 16711680, 'green': 65280,
                            'yellow': 65535, 'orange': 33023, 'purple': 8388736,
                            'black': 0, 'white': 16777215
                        }
                        if color_value.lower() in color_map:
                            char_range.Font.Color.RGB = color_map[color_value.lower()]
                except Exception as e:
                    print(f"Warning: Could not apply color {formatting.get('color')}: {e}")
                    
            if formatting.get('background_color'):
                try:
                    bg_value = formatting['background_color']
                    if bg_value.startswith('#'):
                        hex_color = bg_value[1:]
                        rgb_color = int(hex_color[4:6] + hex_color[2:4] + hex_color[0:2], 16)
                        char_range.Font.Fill.ForeColor.RGB = rgb_color
                except Exception as e:
                    print(f"Warning: Could not apply background color {formatting.get('background_color')}: {e}")
                    
        except Exception as e:
            print(f"Warning: Could not apply formatting to segment {segment}: {e}")


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