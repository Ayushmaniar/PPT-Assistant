"""
Markdown processing module for PowerPoint text formatting.

This module provides functions to parse markdown text and apply formatting
to PowerPoint TextRange objects.
"""

import re


def parse_markdown_text(markdown_text):
    """
    Parse markdown text and return structured formatting data.
    
    Args:
        markdown_text (str): Text with markdown formatting
        
    Returns:
        tuple: (plain_text, formatting_segments)
            - plain_text: Text without markdown syntax
            - formatting_segments: List of formatting instructions
    """
    segments = []
    
    # Define markdown patterns with their PowerPoint equivalents
    # Order matters: process longer/more specific patterns first
    patterns = [
        # Custom colors with nested formatting
        (r'\{color:([^}]+)\}(.*?)\{/color\}', {'color': True}),
        # Custom background with nested formatting
        (r'\{bg:([^}]+)\}(.*?)\{/bg\}', {'background': True}),
        # Bold patterns (process before italic since ** contains *)
        (r'\*\*(.*?)\*\*', {'bold': True}),
        (r'__(.*?)__', {'bold': True}),
        # Italic patterns (process after bold to avoid conflicts)
        (r'(?<!\*)\*([^*]+?)\*(?!\*)', {'italic': True}),  # Avoid matching ** patterns
        (r'(?<!_)_([^_]+?)_(?!_)', {'italic': True}),      # Avoid matching __ patterns
        # Strikethrough
        (r'~~(.*?)~~', {'strikethrough': True}),
        # Custom underline (not standard markdown)
        (r'\[u\](.*?)\[/u\]', {'underline': True}),
    ]
    
    # Start with the original text
    plain_text = markdown_text
    
    # Process patterns in multiple passes to handle nesting
    total_offset = 0
    
    for pattern, base_formatting in patterns:
        if base_formatting.get('color') or base_formatting.get('background'):
            # Handle color/background patterns with potential nested formatting
            matches = list(re.finditer(pattern, plain_text, re.DOTALL))
            
            for match in reversed(matches):  # Process in reverse to maintain indices
                start, end = match.start(), match.end()
                
                if base_formatting.get('color'):
                    color_value = match.group(1)
                    nested_content = match.group(2)
                    format_dict = {'color': color_value}
                elif base_formatting.get('background'):
                    bg_value = match.group(1)
                    nested_content = match.group(2)
                    format_dict = {'background_color': bg_value}
                
                # Recursively process nested content for additional formatting
                nested_plain, nested_segments = parse_markdown_text(nested_content)
                
                # Replace the color/bg block with just the nested plain text
                plain_text = plain_text[:start] + nested_plain + plain_text[end:]
                
                # Add the color/background formatting for the entire segment
                segments.append({
                    'start': start + 1,  # 1-indexed for PowerPoint
                    'length': len(nested_plain),
                    'formatting': format_dict
                })
                
                # Add nested formatting segments, adjusting their positions
                for nested_seg in nested_segments:
                    segments.append({
                        'start': start + nested_seg['start'],  # Adjust position
                        'length': nested_seg['length'],
                        'formatting': nested_seg['formatting']
                    })
        else:
            # Handle other patterns normally
            matches = list(re.finditer(pattern, plain_text, re.DOTALL))
            
            for match in reversed(matches):  # Process in reverse to maintain indices
                start, end = match.start(), match.end()
                content = match.group(1)
                
                # Replace the markdown syntax with just the content
                plain_text = plain_text[:start] + content + plain_text[end:]
                
                # Add formatting segment
                if content:
                    segments.append({
                        'start': start + 1,  # 1-indexed for PowerPoint
                        'length': len(content),
                        'formatting': base_formatting.copy()
                    })
    
    # Sort segments by start position for consistent application
    segments.sort(key=lambda x: x['start'])
    
    return plain_text, segments


def process_markdown_lists(text):
    """
    Process markdown lists and convert to PowerPoint-friendly format.
    
    Args:
        text (str): Text potentially containing markdown lists
        
    Returns:
        tuple: (processed_text, list_info)
    """
    lines = text.split('\n')
    processed_lines = []
    list_info = []
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Bullet list detection
        if stripped.startswith('- ') or stripped.startswith('* '):
            content = stripped[2:].strip()
            processed_lines.append(f"• {content}")
            list_info.append({
                'line': i,
                'type': 'bullet',
                'indent_level': (len(line) - len(line.lstrip())) // 2
            })
        
        # Numbered list detection
        elif re.match(r'^\d+\.\s', stripped):
            match = re.match(r'^(\d+)\.\s(.+)', stripped)
            if match:
                number, content = match.groups()
                processed_lines.append(f"{number}. {content}")
                list_info.append({
                    'line': i,
                    'type': 'numbered',
                    'number': int(number),
                    'indent_level': (len(line) - len(line.lstrip())) // 2
                })
        
        # Headers
        elif stripped.startswith('#'):
            level = len(stripped) - len(stripped.lstrip('#'))
            content = stripped.lstrip('# ').strip()
            processed_lines.append(content)
            list_info.append({
                'line': i,
                'type': 'header',
                'level': level
            })
        
        else:
            processed_lines.append(line)
    
    return '\n'.join(processed_lines), list_info


def apply_markdown_formatting(text_range, plain_text, segments):
    """
    Apply markdown formatting to a PowerPoint TextRange.
    
    Args:
        text_range: PowerPoint TextRange object
        plain_text (str): Plain text content
        segments (list): Formatting segments from parse_markdown_text
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
