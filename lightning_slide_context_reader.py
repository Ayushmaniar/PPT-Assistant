"""
Lightning-Fast PowerPoint Slide Context Reader

This version achieves the perfect balance of speed and accuracy by using a strategic approach:
1. Keep the EXACT same algorithm as the original for HTML generation
2. Apply targeted optimizations that don't affect the output
3. Focus on reducing COM calls without changing the logic
"""

import win32com.client
import pythoncom
import time
import json
from datetime import datetime

class LightningFastPowerPointSlideReader:
    def __init__(self):
        """Initialize the PowerPoint application connection."""
        pythoncom.CoInitialize()
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            self.presentation = self.ppt_app.ActivePresentation
            self.current_slide_index = None
            self.current_slide_context = ""
            print("✅ Connected to PowerPoint successfully!")
        except Exception as e:
            print(f"❌ Error connecting to PowerPoint: {e}")
            print("Make sure PowerPoint is open with an active presentation.")
            self.ppt_app = None
    
    def get_current_slide_index(self):
        """Get the index of the currently selected/active slide."""
        try:
            if not self.ppt_app:
                return None
            
            # Get the active window
            active_window = self.ppt_app.ActiveWindow
            
            # Method 1: Try to get from the current view (most reliable for normal view)
            try:
                if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                    slide_index = active_window.View.Slide.SlideIndex
                    if slide_index > 0:  # Valid slide index
                        return slide_index
            except:
                pass
            
            # Method 2: Try to get from selection (works in slide sorter view)
            try:
                if (hasattr(active_window, 'Selection') and 
                    hasattr(active_window.Selection, 'SlideRange') and
                    active_window.Selection.SlideRange.Count > 0):
                    return active_window.Selection.SlideRange[0].SlideIndex
            except:
                pass
            
            # Method 3: Try to get from active pane (works in some views)
            try:
                if hasattr(active_window, 'ActivePane') and hasattr(active_window.ActivePane, 'View'):
                    pane_view = active_window.ActivePane.View
                    if hasattr(pane_view, 'Slide'):
                        return pane_view.Slide.SlideIndex
            except:
                pass
            
            # Method 4: Try SlideShowWindow if in slideshow mode
            try:
                if hasattr(self.ppt_app, 'SlideShowWindows') and self.ppt_app.SlideShowWindows.Count > 0:
                    slide_show = self.ppt_app.SlideShowWindows(1)
                    if hasattr(slide_show, 'View') and hasattr(slide_show.View, 'CurrentShowPosition'):
                        return slide_show.View.CurrentShowPosition
            except:
                pass
            
            # Fallback: return 1 if presentation exists
            if self.presentation and self.presentation.Slides.Count > 0:
                return 1
            
            return None
            
        except Exception as e:
            print(f"Error getting current slide index: {e}")
            return 1  # Safe fallback
    
    def analyze_shape(self, shape):
        """Analyze a single shape and extract its properties with HTML formatting detection."""
        try:
            shape_info = {
                'name': shape.Name,
                'type': self.get_shape_type_name(shape.Type),
                'left': round(shape.Left, 2),
                'top': round(shape.Top, 2),
                'width': round(shape.Width, 2),
                'height': round(shape.Height, 2),
                'visible': shape.Visible,
                # Static identifiers for reliable object reference
                'static_id': shape.ID,  # Unique static ID that never changes
                'z_order': shape.ZOrderPosition,  # Layer/stacking order position
                'auto_shape_type': getattr(shape, 'AutoShapeType', None),  # AutoShape specific type
            }
            
            # Text content with HTML formatting detection
            if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                try:
                    text_range = shape.TextFrame.TextRange
                    raw_text = text_range.Text
                    
                    # Convert PowerPoint formatting to HTML - LIGHTNING FAST VERSION
                    html_text = self.convert_powerpoint_text_to_html_lightning(text_range)
                    
                    shape_info['text'] = raw_text  # Keep original for compatibility
                    shape_info['html_text'] = html_text  # Add HTML version
                    shape_info['font_name'] = text_range.Font.Name
                    shape_info['font_size'] = text_range.Font.Size
                    shape_info['font_bold'] = bool(text_range.Font.Bold)
                    shape_info['font_italic'] = bool(text_range.Font.Italic)
                    shape_info['font_color'] = self.get_color_info(text_range.Font.Color)
                except:
                    shape_info['text'] = "Could not read text properties"
                    shape_info['html_text'] = "Could not read text properties"
            
            # Fill properties
            if hasattr(shape, 'Fill'):
                try:
                    fill = shape.Fill
                    shape_info['fill_type'] = self.get_fill_type_name(fill.Type)
                    if fill.Type == 1:  # Solid fill
                        shape_info['fill_color'] = self.get_color_info(fill.ForeColor)
                except:
                    pass
            
            # Line properties
            if hasattr(shape, 'Line'):
                try:
                    line = shape.Line
                    shape_info['line_color'] = self.get_color_info(line.ForeColor)
                    shape_info['line_weight'] = round(line.Weight, 2)
                    shape_info['line_style'] = line.Style
                except:
                    pass
            
            # Special handling for different shape types
            if shape.Type == 17:  # Picture
                try:
                    shape_info['picture_format'] = shape.PictureFormat.CompressLevel
                except:
                    pass
            
            elif shape.Type == 3:  # Chart
                try:
                    if hasattr(shape, 'Chart'):
                        shape_info['chart_type'] = shape.Chart.ChartType
                        shape_info['chart_title'] = shape.Chart.ChartTitle.Text if shape.Chart.HasTitle else "No title"
                except:
                    pass
            
            elif shape.Type == 19:  # Table
                try:
                    if hasattr(shape, 'Table'):
                        table = shape.Table
                        shape_info['table_rows'] = table.Rows.Count
                        shape_info['table_columns'] = table.Columns.Count
                        # Read ALL cell content with HTML formatting - LIGHTNING FAST
                        all_cells = []
                        all_cells_html = []
                        for row in range(table.Rows.Count):
                            row_cells = []
                            row_cells_html = []
                            for col in range(table.Columns.Count):
                                try:
                                    cell_shape = table.Cell(row + 1, col + 1).Shape
                                    cell_text = cell_shape.TextFrame.TextRange.Text.strip()
                                    cell_html = self.convert_powerpoint_text_to_html_lightning(cell_shape.TextFrame.TextRange)
                                    
                                    row_cells.append(cell_text if cell_text else "[Empty]")
                                    row_cells_html.append(cell_html if cell_html else "[Empty]")
                                except:
                                    row_cells.append("[Error reading cell]")
                                    row_cells_html.append("[Error reading cell]")
                            all_cells.append(row_cells)
                            all_cells_html.append(row_cells_html)
                        shape_info['table_cells'] = all_cells
                        shape_info['table_cells_html'] = all_cells_html
                except:
                    pass
            
            return shape_info
            
        except Exception as e:
            return {
                'name': f"Shape analysis error: {e}",
                'type': 'Unknown',
                'error': str(e)
            }
    
    def get_layout_name_safe(self, slide):
        """Safely get layout name with error handling."""
        try:
            return slide.Layout.Name
        except:
            return "Unknown Layout"
    
    def convert_powerpoint_text_to_html_lightning(self, text_range):
        """
        LIGHTNING FAST version that uses the .Runs() collection for maximum efficiency.
        This approach iterates through segments of text that share the same formatting,
        dramatically reducing the number of COM calls compared to a character-by-character check.
        """
        try:
            if not hasattr(text_range, 'Runs') or not text_range.Text:
                return text_range.Text if hasattr(text_range, 'Text') else ""

            html_parts = []
            
            # Get the collection of text runs. Each run has uniform formatting.
            runs = text_range.Runs()
            if not runs:
                return text_range.Text # Fallback for empty text ranges

            # OPTIMIZATION: Get default color once from the parent range
            default_color = 0 # Default to black
            try:
                if text_range.Font.Color.Type == 2: # msoColorTypeRGB
                     default_color = text_range.Font.Color.RGB
            except:
                pass

            for run in runs:
                run_font = run.Font
                run_text = run.Text
                
                if not run_text.strip(): # Skip whitespace-only runs
                    html_parts.append(run_text)
                    continue

                open_tags = []
                close_tags = []

                # --- Check formatting properties for the run ---
                if run_font.Bold:
                    open_tags.append('<b>')
                    close_tags.insert(0, '</b>')
                if run_font.Italic:
                    open_tags.append('<i>')
                    close_tags.insert(0, '</i>')
                if run_font.Underline:
                    open_tags.append('<u>')
                    close_tags.insert(0, '</u>')
                
                try:
                    if run_font.Strikethrough: # Not all versions have this
                        open_tags.append('<s>')
                        close_tags.insert(0, '</s>')
                except:
                    pass
                
                # --- Handle color ---
                hex_color = None
                try:
                    # Attempt to get the color as a direct RGB value first
                    color_bgr = run_font.Color.RGB
                    r = color_bgr & 0xFF
                    g = (color_bgr >> 8) & 0xFF
                    b = (color_bgr >> 16) & 0xFF
                    hex_color = f"#{r:02x}{g:02x}{b:02x}"
                except Exception:
                    # If direct RGB fails, it's likely a theme color
                    try:
                        theme_color_index = run_font.Color.ObjectThemeColor
                        theme_color_bgr = self.presentation.SlideMaster.Theme.ThemeColorScheme(theme_color_index).RGB
                        r = theme_color_bgr & 0xFF
                        g = (theme_color_bgr >> 8) & 0xFF
                        b = (theme_color_bgr >> 16) & 0xFF
                        hex_color = f"#{r:02x}{g:02x}{b:02x}"
                    except Exception as theme_error:
                        # If both fail, we cannot determine the color
                        # print(f"DEBUG: Could not read theme color for run '{run_text[:30]}...'. Error: {theme_error}")
                        pass

                # Add span tag if we found a valid, non-black color
                if hex_color and hex_color.lower() != "#000000":
                    open_tags.append(f'<span style="color: {hex_color}">')
                    close_tags.insert(0, '</span>')

                # Convert line breaks to <br> tags FIRST (before HTML escaping to avoid conflicts)
                escaped_text = run_text.replace('\r\n', '<br>').replace('\r', '<br>').replace('\n', '<br>')
                
                # Clean up other problematic whitespace characters that might cause display issues
                escaped_text = escaped_text.replace('\x0b', '<br>')  # Vertical tab
                escaped_text = escaped_text.replace('\x0c', '<br>')  # Form feed
                escaped_text = escaped_text.replace('\xa0', ' ')     # Non-breaking space to regular space
                escaped_text = escaped_text.replace('\u2028', '<br>')  # Line separator
                escaped_text = escaped_text.replace('\u2029', '<br>')  # Paragraph separator
                
                # Then escape HTML special characters
                escaped_text = escaped_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                
                # But preserve our <br> tags by converting them back
                escaped_text = escaped_text.replace('&lt;br&gt;', '<br>')

                # Assemble the final HTML for this run
                formatted_text = ''.join(open_tags) + escaped_text + ''.join(close_tags)
                html_parts.append(formatted_text)

            return ''.join(html_parts)

        except Exception as e:
            # Fallback to plain text on any error
            print(f"Error in HTML conversion, falling back to plain text: {e}")
            return text_range.Text if hasattr(text_range, 'Text') else ""
    
    def get_shape_type_name(self, shape_type):
        """Convert shape type number to readable name."""
        shape_types = {
            1: "AutoShape",
            2: "Callout", 
            3: "Chart",
            4: "Comment",
            5: "Freeform",
            6: "Group",
            7: "Embedded OLE Object",
            8: "Line",
            9: "Linked OLE Object",
            10: "Linked Picture",
            11: "Media",
            12: "OLE Control",
            13: "Picture", 
            14: "Placeholder",
            15: "Text Effect",
            16: "Title",
            17: "Picture",
            18: "Script Anchor",
            19: "Table",
            20: "Canvas",
            21: "Diagram",
            22: "Ink",
            23: "Ink Comment",
            24: "Smart Art",
            25: "Web Video"
        }
        return shape_types.get(shape_type, f"Unknown Type ({shape_type})")
    
    def get_fill_type_name(self, fill_type):
        """Convert fill type number to readable name."""
        fill_types = {
            0: "Mixed",
            1: "Solid",
            2: "Gradient", 
            3: "Textured",
            4: "Pattern",
            5: "Picture",
            6: "Background"
        }
        return fill_types.get(fill_type, f"Unknown Fill ({fill_type})")
    
    def get_color_info(self, color_obj):
        """Extract color information."""
        try:
            return {
                'rgb': color_obj.RGB,
                'type': color_obj.Type
            }
        except:
            return "Could not read color"
    
    def read_slide_content(self, slide_index):
        """Read all content from a specific slide."""
        try:
            if not self.presentation:
                return "No active presentation"
            
            if slide_index > self.presentation.Slides.Count:
                return f"Slide {slide_index} does not exist (total slides: {self.presentation.Slides.Count})"
            
            slide = self.presentation.Slides(slide_index)
            
            slide_info = {
                'slide_index': slide_index,
                'slide_name': slide.Name,
                'layout_name': self.get_layout_name_safe(slide),
                'timestamp': datetime.now().isoformat(),
                'total_shapes': slide.Shapes.Count,
                'shapes': []
            }
            
            # Analyze each shape in the slide
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                shape_info = self.analyze_shape(shape)
                slide_info['shapes'].append(shape_info)
            
            # Check for slide notes
            try:
                if slide.NotesPage.Shapes.Count > 1:  # Usually shape 1 is the slide thumbnail
                    notes_shape = slide.NotesPage.Shapes(2)  # Notes text is usually shape 2
                    if notes_shape.TextFrame.HasText:
                        slide_info['notes'] = notes_shape.TextFrame.TextRange.Text
                else:
                    slide_info['notes'] = ""
            except:
                slide_info['notes'] = "Could not read notes"
            
            # Check for animations
            try:
                timeline = slide.TimeLine
                if timeline.MainSequence.Count > 0:
                    animations = []
                    for j in range(1, timeline.MainSequence.Count + 1):
                        effect = timeline.MainSequence(j)
                        animation_info = {
                            'effect_type': effect.EffectType,
                            'trigger_type': effect.Timing.TriggerType,
                            'shape_name': effect.Shape.Name if hasattr(effect, 'Shape') else "Unknown"
                        }
                        animations.append(animation_info)
                    slide_info['animations'] = animations
                else:
                    slide_info['animations'] = []
            except:
                slide_info['animations'] = "Could not read animations"
            
            return slide_info
            
        except Exception as e:
            return f"Error reading slide {slide_index}: {e}"
    
    def format_slide_context(self, slide_info):
        """
        Format slide information into a readable context string with HTML formatting.
        
        *** IMPORTANT: ALL TEXT CONTENT IN THIS CONTEXT IS PROVIDED IN HTML FORMAT ***
        
        The text content includes HTML tags for formatting:
        - <b>bold text</b> for bold formatting
        - <i>italic text</i> for italic formatting  
        - <u>underlined text</u> for underlined formatting
        - <s>strikethrough text</s> for strikethrough formatting
        - <span style="color: #RRGGBB">colored text</span> for colored text
        
        When modifying text, you can use these same HTML tags in your tool calls.
        """
        if isinstance(slide_info, str):
            return slide_info
        
        context = f"""
=== POWERPOINT SLIDE CONTEXT (HTML FORMATTED) ===
*** IMPORTANT: ALL TEXT CONTENT BELOW IS IN HTML FORMAT ***
*** Use HTML tags like <b>, <i>, <u>, <s>, <span style="color: #RRGGBB"> when modifying text ***

Slide: {slide_info['slide_index']} of {self.presentation.Slides.Count}
Name: {slide_info['slide_name']}
Layout: {slide_info['layout_name']}
Total Objects: {slide_info['total_shapes']}
Last Updated: {slide_info['timestamp']}

=== SLIDE CONTENT (HTML FORMATTED) ===
"""
        
        if slide_info['shapes']:
            for i, shape in enumerate(slide_info['shapes'], 1):
                context += f"\n--- Object {i}: {shape['name']} ---\n"
                context += f"Type: {shape['type']}\n"
                context += f"Position: ({shape.get('left', 'N/A')}, {shape.get('top', 'N/A')})\n"
                context += f"Size: {shape.get('width', 'N/A')} x {shape.get('height', 'N/A')}\n"
                context += f"ID: {shape['static_id']}\n"
                
                if 'html_text' in shape and shape['html_text']:
                    # Show HTML-formatted text as the primary text content
                    context += f"Text: {shape['html_text']}\n"
                    
                    if 'font_name' in shape:
                        context += f"Font: {shape['font_name']}, Size: {shape.get('font_size', 'N/A')}\n"
                        if shape.get('font_bold') or shape.get('font_italic'):
                            styles = []
                            if shape.get('font_bold'): styles.append("Bold")
                            if shape.get('font_italic'): styles.append("Italic")
                            context += f"Base Styles: {', '.join(styles)}\n"
                elif 'text' in shape:
                    # Fallback to plain text if HTML conversion failed
                    context += f"Text: {shape['text']}\n"
                    
                    if 'font_name' in shape:
                        context += f"Font: {shape['font_name']}, Size: {shape.get('font_size', 'N/A')}\n"
                        if shape.get('font_bold') or shape.get('font_italic'):
                            styles = []
                            if shape.get('font_bold'): styles.append("Bold")
                            if shape.get('font_italic'): styles.append("Italic")
                            context += f"Styles: {', '.join(styles)}\n"
                
                if 'table_rows' in shape:
                    context += f"Table: {shape['table_rows']} rows x {shape['table_columns']} columns\n"
                    
                    # Show HTML-formatted table content if available
                    if shape.get('table_cells_html'):
                        context += "Table content:\n"
                        for row_idx, row_data in enumerate(shape['table_cells_html']):
                            row_str = " | ".join(row_data)
                            context += f"  Row {row_idx + 1}: {row_str}\n"
                    
                    # Fallback to plain table content if HTML is not available
                    elif shape.get('table_cells'):
                        context += "Table content:\n"
                        for row_idx, row_data in enumerate(shape['table_cells']):
                            row_str = " | ".join(row_data)
                            context += f"  Row {row_idx + 1}: {row_str}\n"
                
                if 'chart_type' in shape:
                    context += f"Chart Type: {shape['chart_type']}\n"
                    context += f"Chart Title: {shape.get('chart_title', 'No title')}\n"
        else:
            context += "\n[No objects found on this slide]\n"
        
        if slide_info.get('notes'):
            context += f"\n=== SLIDE NOTES (HTML FORMATTED) ===\n{slide_info['notes']}\n"
        
        if slide_info.get('animations') and isinstance(slide_info['animations'], list) and slide_info['animations']:
            context += f"\n=== ANIMATIONS ===\n"
            for i, anim in enumerate(slide_info['animations'], 1):
                context += f"Animation {i}: Type {anim['effect_type']} on {anim['shape_name']}\n"
        
        context += "\n=== END CONTEXT (Remember: Text is HTML formatted!) ===\n"
        
        return context
    
    def monitor_slide_changes(self, interval=2, max_iterations=None):
        """Monitor for slide changes and update context accordingly."""
        print("🔍 Starting slide monitoring...")
        print("Switch between slides in PowerPoint to see context updates.")
        print("Press Ctrl+C to stop monitoring.\n")
        
        iteration = 0
        try:
            while True:
                if max_iterations and iteration >= max_iterations:
                    break
                
                current_slide = self.get_current_slide_index()
                
                if current_slide != self.current_slide_index:
                    print(f"\n📍 Slide changed: {self.current_slide_index} → {current_slide}")
                    print("=" * 60)
                    
                    self.current_slide_index = current_slide
                    slide_info = self.read_slide_content(current_slide)
                    self.current_slide_context = self.format_slide_context(slide_info)
                    
                    print(self.current_slide_context)
                    print("=" * 60)
                
                time.sleep(interval)
                iteration += 1
                
        except KeyboardInterrupt:
            print("\n🛑 Monitoring stopped by user.")
        except Exception as e:
            print(f"\n❌ Error during monitoring: {e}")
    
    def get_current_context(self):
        """Get the current slide context. Always checks for slide changes."""
        try:
            # Always get the current slide index to detect changes
            current_slide = self.get_current_slide_index()
            
            if current_slide is None:
                return "Could not determine current slide"
            
            # Check if the slide has changed or if we don't have cached context
            if current_slide != self.current_slide_index or not self.current_slide_context:
                print(f"🔄 Slide context updating: {self.current_slide_index} → {current_slide}")
                self.current_slide_index = current_slide
                slide_info = self.read_slide_content(current_slide)
                self.current_slide_context = self.format_slide_context(slide_info)
            
            return self.current_slide_context
            
        except Exception as e:
            return f"Error getting current context: {e}"
    
    def force_refresh_context(self):
        """Force refresh the current slide context regardless of slide index."""
        try:
            current_slide = self.get_current_slide_index()
            
            if current_slide is None:
                return "Could not determine current slide"
            
            # Force refresh by reading the slide content again
            print(f"🔄 Force refreshing slide context for slide {current_slide}")
            self.current_slide_index = current_slide
            slide_info = self.read_slide_content(current_slide)
            self.current_slide_context = self.format_slide_context(slide_info)
            
            return self.current_slide_context
            
        except Exception as e:
            return f"Error force refreshing context: {e}"
    
    def clear_context_cache(self):
        """Clear the cached context to force a refresh on next access."""
        print("🗑️ Clearing slide context cache")
        self.current_slide_context = ""
        self.current_slide_index = None


def test_lightning_slide_reader():
    """Test the lightning-fast slide reader functionality."""
    print("🚀 Testing Lightning-Fast PowerPoint Slide Context Reader")
    print("=" * 50)
    
    reader = LightningFastPowerPointSlideReader()
    
    if not reader.ppt_app:
        print("Cannot continue without PowerPoint connection.")
        return
    
    # Test 1: Get current slide context
    print("\n📖 Test 1: Reading current slide context")
    context = reader.get_current_context()
    print(context)
    
    print("\n✅ Lightning-fast testing completed!")


if __name__ == "__main__":
    test_lightning_slide_reader()
