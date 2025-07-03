"""
PowerPoint Slide Context Reader - Test Script

This script tests the functionality to:
1. Get the currently selected slide in PowerPoint
2. Read all objects/shapes/content present in that slide
3. Monitor slide changes and update context accordingly
"""

import win32com.client
import pythoncom
import time
import json
from datetime import datetime

class PowerPointSlideReader:
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
        """Get the index of the currently selected slide."""
        try:
            if not self.ppt_app:
                return None
            
            # Get the active window and current slide
            active_window = self.ppt_app.ActiveWindow
            if hasattr(active_window, 'Selection') and hasattr(active_window.Selection, 'SlideRange'):
                # If slides are selected in slide sorter or normal view
                if active_window.Selection.SlideRange.Count > 0:
                    return active_window.Selection.SlideRange[0].SlideIndex
            
            # Alternative method: get from slide show or normal view
            if hasattr(active_window, 'View') and hasattr(active_window.View, 'Slide'):
                return active_window.View.Slide.SlideIndex
            
            # Fallback: assume first slide
            return 1
            
        except Exception as e:
            print(f"Error getting current slide index: {e}")
            return 1
    
    def analyze_shape(self, shape):
        """Analyze a single shape and extract its properties."""
        try:
            shape_info = {
                'name': shape.Name,
                'type': self.get_shape_type_name(shape.Type),
                'left': round(shape.Left, 2),
                'top': round(shape.Top, 2),
                'width': round(shape.Width, 2),
                'height': round(shape.Height, 2),
                'visible': shape.Visible,
            }
            
            # Text content
            if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                try:
                    text_range = shape.TextFrame.TextRange
                    shape_info['text'] = text_range.Text
                    shape_info['font_name'] = text_range.Font.Name
                    shape_info['font_size'] = text_range.Font.Size
                    shape_info['font_bold'] = bool(text_range.Font.Bold)
                    shape_info['font_italic'] = bool(text_range.Font.Italic)
                    shape_info['font_color'] = self.get_color_info(text_range.Font.Color)
                except:
                    shape_info['text'] = "Could not read text properties"
            
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
                        # Read ALL cell content
                        all_cells = []
                        for row in range(table.Rows.Count):
                            row_cells = []
                            for col in range(table.Columns.Count):
                                try:
                                    cell_text = table.Cell(row + 1, col + 1).Shape.TextFrame.TextRange.Text.strip()
                                    row_cells.append(cell_text if cell_text else "[Empty]")
                                except:
                                    row_cells.append("[Error reading cell]")
                            all_cells.append(row_cells)
                        shape_info['table_cells'] = all_cells
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
        """Format slide information into a readable context string."""
        if isinstance(slide_info, str):
            return slide_info
        
        context = f"""
=== POWERPOINT SLIDE CONTEXT ===
Slide: {slide_info['slide_index']} of {self.presentation.Slides.Count}
Name: {slide_info['slide_name']}
Layout: {slide_info['layout_name']}
Total Objects: {slide_info['total_shapes']}
Last Updated: {slide_info['timestamp']}

=== SLIDE CONTENT ===
"""
        
        if slide_info['shapes']:
            for i, shape in enumerate(slide_info['shapes'], 1):
                context += f"\n--- Object {i}: {shape['name']} ---\n"
                context += f"Type: {shape['type']}\n"
                context += f"Position: ({shape.get('left', 'N/A')}, {shape.get('top', 'N/A')})\n"
                context += f"Size: {shape.get('width', 'N/A')} x {shape.get('height', 'N/A')}\n"
                
                if 'text' in shape:
                    context += f"Text: {shape['text'][:100]}{'...' if len(shape['text']) > 100 else ''}\n"
                    if 'font_name' in shape:
                        context += f"Font: {shape['font_name']}, Size: {shape.get('font_size', 'N/A')}\n"
                        if shape.get('font_bold') or shape.get('font_italic'):
                            styles = []
                            if shape.get('font_bold'): styles.append("Bold")
                            if shape.get('font_italic'): styles.append("Italic")
                            context += f"Styles: {', '.join(styles)}\n"
                
                if 'table_rows' in shape:
                    context += f"Table: {shape['table_rows']} rows x {shape['table_columns']} columns\n"
                    if shape.get('table_cells'):
                        context += "Table content:\n"
                        for row_idx, row_data in enumerate(shape['table_cells']):
                            # Format each row nicely
                            row_str = " | ".join([f"{cell[:30]}{'...' if len(cell) > 30 else ''}" for cell in row_data])
                            context += f"  Row {row_idx + 1}: {row_str}\n"
                
                if 'chart_type' in shape:
                    context += f"Chart Type: {shape['chart_type']}\n"
                    context += f"Chart Title: {shape.get('chart_title', 'No title')}\n"
        else:
            context += "\n[No objects found on this slide]\n"
        
        if slide_info.get('notes'):
            context += f"\n=== SLIDE NOTES ===\n{slide_info['notes']}\n"
        
        if slide_info.get('animations') and isinstance(slide_info['animations'], list) and slide_info['animations']:
            context += f"\n=== ANIMATIONS ===\n"
            for i, anim in enumerate(slide_info['animations'], 1):
                context += f"Animation {i}: Type {anim['effect_type']} on {anim['shape_name']}\n"
        
        context += "\n=== END CONTEXT ===\n"
        
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
        """Get the current slide context."""
        if self.current_slide_index is None:
            current_slide = self.get_current_slide_index()
            if current_slide:
                self.current_slide_index = current_slide
                slide_info = self.read_slide_content(current_slide)
                self.current_slide_context = self.format_slide_context(slide_info)
        
        return self.current_slide_context


def test_slide_reader():
    """Test the slide reader functionality."""
    print("🚀 Testing PowerPoint Slide Context Reader")
    print("=" * 50)
    
    reader = PowerPointSlideReader()
    
    if not reader.ppt_app:
        print("Cannot continue without PowerPoint connection.")
        return
    
    # Test 1: Get current slide context
    print("\n📖 Test 1: Reading current slide context")
    context = reader.get_current_context()
    print(context)
    
    # Test 2: Monitor slide changes for a limited time
    print("\n👀 Test 2: Monitoring slide changes")
    print("Please switch between slides in PowerPoint...")
    reader.monitor_slide_changes(interval=1, max_iterations=30)  # Monitor for 30 seconds
    
    print("\n✅ Testing completed!")


if __name__ == "__main__":
    test_slide_reader()
