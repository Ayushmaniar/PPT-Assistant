import win32com.client
import pythoncom
import cv2
import numpy as np
import os
from slide_context_reader import PowerPointSlideReader
import time
import base64

class SlideVisualizer:
    def __init__(self):
        """
        Initializes the SlideVisualizer by connecting to the PowerPoint application
        via the PowerPointSlideReader.
        """
        self.reader = PowerPointSlideReader()
        if not self.reader.ppt_app:
            raise ConnectionError("PowerPoint is not open or accessible.")
        self.presentation = self.reader.presentation
        self.slide_width_points = self.presentation.PageSetup.SlideWidth
        self.slide_height_points = self.presentation.PageSetup.SlideHeight

    @staticmethod
    def image_to_base64(image, format='JPEG', quality=85):
        """
        Convert an OpenCV image to base64 string for vision model input.
        
        Args:
            image: OpenCV image (numpy array)
            format: Image format ('JPEG', 'PNG')
            quality: JPEG quality (1-100, only used for JPEG)
            
        Returns:
            str: Base64 encoded image string with data URI prefix
        """
        if image is None:
            return None
            
        try:
            # Encode image to bytes
            if format.upper() == 'JPEG':
                encode_param = [int(cv2.IMWRITE_JPEG_QUALITY), quality]
                success, buffer = cv2.imencode('.jpg', image, encode_param)
                mime_type = 'image/jpeg'
            elif format.upper() == 'PNG':
                success, buffer = cv2.imencode('.png', image)
                mime_type = 'image/png'
            else:
                raise ValueError(f"Unsupported format: {format}")
                
            if not success:
                print(f"❌ Failed to encode image to {format}")
                return None
                
            # Convert to base64
            image_base64 = base64.b64encode(buffer).decode('utf-8')
            
            # Return with data URI prefix
            return f"data:{mime_type};base64,{image_base64}"
            
        except Exception as e:
            print(f"❌ Error converting image to base64: {e}")
            return None

    def _get_slide_export_dimensions(self, image_width):
        """Calculates the export height based on a given width to maintain aspect ratio."""
        aspect_ratio = self.slide_height_points / self.slide_width_points
        image_height = int(image_width * aspect_ratio)
        return image_width, image_height

    def get_downsampled_slide_image(self, target_width=512, file_format="JPG"):
        """
        Exports the current slide, downsamples it, and then adds highlights with scaled fonts.
        """
        try:
            # 1. Get slide data and index
            slide_index = self.reader.get_current_slide_index()
            if not slide_index:
                print("❌ Could not get current slide index.")
                return None
            slide = self.presentation.Slides(slide_index)
            slide_info = self.reader.read_slide_content_lean(slide_index)

            # 2. Export and read the high-res image
            temp_file_path = os.path.abspath(f"temp_downsample.{file_format.lower()}")
            export_width, export_height = self._get_slide_export_dimensions(1920)
            slide.Export(temp_file_path, file_format, export_width, export_height)
            image = cv2.imread(temp_file_path)
            if image is None:
                print(f"❌ Failed to load image from: {temp_file_path}")
                os.remove(temp_file_path)
                return None

            # 3. Downsample the image FIRST
            aspect_ratio = image.shape[0] / image.shape[1]
            target_height = int(target_width * aspect_ratio)
            downsampled_image = cv2.resize(image, (target_width, target_height), interpolation=cv2.INTER_AREA)
            
            # 4. Draw overlays onto the downsampled image
            # Recalculate scale factors for the new, smaller dimensions
            scale_x = downsampled_image.shape[1] / self.slide_width_points
            scale_y = downsampled_image.shape[0] / self.slide_height_points
            
            box_color = (0, 255, 0)
            label_bg_color = (0, 255, 255)
            label_text_color = (0, 0, 0)

            # Dynamically calculate font size based on image width
            font_scale = max(0.3, target_width / 2000.0) # Make font smaller
            font_thickness = max(1, int(target_width / 512.0))

            for shape in slide_info.get('shapes', []):
                static_id = shape.get('static_id')
                if static_id is None: continue

                x = int(shape.get('left', 0) * scale_x)
                y = int(shape.get('top', 0) * scale_y)
                w = int(shape.get('width', 0) * scale_x)
                h = int(shape.get('height', 0) * scale_y)

                # Use thinner lines for the smaller image
                cv2.rectangle(downsampled_image, (x, y), (x + w, y + h), box_color, 1)
                
                id_text = f"ID:{static_id}"
                font = cv2.FONT_HERSHEY_SIMPLEX
                # Use the new dynamic font scale and thickness
                (tw, th), _ = cv2.getTextSize(id_text, font, font_scale, font_thickness)
                
                text_x, text_y = x, y - 5
                if text_y < th: text_y = y + h + th + 5

                cv2.rectangle(downsampled_image, (text_x, text_y - th - 2), (text_x + tw + 2, text_y + 2), label_bg_color, -1)
                cv2.putText(downsampled_image, id_text, (text_x + 1, text_y), font, font_scale, label_text_color, font_thickness, cv2.LINE_AA)

            # 5. Clean up the temporary file
            os.remove(temp_file_path)

            print(f"✅ Successfully created downsampled image with overlays of size {downsampled_image.shape[1]}x{downsampled_image.shape[0]}.")
            return downsampled_image

        except Exception as e:
            print(f"❌ An error occurred while generating the downsampled image: {e}")
            return None

    def get_visual_context_for_agent(self, target_width=512, include_description=True):
        """
        Generate visual context for the AI agent including base64 image and description.
        
        Args:
            target_width: Width for the downsampled image
            include_description: Whether to include descriptive text about the image
            
        Returns:
            dict: Contains 'image_base64', 'description', and 'success' keys
        """
        try:
            # Get the downsampled annotated image
            image = self.get_downsampled_slide_image(target_width=target_width)
            
            if image is None:
                return {
                    'success': False,
                    'image_base64': None,
                    'description': 'Failed to generate slide image'
                }
            
            # Convert to base64
            image_base64 = self.image_to_base64(image, format='JPEG', quality=80)
            
            if image_base64 is None:
                return {
                    'success': False,
                    'image_base64': None,
                    'description': 'Failed to convert image to base64'
                }
            
            # Generate description
            description = ""
            if include_description:
                description = f"""
=== VISUAL SLIDE REPRESENTATION ===

This is an annotated visual representation of the current PowerPoint slide with the following features:

🔍 VISUAL ANNOTATIONS:
- Green bounding boxes highlight all interactive objects/shapes on the slide
- Yellow labels show unique object IDs (e.g., "ID:123") for precise reference
- Image is downsampled to {target_width}px width for efficient processing
- All text, images, charts, and other slide elements are visually represented

💡 HOW TO USE THIS IMAGE:
- Use this visual context to understand the spatial layout of slide elements
- Reference object IDs when making modifications (the IDs match the textual context)
- Analyze positioning, sizing, and visual relationships between elements
- Identify visual design issues, alignment problems, or layout improvements
- This complements the textual slide context for comprehensive understanding

⚠️ IMPORTANT NOTES:
- Object IDs in yellow labels correspond exactly to the IDs in the textual context
- Use the textual context for precise measurements and detailed properties
- This image shows the current state of the slide at the time of generation
- Visual and textual contexts are synchronized and represent the same slide state

This visual representation enables you to provide more accurate and contextually aware assistance with slide design, layout, and content positioning.
=== END VISUAL CONTEXT ===
"""
            
            return {
                'success': True,
                'image_base64': image_base64,
                'description': description.strip()
            }
            
        except Exception as e:
            return {
                'success': False,
                'image_base64': None,
                'description': f'Error generating visual context: {str(e)}'
            }

    def create_highlighted_slide_image(self, output_path="highlighted_slide.png", export_width=1920, border_size=50):
        """
        Creates an image of the current slide with rulers and highlighted objects.

        Args:
            output_path (str): The path to save the highlighted image.
            export_width (int): The width in pixels for the exported PNG image.
            border_size (int): The size of the border for drawing rulers.

        Returns:
            str: The path to the saved image, or None if an error occurred.
        """
        t_start = time.time()

        # 1. Get current slide context
        t_read_start = time.time()
        slide_index = self.reader.get_current_slide_index()
        if not slide_index:
            print("❌ Could not get current slide index.")
            return None
        slide_info = self.reader.read_slide_content_lean(slide_index)
        t_read_end = time.time()
        
        if isinstance(slide_info, str):
            print(f"❌ Error reading slide content: {slide_info}")
            return None

        # 2. Export slide as a temporary PNG
        t_export_start = time.time()
        slide = self.presentation.Slides(slide_index)
        temp_png_path = os.path.abspath(f"temp_slide_{slide_index}.png")
        
        width, height = self._get_slide_export_dimensions(export_width)
        slide.Export(temp_png_path, "PNG", width, height)
        
        t_export_end = time.time()

        # 3. Load the exported image with OpenCV
        t_draw_start = time.time()
        slide_image = cv2.imread(temp_png_path)
        if slide_image is None:
            print(f"❌ Failed to load image from: {temp_png_path}")
            if os.path.exists(temp_png_path):
                os.remove(temp_png_path)
            return None
        
        # 4. Create a new canvas with a border
        img_height, img_width, _ = slide_image.shape
        canvas_height = img_height + 2 * border_size
        canvas_width = img_width + 2 * border_size
        canvas = np.full((canvas_height, canvas_width, 3), 255, dtype=np.uint8)
        
        # Paste the slide image onto the canvas
        canvas[border_size : border_size + img_height, border_size : border_size + img_width] = slide_image
        
        # 5. Calculate scaling factors
        scale_x = img_width / self.slide_width_points
        scale_y = img_height / self.slide_height_points

        # 6. Draw rulers on the canvas
        self._draw_rulers(canvas, border_size, img_width, img_height, scale_x, scale_y)
        
        # 7. Draw overlays for each shape on the canvas
        # Define high-contrast colors (BGR format)
        box_color = (0, 255, 0)  # Bright Green
        label_bg_color = (0, 255, 255) # Bright Yellow
        label_text_color = (0, 0, 0) # Black

        shapes = slide_info.get('shapes', [])

        # Dynamically calculate font size based on image width
        font_scale = max(0.5, export_width / 2500.0) # Make font bigger
        font_thickness = max(1, int(export_width / 800.0)) # Make font thicker

        for shape in shapes:
            try:
                static_id = shape.get('static_id')
                if static_id is None:
                    continue

                # Scale coordinates and offset for the canvas
                x = int(shape.get('left', 0) * scale_x) + border_size
                y = int(shape.get('top', 0) * scale_y) + border_size
                w = int(shape.get('width', 0) * scale_x)
                h = int(shape.get('height', 0) * scale_y)

                # Draw bounding box
                cv2.rectangle(canvas, (x, y), (x + w, y + h), box_color, 2)

                # Prepare text label for the shape ID
                id_text = f"ID:{static_id}"
                font = cv2.FONT_HERSHEY_SIMPLEX
                # Use the new dynamic font scale and thickness
                text_size, _ = cv2.getTextSize(id_text, font, font_scale, font_thickness)
                
                text_x = x
                text_y = y - 10
                if text_y < text_size[1]:
                    text_y = y + h + text_size[1] + 5

                # Add filled background for text
                cv2.rectangle(canvas, (text_x, text_y - text_size[1] - 5), 
                              (text_x + text_size[0] + 5, text_y + 5), label_bg_color, -1)
                cv2.putText(canvas, id_text, (text_x + 2, text_y), 
                            font, font_scale, label_text_color, font_thickness, cv2.LINE_AA)

            except Exception as e:
                print(f"⚠️ Error processing shape {shape.get('name', 'N/A')}: {e}")

        t_draw_end = time.time()

        # 8. Save the final image
        t_save_start = time.time()
        cv2.imwrite(output_path, canvas)
        t_save_end = time.time()

        # 9. Clean up temporary file
        os.remove(temp_png_path)
        
        t_end = time.time()

        # Print performance metrics
        print("\n--- Performance ---")
        print(f"Read slide content: {t_read_end - t_read_start:.4f}s")
        print(f"Export slide image: {t_export_end - t_export_start:.4f}s")
        print(f"Draw overlays:      {t_draw_end - t_draw_start:.4f}s")
        print(f"Save final image:   {t_save_end - t_save_start:.4f}s")
        print("---------------------")
        print(f"Total time:         {t_end - t_start:.4f}s")
        print("---------------------\n")
        
        print(f"✅ Highlighted slide image with rulers saved to: {output_path}")

        return output_path
        
    def _draw_rulers(self, canvas, border, width, height, scale_x, scale_y, tick_interval=25):
        """
        Draws X and Y rulers on the canvas with finer ticks and endpoint markers.

        Args:
            canvas: The image canvas to draw on.
            border (int): The border size around the slide image.
            width (int): The pixel width of the slide image.
            height (int): The pixel height of the slide image.
            scale_x (float): The scale factor for the x-axis (pixels per point).
            scale_y (float): The scale factor for the y-axis (pixels per point).
            tick_interval (int): The interval for ruler ticks in points.
        """
        font = cv2.FONT_HERSHEY_SIMPLEX
        font_scale = 0.4
        font_thickness = 1
        tick_length_major = 8
        tick_length_minor = 4
        
        # --- Helper function to draw a single tick ---
        def draw_tick(pos, value_text, is_major=True, is_x_axis=True):
            tick_len = tick_length_major if is_major else tick_length_minor
            text_size = cv2.getTextSize(value_text, font, font_scale, font_thickness)[0]
            
            if is_x_axis:
                cv2.line(canvas, (pos, border), (pos, border - tick_len), (0,0,0), 1)
                if is_major:
                    cv2.putText(canvas, value_text, (pos - text_size[0] // 2, border - tick_len - 5), font, font_scale, (0,0,0), font_thickness)
            else: # Y-axis
                cv2.line(canvas, (border, pos), (border - tick_len, pos), (0,0,0), 1)
                if is_major:
                    cv2.putText(canvas, value_text, (border - tick_len - text_size[0] - 5, pos + text_size[1] // 2), font, font_scale, (0,0,0), font_thickness)

        # --- X-axis Ruler (Top) ---
        cv2.line(canvas, (border, border), (border + width, border), (0,0,0), 1)
        for i in range(0, int(self.slide_width_points), tick_interval):
            px = int(i * scale_x) + border
            is_major_tick = (i % (tick_interval * 2) == 0) # Make every other tick major
            draw_tick(px, str(i), is_major=is_major_tick, is_x_axis=True)
        # Mark the extreme end value for X-axis
        end_x_px = int(self.slide_width_points * scale_x) + border
        draw_tick(end_x_px, str(int(self.slide_width_points)), is_major=True, is_x_axis=True)

        # --- Y-axis Ruler (Left) ---
        cv2.line(canvas, (border, border), (border, border + height), (0,0,0), 1)
        for i in range(0, int(self.slide_height_points), tick_interval):
            py = int(i * scale_y) + border
            is_major_tick = (i % (tick_interval * 2) == 0) # Make every other tick major
            draw_tick(py, str(i), is_major=is_major_tick, is_x_axis=False)
        # Mark the extreme end value for Y-axis
        end_y_py = int(self.slide_height_points * scale_y) + border
        draw_tick(end_y_py, str(int(self.slide_height_points)), is_major=True, is_x_axis=False)

def test_visualizer():
    """
    Tests the SlideVisualizer functionality by creating both a full-resolution
    highlighted image and a downsampled highlighted image.
    """
    print("\n🚀 Testing Slide Visualizer...")
    print("="*50)
    print("Ensure a PowerPoint presentation is open.")
    
    try:
        visualizer = SlideVisualizer()

        # --- Test 1: Full-resolution highlighted image ---
        print("\n🖼️  Test 1: Generating full-resolution highlighted image...")
        full_res_output = visualizer.create_highlighted_slide_image()
        if full_res_output and os.path.exists(full_res_output):
            print(f"🎉 Full-resolution test successful! Check file: {full_res_output}")
        else:
            print("❌ Full-resolution test failed.")
        
        # --- Test 2: Downsampled highlighted image ---
        print("\n🖼️  Test 2: Generating downsampled highlighted image...")
        downsampled_img = visualizer.get_downsampled_slide_image(target_width=512)
        if downsampled_img is not None:
            output_file = "downsampled_slide_preview.jpg"
            cv2.imwrite(output_file, downsampled_img)
            print(f"🎉 Downsampling test successful! Check file: {output_file}")
            # os.startfile(output_file)
        else:
            print("❌ Downsampling test failed.")
            
    except ConnectionError as e:
        print(f"❌ {e}")
        print("Please make sure PowerPoint is running with an open presentation.")
    except Exception as e:
        print(f"❌ An unexpected error occurred during testing: {e}")
    
    print("\n✅ All tests completed!")
    print("="*50)


if __name__ == "__main__":
    test_visualizer() 