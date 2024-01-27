from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from io import BytesIO
from content_generators.qr_generator import QRGenerator
# import qrcode

# class QRGenerator:
#     @staticmethod
#     def generate_qr_code(data: str):
#         img = qrcode.make(data)
#         return img

class SlidesBuilder:
    def __init__(self, pptx_file):
        """
        Initialize the SlidesBuilder with a PowerPoint file.

        :param pptx_file: Path to the PowerPoint file
        """
        # Should auto apply the presentation layout to the new slides added
        self.presentation = Presentation(pptx_file)
        
    def inspect_slide_layouts(self):
        """
        Inspect the slide layouts in the presentation.
        
        :return: Information about the slide layouts in the presentation
        """
        layouts_info = {}

        for i, layout in enumerate(self.presentation.slide_layouts):
            layout_info = {'name': layout.name, 'placeholders': []}
            for placeholder in layout.placeholders:
                layout_info['placeholders'].append({
                    'index': placeholder.placeholder_format.idx,
                    'type': placeholder.placeholder_format.type
                })
            layouts_info[i] = layout_info

        return layouts_info

    def add_slide(self, layout_idx=1, title="New Slide", content=""):
        """
        Add a new slide to the presentation with a title and content.

        :param title: Title of the new slide
        :param content: Content of the new slide
        :return: The new slide
        """
        # Use the first slide layout by default (typically a title slide)
        slide_layout = self.presentation.slide_layouts[layout_idx]

        # Add a slide
        slide = self.presentation.slides.add_slide(slide_layout)

        # Set the title and content if placeholders are available
        if slide.shapes.title:
            slide.shapes.title.text = title

        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1:
                shape.text = content
                
        return slide

    def add_slide_with_title_and_content(self, title, content):
        """
        Add a slide with a title at the top and content following.

        :param title: Title of the slide
        :param content: Content of the slide
        :return: The new slide
        """
        slide_layout = self.presentation.slide_layouts[1]  # A layout with title and content placeholders

        slide = self.presentation.slides.add_slide(slide_layout)
        
        # Assign the title and content if the placeholders exist
        title_placeholder = slide.shapes.title
        content_placeholder = None

        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.type == "BODY":
                content_placeholder = placeholder
                break

        if title_placeholder is not None:
            title_placeholder.text = title
        else:
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = title
            for paragraph in title_frame.paragraphs:
                paragraph.font.size = Pt(44)

        if content_placeholder is not None:
            content_placeholder.text = content
        else:
            content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
            content_box.text_frame.text = content
            
        return slide
                
    def add_qr_slide(self, qr_data, text_dict, slide_title, slide_content=""):
        """
        Add a slide with specified number of QR codes containing the given data
        and corresponding text next to each QR code.

        :param qr_data: Data to be encoded in the QR codes
        :param text_dict: Dictionary with index and text for each QR code
        :param slide_title: Title of the new slide
        :param slide_content: Content of the new slide
        """
        slide = self.add_slide_with_title_and_content(slide_title, slide_content)

        # Define size and spacing
        qr_size = Inches(0.6)  # Size of each QR code
        qr_margin = Inches(0.2)  # Margin between QR codes
        max_qr = min(len(text_dict), 4)  # Limit the number of QR codes to 4 or the number of items in text_dict

        # Assuming QR codes are placed in the content area
        content_top = Inches(1.5)  # Starting position for QR codes
        content_left = Inches(0.5)  # Left margin for QR codes

        for i in range(max_qr):
            # Calculate position for QR code
            top = content_top + i * (qr_size + qr_margin)

            # Generate and add QR code to slide
            qr_img = QRGenerator.generate_qr_code(qr_data)
            qr_bytes = BytesIO()
            qr_img.save(qr_bytes)
            qr_bytes.seek(0)

            slide.shapes.add_picture(qr_bytes, content_left, top, qr_size, qr_size)

            # Add a textbox for each QR code
            if i in text_dict:
                textbox = slide.shapes.add_textbox(content_left + qr_size + qr_margin, top,
                                                   Inches(3), qr_size)
                textbox.text = text_dict[i]

    def save(self, output_file):
        """
        Save the updated presentation to a new file.

        :param output_file: Path to save the updated PowerPoint file
        """
        self.presentation.save(output_file)
    