import json
from pptx import Presentation
from PIL import Image
import io
import pytesseract


def extract_text_from_slide(slide):
    """
    Extract the text from a slide.
    
    :param slide: Slide to extract text from
    :return: Text from the slide
    """
    text = ""
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            text += shape.text + " "
    return text.strip()

def interpret_image(image):
    """
    Interpret an image and return a description of it.
    
    :param image: Image to interpret
    :return: Description of the image
    """
    
    # Convert the image to a format suitable for text recognition
    image = image.convert('L')  # Convert to grayscale

    # Use Tesseract to do OCR on the image
    text = pytesseract.image_to_string(image)

    return text

def process_pptx(pptx_file):
    """
    Prosess a PowerPoint presentation and return a JSON string with the content of each slide.
    
    :param pptx_file: Path to the PowerPoint file
    :return: JSON string with the content of each slide
    """
    presentation = Presentation(pptx_file)
    slides_content = {}

    for i, slide in enumerate(presentation.slides):
        slide_content = {"text": extract_text_from_slide(slide), "images": []}

        for shape in slide.shapes:
            if shape.shape_type == 13:  # This is the type for Picture
                image_stream = io.BytesIO(shape.image.blob)
                image = Image.open(image_stream)
                image_description = interpret_image(image)
                slide_content["images"].append(image_description)

        slides_content[f"Slide {i + 1}"] = slide_content

    return json.dumps(slides_content, indent=4)
