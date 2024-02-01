

from loaders.pptx_loader import PPTXLoader
from slides_builder.slides_builder import SlidesBuilder
import json


ENGAGELI_PREFIX = "engageli-op"

SINGLE_CHOICE_POLL_QR_DATA = f"{ENGAGELI_PREFIX}://{{'op': 'PBM'}}"

def main():
    print("Hello World!")
    builder = SlidesBuilder("../tests/inputs/test2.pptx")
    
    qr_data = SINGLE_CHOICE_POLL_QR_DATA
    text_dict = {
        0: "Drinking diet coke",
        1: "Eating Haagen-Dazs ice cream",
        2: "Smoking cigarettes and drinking alcohol",
        3: "Don't organize the kitchen as you were told to do"
    }
    
    builder.add_qr_slide(qr_data=qr_data, text_dict=text_dict, slide_title="What is more dangerous?")
    slide = builder.add_full_slide_image(image_path="../tests/inputs/logo.png", slide_index=2)
    builder.append_notes_to_slide(slide=slide, notes=json.dumps(text_dict))
    
    notes = PPTXLoader.extract_text_from_notes_between_delimiters(slide=slide)
    print(json.loads(notes[0]))
    
    builder.save("../tests/outputs/test2.pptx")
    


    
if __name__ == "__main__":
    main()