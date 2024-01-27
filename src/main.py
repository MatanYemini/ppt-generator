

from slides_builder.slides_builder import SlidesBuilder


ENGAGELI_PREFIX = "engageli-op"

SINGLE_CHOICE_POLL_QR_DATA = f"{ENGAGELI_PREFIX}://{{'op': 'PBM'}}"

def main():
    print("Hello World!")
    builder = SlidesBuilder("test2.pptx")
    layouts = builder.inspect_slide_layouts()
    qr_data = SINGLE_CHOICE_POLL_QR_DATA
    text_dict = {
        0: "Drinking diet coke",
        1: "Eating Häagen-Dazs ice cream",
        2: "Smoking cigarettes and drinking alcohol",
        3: "Don't organize the kitchen as you were told to do"
    }
    builder.add_qr_slide(qr_data, text_dict, "What is more dangerous?")
    builder.save("test2-res.pptx")
    
if __name__ == "__main__":
    main()