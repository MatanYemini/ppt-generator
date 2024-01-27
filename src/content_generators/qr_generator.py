import qrcode
from PIL import Image, ImageDraw, ImageFont

class QRGenerator:
    @staticmethod
    def generate_qr_code(data: str):
        img = qrcode.make(data)
        img2 = qrcode.make()
        return img

    @staticmethod
    def create_qr_with_text_and_logo(data, logo_path, text):
        # Generate QR code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=1,
        )
        qr.add_data(data)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')

        # Resize logo
        logo_size = 40  # Size of the logo
        logo = Image.open(logo_path)
        logo = logo.resize((logo_size, logo_size), Image.ANTIALIAS)

        # Calculate text size
        font_size = 20
        font = ImageFont.load_default()  # Loading default font
        draw = ImageDraw.Draw(qr_img)
        text_width, text_height = draw.textsize(text, font=font)

        # Calculate combined image size
        combined_height = qr_img.height + max(logo_size, text_height)
        combined_img = Image.new('RGB', (qr_img.width, combined_height), 'white')

        # Place QR code
        combined_img.paste(qr_img, (0, 0))

        # Place text and logo
        draw = ImageDraw.Draw(combined_img)
        text_x = (combined_img.width - text_width) // 2
        text_y = qr_img.height + (max(logo_size, text_height) - text_height) // 2
        draw.text((text_x, text_y), text, font=font, fill='black')

        # Place logo
        logo_x = text_x - logo_size - 10  # 10 pixels space between logo and text
        logo_y = qr_img.height + (max(logo_size, text_height) - logo_size) // 2
        combined_img.paste(logo, (logo_x, logo_y), mask=logo)

        return combined_img
    
# # Example usage
# data = "https://example.com"
# logo_path = "logo.png"
# text = "Single Choice Poll"
# styled_qr = QRGenerator.create_qr_with_text_and_logo(data, logo_path, text)
# styled_qr.save("testqr.png")  # Display the image