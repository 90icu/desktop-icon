import base64
import io
from ttkbootstrap.icons import Icon
from PIL import Image

def extract_icon():
    try:
        # Decode base64 PNG data
        png_data = base64.b64decode(Icon.icon)
        
        # Open as Image
        img = Image.open(io.BytesIO(png_data))
        
        # Save as ICO (containing multiple sizes for better scaling)
        img.save("app.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
        print("Successfully extracted icon to app.ico")
    except Exception as e:
        print(f"Error extracting icon: {e}")

if __name__ == "__main__":
    extract_icon()
