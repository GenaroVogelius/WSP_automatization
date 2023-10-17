
import os
from pathlib import Path


from io import BytesIO
import win32clipboard
from PIL import Image


def send_to_clipboard(image):
    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

BASE_DIR = Path(__file__).resolve().parent

filepath = os.path.join(BASE_DIR, "gondolazo.jpeg")
image = Image.open(filepath)


send_to_clipboard(image)

