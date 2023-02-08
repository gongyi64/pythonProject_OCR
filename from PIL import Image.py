# ocr_card.py
import os
from PIL import Image
import pyocr
import pyocr.builders

# 1.インストール済みのTesseractのパスを通す
path_tesseract = "C:\\Users\\406239\\AppData\\Local\\Tesseract-OCR"
if path_tesseract not in os.environ["PATH"].split(os.pathsep):
    os.environ["PATH"] += os.pathsep + path_tesseract

# 2.OCRエンジンの取得
tools = pyocr.get_available_tools()
tool = tools[0]

# 3.原稿画像の読み込み
img_org = Image.open("./card_image/IMG_8822.jpg")

# 4.ＯＣＲ実行
builder = pyocr.builders.TextBuilder()
result = tool.image_to_string(img_org, lang="jpn", builder=builder)

print(result)