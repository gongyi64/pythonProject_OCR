import pyocr
tools = pyocr.get_available_tools()
tool = tools[0]
builder = pyocr.builders.TextBuilder(tesseract_layout=6)
tool.image_to_string("読み取り対象", lang="読み取り対象言語", builder="オプション")


import os
from PIL import Image
import pyocr
 
#インストールしたTesseract-OCRのパスを環境変数「PATH」へ追記する。
#OS自体に設定してあれば以下の2行は不要
path='C:\\Users\\406239\\AppData\\Local\\Tesseract-OCR'
os.environ['PATH'] = os.environ['PATH'] + path
 
#pyocrへ利用するOCRエンジンをTesseractに指定する。
tools = pyocr.get_available_tools()
tool = tools[0]
 
#OCR対象の画像ファイルを読み込む
img = Image.open("card_image/ocr_test.png")
 
#画像から文字を読み込む
builder = pyocr.builders.TextBuilder(tesseract_layout=6)
text = tool.image_to_string(img, lang="jpn", builder=builder)
 
print(text)