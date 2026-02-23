from PIL import Image
import os

img_path = "C:/Users/elivanilso.junior/.gemini/antigravity/brain/3cae331d-1eff-435b-b0af-95cf7ffb6fab/app_icon_1769523898113.png"
if os.path.exists(img_path):
    img = Image.open(img_path)
    img.save("app_icon.ico", format="ICO", sizes=[(256, 256)])
    print("Icone convertido com sucesso: app_icon.ico")
else:
    print("Imagem nao encontrada.")
