from PIL import Image

img = Image.open("assets/logodocflow.png")
img.save("assets/logodocflow.ico", format="ICO", sizes=[(512, 512)])
