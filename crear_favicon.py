from PIL import Image

# Cambiamos el nombre al que se ve físicamente en tu carpeta
img = Image.open("static/img/Favicon.png.ico").convert("RGBA")
img.thumbnail((220, 220))

canvas = Image.new("RGBA", (256, 256), (255, 255, 255, 0))

x = (256 - img.width) // 2
y = (256 - img.height) // 2

canvas.paste(img, (x, y), img)
canvas.save("static/img/favicon.png")

print("Favicon convertido a PNG correctamente")
