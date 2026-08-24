from PIL import Image

img = Image.open("static/img/favicon.png").convert("RGBA")

img.thumbnail((220, 220))

canvas = Image.new("RGBA", (256, 256), (255, 255, 255, 0))

x = (256 - img.width) // 2
y = (256 - img.height) // 2

canvas.paste(img, (x, y), img)

canvas.save("static/img/favicon_256.png")

print("Favicon creado correctamente")