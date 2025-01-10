from PIL import Image

# Open the PNG file
img = Image.open("icon2.png")

# Convert and save as ICO
img.save("icon2.ico", format="ICO", sizes=[(256, 256)])  # Recommended size: 256x256
