import base64

# Read the PNG file and encode it
with open("icon2.png", "rb") as icon_file:
    encoded_icon = base64.b64encode(icon_file.read()).decode("utf-8")

# Save the encoded string to a Python file
with open("icon_base64.py", "w") as encoded_file:
    encoded_file.write(f'encoded_icon = """{encoded_icon}"""\n')

print("Base64 icon saved to icon_base64.py")
