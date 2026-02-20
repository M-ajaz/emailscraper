from PIL import Image, ImageDraw, ImageFont
import os

# Create 256x256 icon
size = 256
img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Dark background circle
draw.ellipse([8, 8, 248, 248], fill="#0f1117")

# Cyan accent circle border
draw.ellipse([8, 8, 248, 248], outline="#06b6d4", width=8)

# Email envelope shape in white
draw.rectangle([60, 90, 196, 166], fill="#ffffff", outline="#06b6d4", width=3)
draw.polygon([(60,90),(128,140),(196,90)], fill="#06b6d4")

# Save as PNG
output = os.path.join(os.path.dirname(__file__), "icon.png")
img.save(output, "PNG")
print(f"Saved icon to {output}")
