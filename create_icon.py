"""
Script to create application icon
"""
from PIL import Image, ImageDraw, ImageFont

# Create icon image 256x256 (high quality)
size = 256
image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

# Draw gradient background circle
for i in range(size//2, 0, -1):
    alpha = int(255 * (i / (size//2)))
    color = (30, 144, 255, alpha)  # DodgerBlue
    draw.ellipse([size//2 - i, size//2 - i, size//2 + i, size//2 + i], fill=color)

# Draw inner circle (white)
inner_size = int(size * 0.7)
offset = (size - inner_size) // 2
draw.ellipse([offset, offset, offset + inner_size, offset + inner_size], fill=(255, 255, 255, 255))

# Draw connection symbol (arrows and dots)
# Serial port symbol (left)
left_x = size // 4
center_y = size // 2
# Draw plug shape
plug_width = 30
plug_height = 40
draw.rectangle([left_x - plug_width//2, center_y - plug_height//2, 
                left_x + plug_width//2, center_y + plug_height//2], 
               fill=(50, 50, 50, 255))
# Pins
for i in range(-1, 2):
    pin_y = center_y + i * 12
    draw.rectangle([left_x - plug_width//3, pin_y - 2,
                   left_x + plug_width//3, pin_y + 2],
                  fill=(200, 200, 200, 255))

# Arrow (middle)
arrow_y = center_y
arrow_start_x = left_x + plug_width
arrow_end_x = size // 2 + 10
# Line
draw.line([arrow_start_x, arrow_y, arrow_end_x, arrow_y], fill=(50, 50, 50, 255), width=4)
# Arrow head
arrow_size = 12
draw.polygon([
    (arrow_end_x, arrow_y),
    (arrow_end_x - arrow_size, arrow_y - arrow_size//2),
    (arrow_end_x - arrow_size, arrow_y + arrow_size//2)
], fill=(50, 50, 50, 255))

# Window symbol (right)
right_x = size * 3 // 4
window_size = 50
draw.rectangle([right_x - window_size//2, center_y - window_size//2,
               right_x + window_size//2, center_y + window_size//2],
              outline=(50, 50, 50, 255), width=4)
# Window title bar
draw.rectangle([right_x - window_size//2, center_y - window_size//2,
               right_x + window_size//2, center_y - window_size//2 + 10],
              fill=(50, 50, 50, 255))
# Window buttons
for i in range(3):
    btn_x = right_x + window_size//2 - 8 - i * 10
    btn_y = center_y - window_size//2 + 5
    draw.ellipse([btn_x - 3, btn_y - 3, btn_x + 3, btn_y + 3],
                fill=(255, 255, 255, 255))

# Save as ICO with multiple sizes
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
images = []
for icon_size in sizes:
    resized = image.resize(icon_size, Image.Resampling.LANCZOS)
    images.append(resized)

# Save as .ico
images[0].save('app_icon.ico', format='ICO', sizes=[(img.size[0], img.size[1]) for img in images])
print("✅ Icon created successfully: app_icon.ico")

# Also save as PNG for preview
image.save('app_icon.png', format='PNG')
print("✅ PNG preview created: app_icon.png")
