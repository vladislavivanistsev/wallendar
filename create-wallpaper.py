from PIL import Image, ImageDraw, ImageFont
from google.colab import files  # Import the files module for downloading

# Define wallpaper dimensions and icon grid
wallpaper_width = 1920
wallpaper_height = 1080
icon_columns = 20
icon_rows = 10
icon_width = wallpaper_width // icon_columns
icon_height = wallpaper_height // icon_rows

# Create a blank image for the wallpaper
wallpaper = Image.new('RGB', (wallpaper_width, wallpaper_height), color='white')
draw = ImageDraw.Draw(wallpaper)

# Define area colors and titles
areas = [
    {'title': 'Files', 'color': '#C2A5CF', 'start_col': 'A', 'start_row': 1, 'end_col': 'T', 'end_row': 10},
    {'title': 'ToDo', 'color': '#FDB366', 'start_col': 'D', 'start_row': 1, 'end_col': 'L', 'end_row': 5},
    {'title': 'Sort', 'color': '#ACD39E', 'start_col': 'M', 'start_row': 1, 'end_col': 'T', 'end_row': 5},  # Changed color
    {'title': 'Check', 'color': '#FEDA8B', 'start_col': 'D', 'start_row': 6, 'end_col': 'L', 'end_row': 10},  # Changed color
    {'title': 'Later', 'color': '#99DDFF', 'start_col': 'M', 'start_row': 6, 'end_col': 'T', 'end_row': 10},
]

# Load Liberation Sans font
liberation_font_path = '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'  # Update the path accordingly
liberation_font = ImageFont.truetype(liberation_font_path, size=40)

# Loop through each area and draw it on the wallpaper
for area in areas:
    # Calculate area coordinates based on grid positions
    area_x1 = (ord(area['start_col']) - ord('A')) * icon_width
    area_y1 = (area['start_row'] - 1) * icon_height
    area_x2 = (ord(area['end_col']) - ord('A') + 1) * icon_width
    area_y2 = area['end_row'] * icon_height

    # Draw area background with specified color
    draw.rectangle([area_x1, area_y1, area_x2, area_y2], fill=area['color'])

    # Draw area title
    text_bbox = draw.textbbox((area_x1, area_y1, area_x2, area_y2), area['title'], font=liberation_font)
    text_width = text_bbox[2] - text_bbox[0]  # Calculate width from bounding box
    text_height = text_bbox[3] - text_bbox[1]  # Calculate height from bounding box
    title_x = area_x1 + 40  # 40px padding from left edge
    title_y = area_y1 + 40  # 40px padding from top edge
    draw.text((title_x, title_y), area['title'], fill='white', font=liberation_font)

# Save the wallpaper image
wallpaper.save('wallpaper.png')
