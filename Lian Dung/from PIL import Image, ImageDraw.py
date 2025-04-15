from PIL import Image, ImageDraw

def draw_rectangle(image_path, coords, output_path):
    # Open the image
    image = Image.open(image_path)
    
    # Create a drawing object
    draw = ImageDraw.Draw(image)

    # Unpack coordinates
    top_left = coords["top_left"]
    bottom_right = coords["bottom_right"]

    # Draw the rectangle (outline in red)
    draw.rectangle([top_left, bottom_right], outline="red", width=3)

    # Save the image with the rectangle
    image.save(output_path)
    print(f"Image saved with rectangle at {output_path}")

# Coordinates for the rectangle (as calculated earlier)
rectangle_coords = {
    "top_left": (420, 50),
    "bottom_right": (830, 250),
}

# Input and output paths
input_image_path = "/mnt/data/image.png"
output_image_path = "/mnt/data/image_with_rectangle.png"

# Draw the rectangle and save the output
draw_rectangle(input_image_path, rectangle_coords, output_image_path)
