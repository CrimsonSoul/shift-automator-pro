from PIL import Image, ImageDraw

def convert_to_ico(input_path, output_path):
    img = Image.open(input_path).convert("RGBA")
    
    # Create a circular mask
    mask = Image.new("L", img.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0) + img.size, fill=255)
    
    # Apply mask
    img.putalpha(mask)
    
    # Save as .ico with multiple sizes for Windows
    img.save(output_path, format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"Successfully converted {input_path} to {output_path}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        convert_to_ico(sys.argv[1], sys.argv[2])
    else:
        # Default for this project
        convert_to_ico("icon.png", "icon.ico")
