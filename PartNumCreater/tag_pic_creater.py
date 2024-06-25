# import qrcode

from PIL import Image, ImageDraw
from brother_ql.conversion import convert
from brother_ql.backends.helpers import send
from brother_ql.raster import BrotherQLRaster

from PIL import Image, ImageDraw, ImageFont
import qrcode
import barcode
from barcode.writer import ImageWriter
from datetime import datetime

def draw_text(draw, text, position, font, max_width, align):
    # Initialize variables
    lines = []
    words = text.split()
    line = []

    # Create lines of text within the max_width constraint
    for word in words:
        # Test current line with the new word
        test_line = ' '.join(line + [word])
        width = draw.textlength(test_line, font=font)
        if width <= max_width:
            line.append(word)
        else:
            # If the line is too long, finalize the current line and start a new one
            lines.append(' '.join(line))
            line = [word]

    # Append the last line
    lines.append(' '.join(line))

    # Draw the lines on the image
    y = position[1]
    for line in lines:
        draw.text((position[0], y), line, font=font, fill="black", align=align)
        y += 30  # Move to the next line based on the font size

def draw_tag_sticker(pic_name: str, part_code: str, part_spec: str, part_manufacture: str, part_supplier: str, count: int, part_type: str, in_stock_date: str, create_date: str, stockplace: str):
    current_time = datetime.now()

    # Create a new image
    im = Image.new("L", (500, 500), color="white")
    g = ImageDraw.Draw(im)

    # Load fonts
    font_large = ImageFont.truetype('msjh.ttc', size=40)
    font_small = ImageFont.truetype('msjh.ttc', size=25)

    # Draw text with word wrapping
    draw_text(g, part_code, (10, 28), font_large, 280, "left")
    draw_text(g, part_spec, (10, 149), font_small, 250, "left")
    draw_text(g, "製造商:"+part_manufacture, (10, 229), font_small, 250, "left")
    draw_text(g, "供應商:"+str(part_supplier).replace('nan',''), (10, 264), font_small, 250, "left")
    draw_text(g, "入庫日期: "+in_stock_date, (10, 299), font_small, 250, "left")
    draw_text(g, "製表日期: "+create_date, (10, 334), font_small, 250, "left")
    draw_text(g, "物料倉: "+stockplace, (10, 369), font_small, 500, "left")
    draw_text(g, "數量: "+str(count), (322, 28), font_large, 250, "right")
    draw_text(g, part_type, (10, 92), font_small, 250, "right")

    from barcode_generator import placecode, stockroom
    qr = qrcode.make(part_code+in_stock_date.replace('-','')[2:]+str(stockroom[stockplace][1])+str(placecode[stockroom[stockplace][0]])+str(count))
    size = (280,280)
    qr_image = qr.resize(size)
    im.paste(qr_image, (260,82))

    number = part_code+in_stock_date.replace('-','')[2:]+str(stockroom[stockplace][1])+str(placecode[stockroom[stockplace][0]])+str(count)  # 要轉換的數字
    my_code = barcode.Code128(number, writer=ImageWriter())       # 轉換成 barcode
    my_code.save("bar")

    barcode_image = Image.open('bar.png')
    im.paste(barcode_image, (-10, 400))
    from os import remove
    remove('bar.png')

    # Show the image
    im.save(pic_name+".png")

# draw_tag_sticker("1C371043300", "SMT, 0603, X7R, K±10%, 1000P, 50V", "三環","",50, "電容/電容", "2024-01-01", "2024-01-02","竹北物料倉")