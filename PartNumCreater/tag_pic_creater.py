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

import barcode_img_create
import os

def draw_text(draw, text, position, font, max_width, align):
    # Initialize variables
    lines = []
    words = text
    line = []

    # Create lines of text within the max_width constraint
    for word in words:
        # Test current line with the new word
        test_line = ''.join(line + [word])
        width = draw.textlength(test_line, font=font)
        if width <= max_width:
            line.append(word)
        else:
            # If the line is too long, finalize the current line and start a new one
            lines.append(''.join(line))
            line = [word]

    # Append the last line
    lines.append(''.join(line))

    # Draw the lines on the image
    y = position[1]
    for line in lines:
        draw.text((position[0], y), line, font=font, fill="black", align=align)
        y += 30  # Move to the next line based on the font size


def draw_tag_sticker(pic_name: str, part_code: str, part_spec: str, part_manufacture: str, part_supplier: str, count: int, part_type: str, in_stock_date: str, create_date: str, stockplace: str, pn: str, box_num: int, batch_number: str, stcok_org: str):
    from barcode_generator import placecode, stockroom
    inventoryID = part_code+in_stock_date.replace('-','')[2:]+str(placecode[stockroom[stockplace][0]])+str(stockroom[stockplace][1])  # 要轉換的數字
    from execl_handle import number_to_letter
    boxID = inventoryID+number_to_letter(box_num)

    # Create a new image
    im = Image.new("L", (500, 500), color="white")
    g = ImageDraw.Draw(im)

    # Load fonts
    # https://www.sharpgan.com/solve-pyinstaller-cannot-recognize-static-file/
    import sys, os
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    path_to_ttf = os.path.abspath(os.path.join(bundle_dir, 'msjh.ttc'))
    print(path_to_ttf)

    font_large = ImageFont.truetype(path_to_ttf, size=40)
    font_small = ImageFont.truetype(path_to_ttf, size=20)

    
    qr = qrcode.make(boxID)
    size = (200,200)
    qr_image = qr.resize(size)
    im.paste(qr_image, (320,122))

    # Draw text with word wrapping
    draw_text(g, part_code, (10, 35), font_large, 280, "left")
    draw_text(g, part_spec, (10, 159), font_small, 300, "left")
    # draw_text(g, "曜璿東科技股份有限公司 Orient SunTech", (45, 5), font_small, 480, "left")
    draw_text(g, "製造商:"+part_manufacture, (10, 309), font_small, 480, "left")
    draw_text(g, "供應商:"+str(part_supplier).replace('nan',''), (10, 339), font_small, 500, "left")
    # draw_text(g, "入庫日期:"+in_stock_date, (10, 229), font_small, 250, "left")
    draw_text(g, "入庫日期:"+in_stock_date, (10, 249), font_small, 250, "left")
    draw_text(g, "原始倉:"+stcok_org, (10, 279), font_small, 250, "left")
    draw_text(g, "供應商批號: "+batch_number, (10, 369), font_small, 480, "left")
    draw_text(g, "箱號:"+str(box_num), (342, 369), font_small, 250, "right")
    draw_text(g, boxID[11:], (342, 299), font_small, 250, "right")
    draw_text(g, pn+"  "+part_type, (10, 92), font_small, 480, "right")

    
    # code128 = Code128(number, writer=ImageWriter())
    # barcode_image = code128.render(writer_options={"font_size": 1, "text_distance": 1})
    # barcode_file_path = os.path.join('./', "bar.png")
    # barcode_image.save(barcode_file_path, "PNG")
    # EAN = barcode.get_barcode_class('code128')
    # my_ean = EAN(number)
    # my_ean.save('bar')
    # my_ean = EAN(number, writer=ImageWriter(format='png'))
    # my_ean.save('./bar')
    # my_code = barcode.Code128(number, writer=ImageWriter())       # 轉換成 barcode
    # my_code.save("bar")

    if not os.path.exists('sticker'):
        os.makedirs('sticker')
        

    # barcode_img_create.generate_code128_barcode(boxID, "sticker/bar.png")
    # print(boxID)
    Code129 = barcode.get('code128', boxID, writer=ImageWriter())
    Code129.save('sticker/bar', options={'format': 'PNG'})


    barcode_image = Image.open('sticker/bar.png')
    new_size = (450, 200)
    img_resized = barcode_image.resize(new_size, Image.LANCZOS)
    im.paste(img_resized, (0, 400))
    from os import remove
    remove('sticker/bar.png')


    # company_logo_image = Image.open('sticker/Logo.png')
    # new_size = (32, 30)
    # img_resized = company_logo_image.resize(new_size, Image.LANCZOS)

    # im.paste(img_resized, (10, 5))

    company_logo_image = Image.open('sticker/Company_Logo.bmp')
    new_size = (175, 30)
    img_resized = company_logo_image.resize(new_size, Image.LANCZOS)

    im.paste(img_resized, (10, 5))
    
    # Show the image
    im.save(f"{pic_name}.png")

    return [inventoryID, box_num, count, batch_number]
 

# draw_tag_sticker("1C371043300", "SMT, 0603, X7R, K±10%, 1000P, 50V", "三環","",50, "電容/電容", "2024-01-01", "2024-01-02","竹北物料倉")
