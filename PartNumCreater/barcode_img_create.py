
from PIL import Image, ImageDraw, ImageFont

# Code128 字符與條形碼模式映射
code128 = {
    ' ': '11011001100', '!': '11001101100', '"': '11001100110', '#': '10010011000',
    '$': '10010001100', '%': '10001001100', '&': '10011001000', "'": '10011000100',
    '(': '10001100100', ')': '11001001000', '*': '11001000100', '+': '11000100100',
    ',': '10110011100', '-': '10011011100', '.': '10011001110', '/': '10111001100',
    '0': '10011101100', '1': '10011100110', '2': '11001110010', '3': '11001011100',
    '4': '11001001110', '5': '11011100100', '6': '11001110100', '7': '11101101110',
    '8': '11101001100', '9': '11100101100', ':': '11100100110', ';': '11101100100',
    '<': '11100110100', '=': '11100110010', '>': '11011011000', '?': '11011000110',
    '@': '11000110110', 'A': '10100011000', 'B': '10001011000', 'C': '10001000110',
    'D': '10110001000', 'E': '10001101000', 'F': '10001100010', 'G': '11010001000',
    'H': '11000101000', 'I': '11000100010', 'J': '10110111000', 'K': '10110001110',
    'L': '10001101110', 'M': '10111011000', 'N': '10111000110', 'O': '10001110110',
    'P': '11101110110', 'Q': '11010001110', 'R': '11000101110', 'S': '11011101000',
    'T': '11011100010', 'U': '11011101110', 'V': '11101011000', 'W': '11101000110',
    'X': '11100010110', 'Y': '11101101000', 'Z': '11101100010', '[': '11100011010',
    '\\': '11101111010', ']': '11001000010', '^': '11110001010', '_': '10100110000',
    '`': '10100001100', 'a': '10010110000', 'b': '10010000110', 'c': '10000101100',
    'd': '10000100110', 'e': '10110010000', 'f': '10110000100', 'g': '10011010000',
    'h': '10011000010', 'i': '10000110100', 'j': '10000110010', 'k': '11000010010',
    'l': '11001010000', 'm': '11110111010', 'n': '11000010100', 'o': '10001111010',
    'p': '10100111100', 'q': '10010111100', 'r': '10010011110', 's': '10111100100',
    't': '10011110100', 'u': '10011110010', 'v': '11110100100', 'w': '11110010100',
    'x': '11110010010', 'y': '11011011110', 'z': '11011110110', '{': '11110110110',
    '|': '10101111000', '}': '10100011110', '~': '10001011110', 'DEL': '10111101000',
    'FNC3': '10111100010', 'FNC2': '11110101000', 'SHIFT': '11110100010', 'Code C': '10111011110',
    'FNC4': '10111101110', 'FNC1': '11101011110', 'Start Code A': '11010000100', 'Start Code B': '11010010000',
    'Start Code C': '11010011100', 'Stop': '11000111010'
}

def generate_code128_barcode(data, output_path='code128_barcode.png'):
    # 基本設置
    bar_width = 2
    bar_height = 150
    padding = 5

    # 生成條形碼
    barcode = '11010010000'  # Start Code B
    for char in data:
        barcode += code128[char]
    barcode += '1100011101011'  # Stop

    # 計算圖片寬度
    barcode_width = len(barcode) * bar_width + padding * 2

    # 創建圖片對象
    image = Image.new('RGB', (barcode_width, bar_height + 30), 'white')
    draw = ImageDraw.Draw(image)

    # 繪製條形碼
    x = padding
    for bit in barcode:
        if bit == '1':
            draw.rectangle([(x, padding), (x + bar_width - 1, bar_height)], fill='black')
        x += bar_width

    # 在條形碼下方添加數字
    # font = ImageFont.load_default()
    # _,_,text_width, text_height = draw.textbbox((0, 0), text=data, font=font)
    # text_x = (barcode_width - text_width) / 2
    # text_y = bar_height + padding
    # draw.text((text_x, text_y), data, fill='black', font=font)

    # 保存條形碼圖片
    image.save(output_path)


# 測試生成條形碼
# generate_code128_barcode('Example text', 'code128_barcode.png')