from PIL import Image, ImageDraw
from brother_ql.conversion import convert
from brother_ql.backends.helpers import send
from brother_ql.raster import BrotherQLRaster


def tag_printer(pic: str, ):
    label_images = []

    im = Image.open(pic)
    label_images.append(im)

    # Setting Printer Specificiations
    backend = 'pyusb'    # 'pyusb', 'linux_kernal', 'network'
    model = 'QL-1100' 
    printer = 'usb://0x04f9:0x20a7'  

    qlr = BrotherQLRaster(model)
    qlr.exception_on_warning = True

    # Converting print instructions for the Brother printer
    instructions = convert(
            qlr=qlr, 
            images=label_images,    #  Takes a list of file names or PIL objects.
            label='50', 
            rotate='0',    # 'Auto', '0', '90', '270'
            threshold=70.0,    # Black and white threshold in percent.
            dither=False, 
            compress=False, 
            dpi_600=False, 
            hq=True,    # False for low quality.
            cut=True
    )

    send(instructions=instructions, printer_identifier=printer, backend_identifier=backend, blocking=True)