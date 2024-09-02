import usb
d= usb.core.find(find_all=1,)
print(d)
print(list(d))

print(usb.core.show_devices())


import usb.backend.libusb1
backend = usb.backend.libusb1.get_backend(find_library=lambda x: r"C:\Users\j1618\AppData\Local\Programs\Python\Python312\Lib\site-packages\libusb\_platform\_windows\x64\libusb-1.0.dll")  # adapt to your path
usb_devices = usb.core.find(backend=backend, find_all=True, idVendor=0x04f9)
for usb_device in usb_devices:
    print(usb_device)