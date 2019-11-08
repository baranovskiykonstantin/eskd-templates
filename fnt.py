#! /usr/bin/python
import tkinter
import tkinter.font

import codecs

root = tkinter.Tk()

mm_per_pix = 1 / root.winfo_fpixels('1m')

font = tkinter.font.Font(family="OpenGost Type B TT", size=100)

chars = list(range(0x0020, 0x007f))
chars.append(0x00ab)
chars.append(0x00b0)
chars.append(0x00b1)
chars.append(0x00b5)
chars.append(0x00b7)
chars.append(0x00bb)
chars.extend(range(0x0410, 0x0450))
chars.append(0x0401)
chars.append(0x0404)
chars.append(0x0406)
chars.append(0x0407)
chars.append(0x0451)
chars.append(0x0454)
chars.append(0x0456)
chars.append(0x0457)
chars.append(0x2015)
chars.append(0x2026)
chars.append(0x2103)
chars.append(0x2109)
chars.append(0x2116)

out = codecs.open("fnt.txt", 'w', "utf-8")

minWidth = 1e6
maxWidth = 0
out.write("CHARWIDTH_MM_PER_POINT = {\n")
out.write("    '\\r': 0,\n")
out.write("    '\\n': 0,\n")
for char in chars:
    width = font.measure(chr(char)) * mm_per_pix / 100
    if minWidth > width:
        minWidth = width
    if maxWidth < width:
        maxWidth = width
    if char in (0x0027, 0x005c):
        out.write("    '\\{}': {},\n".format(chr(char), width))
    else:
        out.write("    '{}': {},\n".format(chr(char), width))
out.write('    "min": {},\n'.format(minWidth))
out.write('    "max": {},\n'.format(maxWidth))
out.write('}\n')
out.close()
root.destroy()
