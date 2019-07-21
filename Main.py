import xlsxwriter
import os, os.path
from tkinter import Tk
from tkinter.filedialog import askdirectory
import sys

# we don't want a full GUI, so keep the root window from appearing
Tk().withdraw()
# show an "Open" dialog box and return the path to the selected file
PATH = askdirectory()
if not PATH:
    sys.exit(0)
print(PATH)

'''
    Initial
'''

# Get py script file path
#PATH = os.path.dirname(os.path.realpath(__file__))
# Create a new workbook and add a worksheet
WORKBOOK = xlsxwriter.Workbook('Image_Lib.xlsx')
WORKSHEET = WORKBOOK.add_worksheet('Img_lib')
# Format the column & ROW
WORKSHEET.set_column('A:A', 100)
WORKSHEET.set_column('B:B', 50)
WORKSHEET.set_default_row(50)
ROW = 0
COL = 0

'''
    Work function
'''


def writeExcel(path):
    """
        path: image file path
    """
    global ROW, COL

    # Write image hyperlinks and move to next column
    WORKSHEET.write_url(row=ROW, col=COL, url=path, string=path)
    COL += 1

    # Insert an image and move to next row
    WORKSHEET.insert_image(row=ROW, col=COL, filename=path, options={'x_scale': 0.2, 'y_scale': 0.2,
                                                                     'url': 'external:'+path, 'tip': path})
    ROW += 1
    COL -= 1


def main():
    """
    :note:
        dirPath, dirNames, fileNames = os.walk(PATH)
        file, ext = os.path.splitext(filename)

    :return: n/a
    """
    # only search image
    valid_images = [".jpg", ".gif", ".png", ".bmp"]
    # Search include chosen directory and subdirectory
    for dirPath, dirNames, fileNames in os.walk(PATH):
        for f in fileNames:
            ext = os.path.splitext(f)[1]
            if ext.lower() not in valid_images:
                continue
            writeExcel(os.path.join(dirPath, f))

    # for f in os.listdir(PATH):
    #     ext = os.path.splitext(f)[1]
    #     if ext.lower() not in valid_images:
    #         continue
    #     writeExcel(os.path.join(PATH, f))


if __name__ == '__main__':
    main()
    WORKBOOK.close()
