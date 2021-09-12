from PIL import Image
import numpy as np
import xlwings as xw
from Configuration import *


def CellColumnString(columns):
    # convert the column index of the image array to a corresponding string. It is needed to convert the indexes to
    # corresponding column address in Excel
    if columns < 26:
        return "{0}".format(ALPHABETS[columns])
    else:
        return CellColumnString((columns // 26) - 1) + ALPHABETS[columns % 26]


def GetCellAddress(rows, columns):
    # Generate cell adress for a particular element in array
    return CellColumnString(columns) + str(rows + 1)


def ImageToArray(PathToImage):
    # Load the image into memory and read the pixel values into a numpy array
    Img = Image.open(PathToImage, 'r')
    Pixel_Color = list(Img.getdata())
    # read value of image resolution
    (width, height) = Img.size
    # make 3D array in order to store 3 color values corresponding to each pixel
    PixelArray = np.zeros((height, width, 3))
    for row in range(0, height - 1):
        for col in range(0, width - 1):
            for i in range(0, 3):
                PixelArray[row, col, i] = Pixel_Color[row * width + col][i]

    return PixelArray


def Render(ColorArray, PathToOutput):
    # read the color values from color array (formed in previous function) and fill the corresponding cells in the target
    # Excel file
    file = xw.Book(PathToOutput)
    Sheet = file.sheets['Sheet1']
    for x in range(0, ColorArray.shape[0] - 1):
        for y in range(0, ColorArray.shape[1] - 1):
            Cell = GetCellAddress(x, y)
            Sheet.range(Cell).color = (ColorArray[x, y, 0], ColorArray[x, y, 1], ColorArray[x, y, 2])


if __name__ == "__main__":
    Render(ImageToArray(PATH_TO_INPUT), PATH_TO_OUTPUT)
