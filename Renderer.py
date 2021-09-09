from PIL import Image
import numpy as np
import xlwings as xw

wb = xw.Book(r'C:\Users\Saif Qadeer\PycharmProjects\Excel_Image_Renderer\Output\output.xlsx')
sht = wb.sheets['Sheet1']
Img = Image.open("test.jpg",'r')
Pixel_Color = list(Img.getdata())
s=(360,480,3)
Pix = np.zeros(s)
print(Pixel_Color)
dict = {0:'A',1:'B',2:'C',3:'D',4:'E',5:'F',6:'G',7:'H',8:'I',9:'J',10:'K',11:'L',12:'M',13:'N',14:'O',15:'P',16:'Q',
        17:'R',18:'S',19:'T',20:'U',21:'V',22:'W',23:'X',24:'Y',25:'Z'}


def GetCellAdd(rows,columns):
    quotienty = columns//26
    remaindery = columns%26
    if (quotienty==0):
        return "{0}{1}".format(dict[remaindery],str(rows+1))
    else:
        return "{0}{1}".format(dict[quotienty-1]+dict[remaindery], str(rows + 1))


for row in range(0,359):
        for col in range(0,479):
            for i in range (0,3):
                Pix[row, col, i]= Pixel_Color[row * 480 + col][i]

for x in range(0,479):
        for y in range(0,359):
            Cell = GetCellAdd(y,x)
            sht.range(Cell).color = (Pix[y, x, 0], Pix[y, x, 1], Pix[y, x, 2])


