import openpyxl
from openpyxl_image_loader import SheetImageLoader
import pandas as pd
#loading the Excel File and the sheet
pxl_doc = openpyxl.load_workbook('myfile.xlsx')
sheet = pxl_doc['Sheet1']

#calling the image_loader
image_loader = SheetImageLoader(sheet)

#get the image (put the cell you need instead of 'A1')


sheet=pd.read_excel('myfile.xlsx',sheet_name='Sheet1')
# print(sheet)
# roll=(sheet['ROLL NO.'][0])

#showing the image
# image.show()

#saving the image
# roll=str(roll)
# image.save('my_path/'+roll+'.jpg')

# image=image_loader.get()
# image.show()  
for sr in sheet['S.NO.']:
    try:
        print(sr)
        print(sheet['ROLL NO.'][sr-1])
        image=image_loader.get('F'+str(sr+1))
        image.save('my_path/'+str(sheet['ROLL NO.'][sr-1])+'.jpg')
    except:
        continue
