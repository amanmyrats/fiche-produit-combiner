import win32com.client       # Need pywin32 from pip
from PIL import ImageGrab    # Need PIL as well
import os
import datetime
import openpyxl
import string
import general_functions
import pyperclip

from ctypes import windll

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from pathlib import Path

import image_extracter_with_openpyxl_class


class FP_Image_Extractor:
    def __init__(self,*args, **kwargs):
        self.start_time=datetime.datetime.now()
        self.xpath=Path(kwargs['xpath'])
        self.result_path=Path(kwargs['result_path'])
        self.images_folder='images'
        self.letters=string.ascii_uppercase

        check_directory=self.result_path / self.images_folder
        if not check_directory.exists():
            check_directory.mkdir(parents=True)

    def xlsx_image_extracter(self, **kwargs):
        self.workbook=kwargs['workbook']
        self.sheetname=kwargs['sheetname']
        self.image_name=kwargs['image_name']
        self.fp_no=kwargs['fp_no']
        self.excel_name=kwargs['excel_name']

        self.sheet=self.workbook[self.sheetname]
        print ("Extracting images from ", self.excel_name , " - ", self.sheet.title)
    
        try:
            self.images=SheetImageLoader(self.sheet)
        except:
            pass

        try:
            print(f'image keys: {self.images._images.keys()}')
            for i in range(10,125):
                for j in self.letters:
                    self.xcell=j + str(i)
                    if self.xcell in self.images._images.keys():
                        self.image=self.images.get(self.xcell)
                        self.image_name_raw=str(self.image_name) + "_" + str(self.fp_no) + "_(" + self.xcell + ").png"
                        self.image_name_refined=str(general_functions.convert_to_file_folder_name(self.image_name_raw))
                        self.image.save(self.result_path / self.images_folder / self.image_name_refined)

                        # Update result dictionary with refined image name
                        kwargs['df_for_excel'][24][len(kwargs['df_for_excel'][24])-1]=self.image_name_refined
                        # print('I saved image: ', self.image_name_refined)
                        # os.system('pause')
        except:
            pass
        finally:
            try:
                self.images._images.clear()
            except:
                pass

    def xls_image_extracter(self, **kwargs):
        # what can you do with this function
        # 1. Extract image from excel file
        # 2. Save it in images folder
        # 3. Update result dictionary with refined image name
        # 4. Clear clipboard
        # 5. Return None
        # Refactor this function

        
        self.xls_image_name=kwargs['image_name']
        self.xls_fp_no=kwargs['fp_no']
        self.xls_excel_name=kwargs['excel_name']

        # self.excel_app=win32com.client.Dispatch('Excel.Application')
        self.xls_workbook=kwargs['workbook']
        self.xls_sheet=kwargs['sheetname']

        for n, shape in enumerate(self.xls_sheet.Shapes):
            print(f"n:{n}, shape:{shape}")
            # print(shape.TopLeftCell)
            print(f"shape name: {shape.Name}")
            print(f"shape width: {shape.Width}")
            print(f"shape height: {shape.Height}")
            if shape.Width<1 or shape.Height<1:
                continue
            # if any of shape dimension is not bigger than 30 then skip this shape
            if shape.Width<30 and shape.Height<30:
                continue
            # If shape name contains rectangle then skip this shape as well, check it with lower case
            if 'rectangle' in shape.Name.lower():
                continue
            # if shape is not picture or image then skip this shape
            if shape.Type!=13:
                continue
            
            try:
                print(shape.TopLeftCell.row)
            except Exception as e:
                print(f"Exception after: print(shape.TopLeftCell.row)")
                print(str(e))
                continue
            try:
                if shape.TopLeftCell.row>3:
                    print(f"Shape/'s row is bigger that 3")
                    shape.Copy() # Copies from Excel to Windows clipboard
                    print('Copied...')
                    try:
                        self.xls_image = ImageGrab.grabclipboard()
                    except Exception as e:
                        print(f"Exception in self.xls_image = ImageGrab.grabclipboard()")
                        print(str(e))
                        continue
                    print(f"self.xls_image:{self.xls_image}")
                    self.xls_image_name_raw=str(self.xls_image_name) + "_" + str(self.xls_fp_no) + "_(" + str(n) + ").png"
                    print(f"self.xls_image_name_raw:{self.xls_image_name_raw}")
                    self.xls_image_name_refined=str(general_functions.convert_to_file_folder_name(self.xls_image_name_raw))
                    print(f"self.xls_image_name_refined:{self.xls_image_name_refined}")
                    # Save image, it is now pillow image
                    self.xls_image.save(self.result_path / self.images_folder / self.xls_image_name_refined,'png')
                    print(f'Saved image successfully,{self.xls_image_name_refined}.png')
                # Update result dictionary with refined image name
                kwargs['df_for_excel'][24][len(kwargs['df_for_excel'][24])-1]=self.xls_image_name_refined
                print('Updated result dictionary with refined image name')
                if windll.user32.OpenClipboard(None):
                    windll.user32.EmptyClipboard()
                    windll.user32.CloseClipboard()
                print('Clipboard cleared-2.')
            except Exception as e:
                print('I could not copy this image: ', self.xls_image_name, ' - from excel file: ', self.xls_excel_name)
                print(str(e))

if __name__=="__main__":
    test=FP_Image_Extractor(xpath="D:\\bckup Aman\\coding\\fiche_produit\\cra_list_test", prd_name_row=2, prd_name_col=7)
    test.xlsx_image_extracter()
