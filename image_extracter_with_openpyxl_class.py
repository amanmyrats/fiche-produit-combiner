import win32com.client       # Need pywin32 from pip
from PIL import ImageGrab    # Need PIL as well
import os
import datetime
import openpyxl
import string
import general_functions

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
        
        
        self.xls_image_name=kwargs['image_name']
        self.xls_fp_no=kwargs['fp_no']
        self.xls_excel_name=kwargs['excel_name']

        self.excel_app=win32com.client.Dispatch('Excel.Application')
        self.xls_workbook=kwargs['workbook']
        self.xls_sheet=kwargs['sheetname']

        for n, shape in enumerate(self.xls_sheet.Shapes):
            if shape.TopLeftCell.row>3:
                try:
                    shape.Copy() # Copies from Excel to Windows clipboard

                    self.xls_image = ImageGrab.grabclipboard()
                    self.xls_image_name_raw=str(self.xls_image_name) + "_" + str(self.xls_fp_no) + "_(" + str(n) + ").png"
                    self.xls_image_name_refined=str(general_functions.convert_to_file_folder_name(self.xls_image_name_raw))
                    # Save image, it is now pillow image
                    self.xls_image.save(self.result_path / self.images_folder / self.xls_image_name_refined,'png')
                except:
                    print('I could not copy this image: ', self.xls_image_name, ' - from excel file: ', self.xls_excel_name)

                # Update result dictionary with refined image name
                kwargs['df_for_excel'][24][len(kwargs['df_for_excel'][24])-1]=self.xls_image_name_refined

if __name__=="__main__":
    test=FP_Image_Extractor(xpath="D:\\bckup Aman\\coding\\fiche_produit\\cra_list_test", prd_name_row=2, prd_name_col=7)
    test.xlsx_image_extracter()