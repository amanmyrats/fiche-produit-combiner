import win32com.client       # Need pywin32 from pip
from PIL import ImageGrab    # Need PIL as well
import os
import datetime

from pathlib import Path
import general_functions

start_time=datetime.datetime.now()

def image_extracter(xpath):
    dest_folder=Path(xpath)
    for excel_file in dest_folder.rglob("*.xls*"):
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            workbook = excel.Workbooks.Open(excel_file)
        except:
            continue

        wb_folder = workbook.Path
        wb_name = workbook.Name
        wb_path = os.path.join(wb_folder, wb_name)

        print ("Extracting images from ", wb_path)

        image_no = 0

        for sheet in workbook.Worksheets:
            for n, shape in enumerate(sheet.Shapes):
                #if shape.Name.startswith("Picture"):
                # Some debug output for console
                image_no += 1
                print ("---- Image No. ", image_no)

                # Sequence number the pictures, if there's more than one
                # num = "" if n == 0 else "_%03i" % n
                print("Top Left Cell: ", shape.TopLeftCell.Column)
                print("This is n: ", n)
                if n>1:
                    if shape.TopLeftCell.Column>0:
                        if str(sheet.Cells(2,7))=="None":
                            continue
                        
                        filename = str(general_functions.convert_to_file_folder_name(str(sheet.Cells(2,7)))) + " NO-" + str(n)  + ".png"
                        file_path = os.path.join (Path(wb_folder) / "result_image", filename)

                        print ("Saving as ", file_path)    # Debug output

                        shape.Copy() # Copies from Excel to Windows clipboard

                        # Use PIL (python imaging library) to save from Windows clipboard
                        # to a file
                        image = ImageGrab.grabclipboard()
                        image.save(file_path,'png')

        workbook.Close(False)


image_extracter("D:\\bckup Aman\\coding\\fiche_produit\\cra_list_test")

end_time=datetime.datetime.now()

print("Total time: ", end_time-start_time)