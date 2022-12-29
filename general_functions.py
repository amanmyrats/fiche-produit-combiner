def is_there_value(xstr):
    if str(xstr)=="nan":
        return False
    if len(str(xstr))>0:
        return True
    else:
        return False

def convert_to_file_folder_name(xstr):
    unwanted=("/","\\", "*", ":", "?", "\"", "<", ">", "|", "\n")
    for uwchar in unwanted:
        xstr=str(xstr).replace(uwchar, "-")
    return xstr

def is_sheet_fiche_produit(*args, **kwargs):
    fp_at_least=["FICHE PRODUIT - КАРТОЧКА ИЗДЕЛИЯ", "Nom du produit", "Name of Product"]
    if kwargs['type']=='dataframe':
        found_match=0
        for index in range(kwargs['df'].shape[1]):
            for index, value in kwargs['df'].iloc[:, index].items():
                for atleast in fp_at_least:
                    if str(value).__contains__(atleast):
                        found_match+=1
    elif kwargs['type']=='openpyxl':
        found_match=0
        for row in kwargs['sheet'].iter_rows():
            for cell in row:
                for atleast in fp_at_least:
                    if str(cell.value).__contains__(atleast):
                        found_match+=1

    if found_match<1:
        return False
    else:
        print("This one contains one of minimum requirements of fiche produit.")
        return True

def is_integer(n):
    try:
        float(n)
    except ValueError:
        return False
    else:
        return float(n).is_integer()