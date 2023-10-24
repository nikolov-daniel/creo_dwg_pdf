# PDF DWG creoson
print ("Create DWG and PDF from drw in creo\n")
print ("Made by Daniel Nikolov")
#v1.0
# make folders Autocad, pdf
# making dwg and PDF files
# merging all pdf files

# v1.1
# making log file 0_dwg_pdf.log
# if don't exist pdf files in pdf_(format) don't make pdf_file_(format)
# added message to start creo if not started
# rename folders in exact name if folders with same name exist ("AutoCAD", PDF_A4, PDF_A3... not "autocad" or pdf_A3)

#V1.2
#bug fix merging more files

#v1.3
#automaticly start creoson with running the program from c:\users\USER\Desktop\creoson
#split export dwg or pdf or export all (dwg and pdf)
#must exist creoson folder on desktop!

#v1.4
#improved speed by not opening model from drawing - only read its parameters
#remove log files from autocad folder
#remove pdf folders if empty
#if parameter doesn't exist added 0 value for poz parameter
#made 'erase not displayed' on demand

#v1.5 added pause for merging pdfs
#add made by Daniel Nikolov
#added format a0 plus
#changed directory for saving pdfs in PDF_PRINT


import os
from datetime import datetime
from pathlib import Path
import creopyson
from functools import reduce  # Python 3 compatibility
from PyPDF2 import PdfFileMerger, PdfFileReader
import logging
import sys
import time
import getpass
############################## v1.2
sys.setrecursionlimit(150000)

def f_delete_logs():
    for file in os.listdir(pateka_autocad):
        if file[-6:-1] == ".log.":
            file_path = os.path.join(pateka_autocad, file)
            os.remove(file_path)
        elif file[-7:-2] == ".log.":
            file_path = os.path.join(pateka_autocad, file)
            os.remove(file_path)
        elif file[-8:-3] == ".log.":
            file_path = os.path.join(pateka_autocad, file)
            os.remove(file_path)
    print("removed log files from AutoCAD folder")

#upatstvo
print("""
With this program you are:

- Saving drawings as DWG files in AutoCAD folder in Working Directory and deleting log files
- Saving drawings as PDF files in PDF folders in Working Directory
- Merging all pdf files depending on format

---------!!! You must have opened CREO and CREOSON server on desktop !!!---------

Recommended:
- Have opened top assembly in session.
- Have ONLY drawings that you need in working directory or workspace!
- if parameter 'poz' doesn't exist in the part or assembly of the drawing, it's considered 0!

""")


yes = ("yes", "YES", "Yes", "Y", "y")
no = ("no", "NO", "No", "N", "n")
var_pdf = ("pdf", "p", "P", "PDF")
var_dwg = ("dwg", "d", "D", "DWG")

####################### v1.3
user = getpass.getuser()

#starting creoson
CREOSON_PATH = Path(f"C:\\Users\\{user}\\Desktop\\creoson")

os.chdir(CREOSON_PATH)
run_path = Path(CREOSON_PATH) / "creoson_run.bat"
os.startfile(run_path.resolve().as_posix())
print("Creoson started!")
c = creopyson.Client()
c.connect()

while c.is_creo_running()==False:
    print("""
--------------------
Creo is not running or more than one creo is running!
Please start Creo!
--------------------""")
    time.sleep(10)
else:
    print("Creo is running!")


print("You are using creo Windchill or Disk?")
creo_windchill = ("w", "W", "windchill", "Windchill", "WINDCHILL")
creo_disc = ("d", "D", "DISK", "Disk", "disk")
creo = input("Enter 'w' for windchill or 'd' for disk: ")

if creo in creo_disc:
    print("You chose creo disc!")
    list_of_drawings = c.creo_list_files("*.drw")
    print("Number od drawings: ", len(list_of_drawings))
elif creo in creo_windchill:
    print("You chose creo windchill!")
    if user == "user1":
        password = "password1"
    elif user == "user2":
        password = "password2"
    else:
        print("username:", user)
        password = input("Enter your password: ")
    c.windchill_authorize(user, password)

#   da napravam dokolku e pogresen password da se logira povtorno
#    while c.windchill_authorize(user, password) == False:
#        user = input("enter user: ")
#        password = input("Enter your password: ")

    workspace = c.windchill_get_workspace()

    list_of_drawings = c.windchill_list_workspace_files(workspace, "*.drw")
    print("Number od drawings: ", len(list_of_drawings))
else:
    creo = input("enter w or d: ")

pdf_dwg = input('Enter "p" to get only PDF files, "d" to get only DWG files, or ANY KEY to get PDF and DWG: ')

c.server_pwd()  #lokacija na server creoson
# working directory
current_directory = c.creo_pwd() #  return current working directory.
logging.basicConfig(filename=f"{current_directory}0_dwg_pdf.log", level=logging.INFO)

#print("list of drw files: ", list_of_drawings)  #  return a list in the working directory.
if pdf_dwg not in var_pdf:
    autocad = "AutoCAD"
    pateka = Path(f"{current_directory}{autocad}")  #  create and use Autocad folder in WD
    if not Path.is_dir(pateka):
        Path.mkdir(pateka)
        print("AutoCAD folder was made!")
    else:
        print("AutoCAD folder exists!")
        pateka.rename(Path(current_directory, autocad))
    pateka_autocad = pateka

pateka_pdf_print = Path(f"{current_directory}PDF_PRINT")
pateka_pdf_a4 = Path(f"{pateka_pdf_print}/PDF_A4")
pateka_pdf_a3 = Path(f"{pateka_pdf_print}/PDF_A3")
pateka_pdf_a2 = Path(f"{pateka_pdf_print}/PDF_A2")
pateka_pdf_a1 = Path(f"{pateka_pdf_print}/PDF_A1")
pateka_pdf_a0 = Path(f"{pateka_pdf_print}/PDF_A0")
pateka_pdf_a0_plus = Path(f"{pateka_pdf_print}/PDF_A0_PLUS")
pateka_pdf_nestandarden_format = Path(f"{pateka_pdf_print}/PDF_nestandardni_formati")
if pdf_dwg not in var_dwg:

    if not Path.is_dir(pateka_pdf_print):
        Path.mkdir(pateka_pdf_print)
        print("PDF_PRINT folder was made!")
    else:
        pateka_pdf_print.rename(Path(current_directory, "PDF_PRINT"))
        print("PDF_PRINT folder exist!")

    if not Path.is_dir(pateka_pdf_a4):
        Path.mkdir(pateka_pdf_a4)
        print("PDF_A4 folder was made!")
    else:
        pateka_pdf_a4.rename(Path(pateka_pdf_print, "PDF_A4"))
        print("PDF_A4 folder exist!")

    if not Path.is_dir(pateka_pdf_a3):
        Path.mkdir(pateka_pdf_a3)
        print("PDF_A3 folder was made!")
    else:
        pateka_pdf_a3.rename(Path(pateka_pdf_print, "PDF_A3"))
        print("PDF_A3 folder exist!")

    if not Path.is_dir(pateka_pdf_a2):
        Path.mkdir(pateka_pdf_a2)
        print("PDF_A2 folder was made!")
    else:
        pateka_pdf_a2.rename(Path(pateka_pdf_print, "PDF_A2"))
        print("PDF_A2 folder exist!")

    if not Path.is_dir(pateka_pdf_a1):
        Path.mkdir(pateka_pdf_a1)
        print("PDF_A1 folder was made!")
    else:
        pateka_pdf_a1.rename(Path(pateka_pdf_print, "PDF_A1"))
        print("PDF_A1 folder exist!")

    if not Path.is_dir(pateka_pdf_a0):
        Path.mkdir(pateka_pdf_a0)
        print("PDF_A0 folder was made!")
    else:
        pateka_pdf_a0.rename(Path(pateka_pdf_print, "PDF_A0"))
        print("PDF_A0 folder exist!")

    if not Path.is_dir(pateka_pdf_a0_plus):
        Path.mkdir(pateka_pdf_a0_plus)
        print("PDF_A0_PLUS folder was made!")
    else:
        pateka_pdf_a0_plus.rename(Path(pateka_pdf_print, "PDF_A0_PLUS"))
        print("PDF_A0_PLUS folder exist!")

    if not Path.is_dir(pateka_pdf_nestandarden_format):
        Path.mkdir(pateka_pdf_nestandarden_format)
        print("PDF_nestandardni_formati folder was made!")
    else:
        pateka_pdf_nestandarden_format.rename(Path(pateka_pdf_print, "PDF_nestandardni_formati"))
        print("PDF_nestandardni_formati folder exist!")

# go otvara sekoj .drw file
for file in list_of_drawings:
    print("==========================================================")
#    c.file_exists(file)  #  verify if "file" exists.
    c.file_open(file, display=True)  #  Open "file" in Creo.

# loop za da gi dobieme site sheets od crtez
    i = 1
    model_file = c.drawing_get_cur_model()
    #c.file_open(model_file, display = True)
    parametar_pozcija_raw = c.parameter_list(file_=model_file, name="poz")
    #c.file_close_window(model_file)
    c.file_open(file, display=True)  #  Open "file" in Creo.
    drawing_sheets = c.drawing_get_num_sheets(file)  #  get number of sheets in drawing
    print("opening file: ", file)
    print("have number of sheets: ", drawing_sheets)
    check_parameter_exist = c.parameter_exists(file_=model_file, name="poz")
    if check_parameter_exist == True:
        pozicija = reduce(lambda a, b: dict(a, **b), parametar_pozcija_raw)
        pozicija = (pozicija["value"])
    else:
        print("Parameter poz doesn't exist in model:", model_file, "poz = 0")
        pozicija = 0


#parametar pozicija dodava nuli napred za polesno sortiranje
    if pozicija < 10:
        pozicija = f"00{pozicija}"
    elif pozicija >= 10 and pozicija < 100:
        pozicija = f"0{pozicija}"

    while i <= drawing_sheets:
        print("_______________________________________________________")
        c.drawing_select_sheet(i)
        sheet_format = c.drawing_get_sheet_format(sheet=i)
        new_list = list(sheet_format.values())
        file_name = list(file.rsplit(".", maxsplit=1))
        file_name = file_name[0]
        if pdf_dwg not in var_pdf:
            print("saving as DWG file!")
            if drawing_sheets == 1:
                file_name_autocad = f"{pozicija}_{file_name}"
                file_name_pdf = f"{pozicija}_{file_name}"
            elif drawing_sheets > 1:
                file_name_autocad = f"{pozicija}_{file_name}_{i}"
                file_name_pdf = f"{pozicija}_{file_name}_{i}"
            save_as_dwg = f"""
            ~ Close `main_dlg_cur` `appl_casc`;~ Command `ProCmdModelSaveAs` ;\
            ~ Open `file_saveas` `type_option`;~ Close `file_saveas` `type_option`;\
            ~ Select `file_saveas` `type_option` 1 `db_560`;\
            ~ Activate `file_saveas` `check_is_secondary` 0;\
            ~ Activate `file_saveas` `Current Dir`;\
            ~ Select `file_saveas` `ph_list.Filelist` 1 \
            `{pateka_autocad}`;\
            ~ Activate `file_saveas` `ph_list.Filelist` 1 \
            `{pateka_autocad}`;\
            ~ Select `file_saveas` `ph_list.Filelist` 1 `AutoCAD`;\
            ~ Select `file_saveas` `ph_list.Filelist` 1 `AutoCAD`;\
            ~ Activate `file_saveas` `ph_list.Filelist` 1 `AutoCAD`;\
            ~ Update `file_saveas` `Inputname` `{file_name_autocad}`;\
            ~ Select `file_saveas` `ph_list.Filelist` 1 `{file_name_autocad}`;\
            ~ Activate `file_saveas` `OK`;~ Activate `export_2d_dlg` `OK_Button`;\
            ~ Activate `UI Message Dialog` `ok`;
    
            """

            c.interface_mapkey(save_as_dwg)
            print("file: ", f"{file_name_autocad}.dwg", "saved at: ", autocad)

#Koj format se koristi vo crtezot
        if pdf_dwg not in var_dwg:

            if any("A4" in word for word in new_list):
                format_crtez = "A4"
                pateka_pdf = "PDF_A4"
            elif any("A3" in word for word in new_list):
                format_crtez = "A3"
                pateka_pdf = "PDF_A3"
            elif any("A2" in word for word in new_list):
                format_crtez = "A2"
                pateka_pdf = "PDF_A2"
            elif any("A1" in word for word in new_list):
                format_crtez = "A1"
                pateka_pdf = "PDF_A1"
            elif any("A0_POD" in word for word in new_list):
                format_crtez = "A0 PLUS"
                pateka_pdf = "PDF_A0_PLUS"
            elif any("A0" in word for word in new_list):
                format_crtez = "A0"
                pateka_pdf = "PDF_A0"
            else:
                format_crtez = "NONSTANDARD"
                print("nonstandard format: ", file, "sheet: ", i)
                pateka_pdf = "PDF_nestandardni_formati"
            print("format of drawing: ", format_crtez)
            print("sheet number: ", i)
            print("saving as PDF file!")
            if drawing_sheets == 1:
                file_name_pdf = f"{pozicija}_{file_name}"
                save_as_pdf = f"mapkey 1p ~ Close `main_dlg_cur` `appl_casc`;~ Command `ProCmdModelSaveAs` ;\
                ~ Open `file_saveas` `type_option`;~ Close `file_saveas` `type_option`;\
                ~ Select `file_saveas` `type_option` 1 `db_617`;\
                ~ Activate `file_saveas` `check_is_secondary` 0;\
                ~ Activate `file_saveas` `Current Dir`;\
                ~ Select `file_saveas` `ph_list.Filelist` 1 `PDF_PRINT`;\
                ~ Activate `file_saveas` `ph_list.Filelist` 1 `PDF_PRINT`;\
                ~ Select `file_saveas` `ph_list.Filelist` 1 `{pateka_pdf}`;\
                ~ Activate `file_saveas` `ph_list.Filelist` 1 `{pateka_pdf}`;\
                ~ Update `file_saveas` `Inputname` `{file_name_pdf}`;\
                ~ Activate `file_saveas` `OK`;~ Activate `UI Message Dialog` `ok`;\
                ~ Activate `intf_profile` `pdf_export.pdf_launch_viewer` 0;\
                ~ Activate `intf_profile` `OkPshBtn`;"
            elif drawing_sheets > 1:
                file_name_pdf = f"{pozicija}_{file_name}_{i}"
                save_as_pdf = f"~ Close `main_dlg_cur` `appl_casc`;~ Command `ProCmdModelSaveAs` ;\
                ~ Open `file_saveas` `type_option`;~ Close `file_saveas` `type_option`;\
                ~ Select `file_saveas` `type_option` 1 `db_617`;\
                ~ Activate `file_saveas` `check_is_secondary` 0;\
                ~ Activate `file_saveas` `Current Dir`;\
                ~ Select `file_saveas` `ph_list.Filelist` 1 `{pateka_pdf}`;\
                ~ Activate `file_saveas` `ph_list.Filelist` 1 `{pateka_pdf}`;\
                ~ Update `file_saveas` `Inputname` `{file_name_pdf}`;\
                ~ Activate `file_saveas` `OK`;~ Activate `UI Message Dialog` `ok`;\
                ~ Select `intf_profile` `pdf_export.pdf_sheets_choice` 1 `current`;\
                ~ Activate `intf_profile` `pdf_export.pdf_launch_viewer` 0;\
                ~ Activate `intf_profile` `OkPshBtn`;"

            c.interface_mapkey(save_as_pdf)
            print("file: ", f"{file_name_pdf}.pdf", "saved at", pateka_pdf)
            now = datetime.now()
            if pdf_dwg in var_pdf:
                logging.info(f'file: {file_name}, saved as PDF at {pateka_pdf}; time: {now}')
            elif pdf_dwg in var_dwg:
                logging.info(f'file: {file_name}, sheet number: {i} saved as DRW at: {autocad}: time: {now}')
            else:
                logging.info(f'file: {file_name}, sheet number: {i} saved as DRW at: {autocad}; saved as PDF at {pateka_pdf}; time: {now}')

        i=i+1
    c.file_close_window(file)
    print("closing file: ", file)
print("")

#print(list_of_drawings)
print("_____________________________________________")
merge_pdf_ = input("prepare PDF folders and press any key to merge PDFs: ")
if pdf_dwg not in var_dwg:
    # a4
    is_empty_a4 = not any(Path(f"{pateka_pdf_a4}").iterdir())
    if is_empty_a4 == False:
        print("merging A4 PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_a4}"):
            item_path = os.path.join(pateka_pdf_a4,item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/A4.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_a4)
        print("folder PDF_A4 is empty and have been removed")
    # a3
    is_empty_a3 = not any(Path(f"{pateka_pdf_a3}").iterdir())
    if is_empty_a3 == False:
        print("merging A3 PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_a3}"):
            item_path = os.path.join(pateka_pdf_a3,item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/A3.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_a3)
        print("folder PDF_A3 is empty and have been removed")
    # a2
    is_empty_a2 = not any(Path(f"{pateka_pdf_a2}").iterdir())
    if is_empty_a2 == False:
        print("merging A2 PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_a2}"):
            item_path = os.path.join(pateka_pdf_a2, item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/A2.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_a2)
        print("folder PDF_A2 is empty and have been removed")
    # a1
    is_empty_a1 = not any(Path(f"{pateka_pdf_a1}").iterdir())
    if is_empty_a1 == False:
        print("merging A1 PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_a1}"):
            item_path = os.path.join(pateka_pdf_a1, item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/A1.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_a1)
        print("folder PDF_A1 is empty and have been removed")
    # a0
    is_empty_a0 = not any(Path(f"{pateka_pdf_a0}").iterdir())
    if is_empty_a0 == False:
        print("merging A0 PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_a0}"):
            item_path = os.path.join(pateka_pdf_a0, item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/A0.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_a0)
        print("folder PDF_A0 is empty and have been removed")

    is_empty_a0_plus = not any(Path(f"{pateka_pdf_a0_plus}").iterdir())
    if is_empty_a0_plus == False:
        print("merging A0 plus PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_a0_plus}"):
            item_path = os.path.join(pateka_pdf_a0_plus, item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/A0_PLUS.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_a0_plus)
        print("folder PDF_A0_PLUS is empty and have been removed")

    pateka_pdf_nestandarden_format_is_empty = not any(Path(f"{pateka_pdf_nestandarden_format}").iterdir())
    if pateka_pdf_nestandarden_format_is_empty == False:
        print("merging nonstandard PDFs")
        merger = PdfFileMerger()
        for item in os.listdir(f"{pateka_pdf_nestandarden_format}"):
            item_path = os.path.join(pateka_pdf_nestandarden_format, item)
            if item.endswith(".pdf"):
                merger.append(item_path)
        merger.write(f"{pateka_pdf_print}/NONSTANDARD.pdf")
        merger.close()
    else:
        os.rmdir(pateka_pdf_nestandarden_format)
        print("folder PDF_nestandardni_formati is empty and have been removed")
print("DONE!")

#delete log files from autocad folder
if pdf_dwg not in var_pdf:
    f_delete_logs()


print("-------------Do you want to 'erase not displayed memory' --------------")
q = input("Enter 'y' to erase not displayed: ")
if q in yes:
    c.file_erase_not_displayed()
    print("Memory was erased")
else:
    print("Memory wasn't erased")

q_list = ("q", "Q", "")
q = input("press q to exit: ")
while q not in q_list:
    q = input("press q to exit: ")