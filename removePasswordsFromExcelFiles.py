# imports
import os, win32com.client, pythoncom 

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# objects needed: USER TO MODIFY
source = r'C:\Users\mahmudur\Desktop\TempData\All OB\4002'                                              # source folder
destination = r'C:\Users\mahmudur\Desktop\TempData\All OB\4002\Operation Bulletine\password_free'       # target directory
password = 'fkl'
                                                                                        # password to open excel files
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# Routine for iterating through the folder strucures in the source folder to create a list of paths to excel files. 
# In the event that there are two files with the same name the larger of the two will be kept and the other discarded.

os.chdir(source)                                                              # set the working directory
all_files = []                                                                # container for all files
file_list = []                                                                # container for 'unique' file names
full_path_dict = {}                                                           # container for full path to file
for root, dirs, files in os.walk(source):                                     # walk through the folder structure
    for file_ in files:
        if file_.startswith('~$'):                                            # ignore temporary files
            continue
        if file_.lower().endswith((".xls", ".xlsx")):                         # find excel files
            file_info = os.stat(root + '\\' + file_)                          # get info about the file
            file_size = file_info.st_size                                     # get size of file
            all_files.append(file_)                                           # append file name to all_files
            if file_ not in file_list:                                        # if file not in file_list then update list and dict
                 file_list.append(file_)
                 full_path_dict[file_] = ((root + '\\' + file_), file_info.st_size)
            else:
                if full_path_dict[file_][1] > file_size:                      # else if size is greater than existing entry update dict
                    full_path_dict[file_] = ((root + '\\' + file_), file_info.st_size)

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# Routine for iterating the paths object created above. For each file the routine opens it with the password,
# and then saves in the destination folder without the password. 

errors = []                                               # container for files that open/save with error
xcl = win32com.client.Dispatch('Excel.Application')       # Create the client that will read the excel file
xcl.visible = False                                       # This option means that excel will not actually appear to be open
xcl.DisplayAlerts = False                                 # This disables pop ups.

for path_tuple in full_path_dict.values():                # iterate through full_path_dict
    filename = path_tuple[0]                              # set filename
    try:
        wb = xcl.workbooks.open(filename, 0, False, None, password, password)                                  # open the workbook
        wb.SaveAs(destination + '\\' + filename.split('\\')[-1], None, Password = '', WriteResPassword = '')   # save it
        wb.Close()                                                                                             # close it 
    except pythoncom.com_error:
        errors.append(filename)
else:
    print 'All Done'
    print len(errors)
    xcl.Quit()
        



