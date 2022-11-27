import os
import time
import pythoncom
import win32com.client as win32

from datetime import datetime

#Function that returns the path where the script function is executing
def get_script_path():
    return os.getcwd()

#Function that reads a txt file and returns all lines as an array without line jumps (\n)
def read_txt(route):
    content = []
    with open(route, "r") as txt:
        for lines in txt:
            content.append(lines)
    return content

#Function that writes in a txt. If file doesn't exists, it creates it.
def write_txt(route, content):
    with open(route, "w") as txt:
        txt.write(content)

#Function that gets the modification date of a file
def get_file_mday(route):
    return datetime.fromtimestamp(os.path.getmtime(route))

#Function to open and refresh data that came from a query connection in excel
def open_close_as_excel(file_path):
    try:
        pythoncom.CoInitialize()
        Xlsx = win32.DispatchEx('Excel.Application')
        Xlsx.DisplayAlerts = False
        Xlsx.Visible = False
        book = Xlsx.Workbooks.Open(file_path)
        book.RefreshAll()
        Xlsx.CalculateUntilAsyncQueriesDone()
        book.Save()
        time.sleep(5)
        book.Close(SaveChanges=True)
        time.sleep(5)
        Xlsx.Quit()
        pythoncom.CoUninitialize()

        book = None
        Xlsx = None
        del book
        del Xlsx
        print("-- Opened/Closed as Excel --")

    except Exception as e:
        print(e)

    finally:
        # RELEASES RESOURCES
        book = None
        Xlsx = None

def compare_dates(consolidated_path, data_dir):
    aux = consolidated_path.split("\\")
    source_dir = ""
    shift_wbs = os.path.join(data_dir, "source_names.txt")
    consolidated_mday = get_file_mday(consolidated_path)
    for i in range(len(aux)-1):
        source_dir += aux[i] + "\\"

    files = read_txt(shift_wbs)
    
    for file in files:
        file_directory = os.path.join(source_dir, file.strip())
        #print(file_directory)
        file_mday = get_file_mday(file_directory)
        
        if file_mday > consolidated_mday:
            open_close_as_excel(consolidated_path) 
            break
        print(file_mday > consolidated_mday)


def refresh(): 
    script_dir = get_script_path()
    data_dir = os.path.join(script_dir, "data")
    #print(data_dir)
    route_file = "routes.txt"
    routes_text_path = os.path.join(data_dir, route_file)
    #print(routes_text)

    while True:
        files = os.listdir(data_dir)
        #print(files)
        if route_file not in files:
            print("No path associated, please insert route to consolidated file as show below")
            print("C:\\Users\\e0123456\\Consolidated.xlsx\n")
            path = input()
            write_txt(routes_text_path, path)
        else:
            print("Path associated")
            consolidated_path = read_txt(routes_text_path)
            compare_dates(consolidated_path[0], data_dir)
            print(files, "\n")
        

        time.sleep(10)
        
refresh()
            
