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
        print("Path associated")
        consolidated_path = read_txt(routes_text_path)
        print(consolidated_path)
        try:
            compare_dates(consolidated_path[0], data_dir)
            print(files, "\n")
        except:
            print("Files don't have corresponding data to use. Please check 'routes.txt' and 'source_names.txt' in 'data' folder")
        

        time.sleep(43200)

def create_file(file_name):
    with open(file_name, "w") as txt:
        print("Created", file_name)

def validation():
    script_dir = get_script_path()
    data_folder = "data"
    script_files = os.listdir()
    if data_folder not in script_files:
        os.mkdir("data")
    else:
        data_dir = os.path.join(script_dir, data_folder)
        data_files = os.listdir(data_dir)
        if len(data_files) < 2:
            try: 
                create_file(os.path.join(data_dir, "routes.txt"))
                create_file(os.path.join(data_dir, "source_names.txt"))
            except:
                print("Files couldn't be created")
        else:
            size = []
            for file in data_files:
                file_size = os.stat(os.path.join(data_dir, file)).st_size
                if file_size <= 0:
                    print(file, "is empty")
                    size.append(file_size)
            if 0 in size:
                return True
            else:
                return False
    return True
            
def main():
    val = True
    while val:
        val = validation()
        print(val)
        time.sleep(10)
    
    refresh()

main()
