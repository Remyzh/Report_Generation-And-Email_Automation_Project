import os

os.chdir("C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test")

a = open("Sheet_names.txt", "w")

for path, subdirs, files in os.walk(r'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers'):
    for filename in files:
        file_name, file_type = os.path.splitext(filename)
        
        ManagerName,file_info = file_name.split("'")
        a.write(str(ManagerName) + str(" Results"))
        a.write("\n")
