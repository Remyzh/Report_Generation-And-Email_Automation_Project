import os

os.chdir("C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test")

a = open("File_names.txt", "w")

for path, subdirs, files in os.walk(r'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers'):
    for filename in files:
        a.write(str(filename))
        a.write("\n")
