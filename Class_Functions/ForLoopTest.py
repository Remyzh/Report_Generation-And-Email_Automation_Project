
lines = []

with open("File_names.txt", 'r') as f:
    lines = [line.rstrip() for line in f] 

print(lines)
    
