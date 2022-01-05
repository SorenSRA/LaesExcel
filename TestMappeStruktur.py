import os


path_root = 'C:\\SRA'

for root, dirs, files in os.walk(path_root, topdown=True):
    if len(dirs) == 0:
        print(f'Dirs {root} XXXX {dirs} er et Dir')
