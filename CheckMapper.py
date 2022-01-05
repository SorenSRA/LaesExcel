import os

import forfitpartneroversigt
from forfitpartneroversigt import partnerliste, pathbase, pathspecff

print(pathbase)
print(pathspecff)

for key, value in partnerliste.items():
    print(key, value)


print(os.path.abspath(os.getcwd()))
print(os.path.basename(os.getcwd()))
print(os.path.dirname(os.getcwd()))
print(os.path.expanduser(os.getcwd()))
print(os.path.expandvars(os.getcwd()))
print(os.path.getsize(os.getcwd()))
print(os.path.isfile(os.getcwd()))
print(os.path.isdir(os.getcwd()))
print(os.path.normcase(os.getcwd()))
print(os.path.split(os.getcwd()))
print(os.path.splitdrive(os.getcwd()))
print(os.path.splitext(os.getcwd()))





