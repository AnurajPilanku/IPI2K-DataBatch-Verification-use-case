'''
Author   :    Anuraj Pilanku
Usecase  :    IPI 2K DBV
'''
import shutil
import sys
import os
from datetime import date,datetime
today=date.today()
month=datetime.now().strftime('%B')
year=today.year

destination=r"\\PMCSNTTEST64\dropBox\Cognizant\IPI2K Daily Status\L1.5\_{year}\{month}_{year}".format(year=year,month=month)
path=r"\\acprd01\E\3M_CAC\IPI2K_DBV\mail_attachments"
files=os.listdir(path)
for file in files:
    source=os.path.join(path,file)
    if os.path.isdir(destination):
        shutil.copy(source,destination)
    else:
        os.mkdir(destination)
        shutil.copy(source, destination)
print("files placed successfully in :"+destination)

