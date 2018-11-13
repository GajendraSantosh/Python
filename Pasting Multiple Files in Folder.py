'''
This script will copy a file and Pasting into Multiple Files in a Folder.

Note: -> Folder name and File name must be updated in 'Names.csv' file, 
		 The .csv file must be in root directory of this script,
		 Enter the valid path.

'''

#Import Libraries
import pandas as pd
import shutil
import os

# Collecting Data from .CSV
names = pd.read_csv('Names.csv')

CopyFilePath = input('Copying file path')
PastePath = input('Paste Location path')

for region in sorted(set(names['Folders'])):
    for emp in sorted(names['File Name'][(names['Folders']==region)]):
        if not os.path.exists(os.path.join(PastePath,region)):
            os.makedirs(os.path.join(PastePath,region))
        shutil.copy(CopyFilePath,os.path.join(PastePath,region,emp))
    print("Files Created in {}".format(region))
	

