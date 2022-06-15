# Python 3.9
# Program to obtain spreadsheet with data from RIS files
# Execute  : In the command line, write: 
           # python3 ris_reader.py <source> <path>
           # <source> : Directory with RIS files
           # <path>   : Results Logger with information
# Example  : python3 news2csv.py RIS_files affiliations
# Author   : Alejandra J. Perea Rojas

import os
import sys
import csv
import rispy
import pandas as pd
from rispy import TAG_KEY_MAPPING as mapping
from pprint import pprint

class bcolors:
    FAIL = '\033[91m'
    ENDC = '\033[0m'

# Print tags
pprint(mapping)

folder = sys.argv[1]
results = sys.argv[2]

# Clears results file and writes CSV header
with open(results+'.csv','w') as f:
    f.write('DOI,Title,Year,First Authors,Affiliation\n')

def reader(filepath, counter):

  # Start counter for skipped entries
  count = 0 
  
  with open(filepath, 'r') as bibliography_file:

    try:
      entries = rispy.load(bibliography_file)
    except UnicodeDecodeError:
      print("UnicodeDecodeError at:", filepath, "Skipped.")
      print('Might be a DS_store file, \
            <find . -name ".DS_Store" -delete> could help')
      sys.exit()

    for entry in entries:
      try:
        f = csv.writer(open(results+'.csv', "a"))
        addr = "https://doi.org/" + entry['doi']
        f.writerow([addr,entry['title'],\
                    entry['year'],\
                    entry['authors'],\
                    entry['custom3']])
        # f.writerow([entry['title']])# doesn't skip any
      except KeyError:
        # print(bcolors.FAIL + \
        #       "\nMissing key, skipped\n" + bcolors.ENDC)
        count = count + 1
    
  print("Entries skipped for missing key:", count)
  return count
    
# Total count of skipped entries
counter = 0

for filename in os.listdir(folder):
    filepath = os.path.join(folder, filename)
    counter = counter + reader(filepath, counter)

read_file = pd.read_csv(results+'.csv')
read_file.to_excel(results+'.xlsx', \
                   index = None, header=True)

print(counter)