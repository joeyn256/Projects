import sys
import os
import subprocess
import re


#WARNING DO NOT RUN IF YOU HAVE SENSITIVE INFORMATION ON COPIES

#this deletes all copies of documents while keeping the original

directory = '/Users/joeyn256/Documents'

l_dir = subprocess.check_output(['ls', directory])
l_dir = l_dir.decode('utf-8')
l_dir = l_dir.split('\n')

repeated_documents = {}

for i in l_dir:
    repeats = i.split('.')
    if len(repeats) == 2:
        try: 
            repeated_documents[repeats[1]].append(repeats[0])
        except: 
            repeated_documents[repeats[1]] = [repeats[0]]

for j in repeated_documents:
    type_of_file = j # can choose to target specific files
    pattern = r'\(([1-9]|[1-9][0-9])\)$'
    for s in repeated_documents[type_of_file]:
        if re.search(pattern, s):
            try:
                os.remove(os.path.join(directory, s+'.'+type_of_file))
                print(f'{s}.{type_of_file} deleted successfully')
            except OSError as error:
                print(f'Error deleting {s}.{type_of_file}: {error}')
