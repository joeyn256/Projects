import sys
import os
import subprocess
import shutil
import re

directory = '/Users/joeyn256/Downloads'

l_dir = subprocess.check_output(['ls', directory])
l_dir = l_dir.decode('utf-8')
l_dir = l_dir.split('\n')


#below deletes directories that have repeats (1) and (2), etc.
"""
pattern = r'\(([1-9]|[1-9][0-9])\)$'

for s in l_dir:
    if re.search(pattern, s):
        try:
            shutil.rmtree(os.path.join(directory, s))
            print(f'{s} deleted successfully')
        except OSError as error:
            print(f'Error deleting {s}: {error}')
"""

#this will delete directories the start with 'hw'
"""
for i in l_dir:
    if i[:2] == 'hw':
        try:
            shutil.rmtree(os.path.join(directory, i))
            print(f'{i} deleted successfully')
        except OSError as error:
            print(f'Error deleting {i}: {error}')
"""
