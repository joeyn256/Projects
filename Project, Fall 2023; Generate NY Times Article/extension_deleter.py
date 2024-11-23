import sys
import os
import subprocess
import shutil
import re

directory = '/Users/joeyn256/Downloads'

l_dir = os.listdir(directory)
extension = input('enter the extension you want to delete (example: jpg): ')
print(f'input entered: {extension}')
extension = '.' + extension
print(f'Looking for inputs {extension} :] ')

extensions_found = 0
specific_extensions = []

for file_name in l_dir:
    file_path = os.path.join(directory, file_name)
    if os.path.isfile(file_path) and os.path.splitext(file_path)[1] == extension:
        extensions_found += 1
        specific_extensions.append(file_name)

print(f'{extensions_found} files found with extension {extension}')

user_input1 = input(f'Do you want to see the file names (Y/N): ')
bool2 = True

# block if you want to see the file names you are about to delete
while(bool2):
    if user_input1 == 'Y':
        print(f'file names = {specific_extensions}')

        # block to remove an item you want to save from the list
        user_input2 = input(f'Do you want to delete any file names from that list (Y/N): ')
        bool3 = True

        while(bool3):
            if user_input2 == 'Y':
                # block to remove the file name from the list and prompt if you want to save another file
                user_input3 = input(f'What file name do you want to delete: ')
                bool4 = True

                while(bool4):
                    if user_input3 in specific_extensions:
                        try:
                            specific_extensions.remove(user_input3)
                            print(f'{user_input3} deleted successfully')
                            print(f'new list = {specific_extensions}')
                            
                            #block to delete another file
                            user_input4 = input('Do you want to delete another file(Y/N): ')
                            if user_input4 == 'N': 
                                bool4 = False
                                bool3 = False
                            else:
                                user_input3 == input(f'What is the next file name that you want to delete: ')

                        except OSError as error:
                            print(f'Error deleting {user_input3}: {error}')
                            user_input3 = input(f'What file name do you want to delete: ')
                    else:
                        print(f'file name not found')
                        user_input4 = input('Try again (Y/N): ')
                        if user_input4 == 'Y':
                            user_input3 = input(f'What file name do you want to delete: ')
                        else:
                            bool4 = False
                            bool3 = False              
            elif user_input2 == 'N':
                bool3 = False
            else:
                print('Input not recognized')
                user_input2 = input(f'Do you want to delete any file names from that list (Y/N): ')

        bool2 = False
    elif user_input1 == 'N':
        bool2 = False
    else:
        print('Input not recognized')
        user_input1 = input(f'Do you want to see the file names (Y/N): ')

num_of_extensions = len(specific_extensions)
user_input = input(f'Do you want to delete all {num_of_extensions} extensions_found (Y/N): ')

bool1 = True

while(bool1):
    if user_input == 'Y':
        print(f'Files you are about to delete: {specific_extensions}')
        goodbye_files = input('Are you sure (Y/N): ')
        if goodbye_files == 'Y':
            for file_name in specific_extensions:
                try:
                    os.remove(os.path.join(directory, file_name))
                    print(f'{file_name} deleted successfully')
                except OSError as error:
                    print(f'Error deleting {file_name}: {error}')
            bool1 = False
        else:
            print('Files not deleted')
            bool1 = False
    elif user_input == 'N':
        print('Files not deleted')
        bool1 = False
    else:
        print('Input not recognized')
        user_input = input(f'Do you want to delete the {extensions_found} repeats (Y/N): ')
