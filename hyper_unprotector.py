# -*- coding: utf-8 -*-
"""
Created on Sat Mar 18 13:54:50 2023

@author: u235
"""
import subprocess
import os
import argparse
import re
import shutil


def get_djvulibre_path():

    if os.name == 'nt':
        import winreg
        import platform
        bits = platform.architecture()[0]
        if bits == '64bit':
            _djvulibre_key = r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\DjVuLibre+DjView'
        else:
            _djvulibre_key = r'Software\Microsoft\Windows\CurrentVersion\Uninstall\DjVuLibre+DjView'
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as registry:
            with winreg.OpenKey(registry, _djvulibre_key) as regkey:
                path, _ = winreg.QueryValueEx(regkey, 'UninstallString')
        path = os.path.dirname(path)
        if os.path.isfile(os.path.join(path, 'djvused.exe')):
            djvused = '"'+path+'\\djvused.exe "'
    else:
        djvused = r'djvused '
    # djvused=r'"c:\Program Files (x86)\DjVuLibre\djvused.exe "' #manual set path
    return djvused


def backup_djvu(book_name):
    name, ext = os.path.splitext(book_name)
    new_book_name = ''.join([name, '_new', ext])
    shutil.copy(book_name, new_book_name)
    return new_book_name


def get_links(book_name, djvused):
    proc = subprocess.Popen(' '.join(
        [djvused, book_name, ' -u -e "output-ant" ']), stdout=subprocess.PIPE, shell=True)
    output = proc.stdout.read()
    txt = output.decode('utf-8')
    txt = re.sub(r' \(maparea', r'\n(maparea', txt)  # split by spaces
    return txt


def get_page_names(book_name, djvused):

    proc = subprocess.Popen(' '.join(
        [djvused, book_name, ' -u -e "ls" ']), stdout=subprocess.PIPE, shell=True)
    output = proc.stdout.read()
    txt = output.decode('utf-8')
    txt = txt.split('\n')
    result = []
    for row in txt:
        lst = row.split()
        if len(lst) > 2 and lst[1] == 'P':
            result.append([lst[0], lst[3]])
    return result


def save_links(book_name, links, djvused):
    with open('myfile.dsed', encoding="utf-8", mode='w') as f:
        f.write(links)
    subprocess.Popen(' '.join(
        [djvused, book_name, ' -f myfile.dsed -s ']), stdout=subprocess.PIPE, shell=True)


def edit_ant(txt, LUT):
    # replace page name by page number
    for s in LUT:
        txt = re.sub(r'\(maparea "#' +
                     str(s[1])+r'"', r'(maparea "#'+str(s[0])+r'"', txt)
        txt = txt.replace('\r', "")
    return(txt)


def delete_old_links(txt):
    new_txt = []
    txt = txt.split('\n')
    pattern1 = r'\(maparea "#[0-9]+"'
    pattern2 = r'\(maparea "#"'
    for row in txt:
        if re.search(pattern2, row):
            if re.search(pattern1, row):
                new_txt.append(row)
            else:
                print(f'delete row: {row}')
        else:
            new_txt.append(row)

    new_txt = '\n'.join(new_txt)
    return new_txt
def print_LUT(LUT):
    print('page num\tpage name')
    print('-'*12)
    for row in LUT:
        print(f'{row[0]}\t{row[1]}')

def main():
    parser = argparse.ArgumentParser(
        prog='Hyperlink djvu unprotector',
        description='This script may be use for unprotect hyperlinks in djvu file, after add/delete/shuffle pages',
        epilog='--18.03.2023--')
    parser.add_argument('-f', '--file')
    args = parser.parse_args()
    book_name = args.file
    djvused = get_djvulibre_path()
    new_book_name = backup_djvu(book_name)
    txt = get_links(new_book_name, djvused)
    LUT = get_page_names(book_name, djvused)
    #print_LUT(LUT)
    new_links = edit_ant(txt, LUT)
    new_links=delete_old_links(new_links)
    save_links(new_book_name, new_links, djvused)


if __name__ == "__main__":
    main()
