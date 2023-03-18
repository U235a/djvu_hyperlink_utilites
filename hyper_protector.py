# -*- coding: utf-8 -*-
"""
Created on Tue Mar 14 21:10:14 2023

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
    for s in LUT:
        if s[0] != s[1]:
            txt = re.sub(r'\(maparea "#' +
                         str(s[0])+r'"', r'(maparea "#_'+str(s[1])+r'"', txt)
    txt = re.sub(r'#_', r'#', txt)
    return(txt)


def main():
    parser = argparse.ArgumentParser(
        prog='Hyperlink djvu protector',
        description='This script may be use for protect hyperlinks in djvu file, after add/delete/shuffle pages',
        epilog='--14.03.2023--')
    parser.add_argument('-f', '--file')
    args = parser.parse_args()
    book_name = args.file
    djvused = get_djvulibre_path()
    new_book_name = backup_djvu(book_name)
    txt = get_links(new_book_name, djvused)
    LUT = get_page_names(book_name, djvused)
    new_links = edit_ant(txt, LUT)
    save_links(new_book_name, new_links, djvused)


if __name__ == "__main__":
    main()
