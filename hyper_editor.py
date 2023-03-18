# -*- coding: utf-8 -*-
"""
Created on Sat Mar 11 07:58:03 2023

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
    #djvused=r'"c:\Program Files (x86)\DjVuLibre\djvused.exe "' #manual set path
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


def save_links(book_name, links, djvused):
    with open('myfile.dsed', encoding="utf-8", mode='w') as f:
        f.write(links)
    subprocess.Popen(' '.join(
        [djvused, book_name, ' -f myfile.dsed -s ']), stdout=subprocess.PIPE, shell=True)


def dict_links(txt):
    table = []
    current_page = None
    for row in txt.split('\n'):
        page = re.search(r'^select .* # page (\d+)', row)
        if page:
            current_page = page[1]
        link_page = re.search(r'^\(maparea "#(\d+).*', row)
        if link_page:
            table.append([int(current_page), int(link_page[1])])
    return(table)


def increment_links(table, start_page, shift):
    result=[]
    if shift>=0:
        result=[[i[1], i[1]] if i[1] < start_page else [i[1], i[1]+shift]for i in table]
    else:
        for i in table:
            if i[1]<start_page:
                result.append([i[1], i[1]])
            elif i[1]<start_page-shift:
                result.append([i[1], None])
            else:
                result.append([i[1], i[1]+shift])
    return result

def del_links(txt, LUT):
    for s in LUT:
         if s[1] == None:
             txt = re.sub(r'\(maparea "#' + str(s[0])+r'".*\n', r'', txt)
    return txt

def edit_ant(txt, LUT):
    for s in LUT:
        if s[0] != s[1]:
            txt = re.sub(r'\(maparea "#' +
                         str(s[0])+r'"', r'(maparea "#_'+str(s[1])+r'"', txt)
    txt = re.sub(r'#_', r'#', txt)
    return(txt)


def main():
    parser = argparse.ArgumentParser(
        prog='Hyperlink djvu editor',
        description='This script may be use for corrected hyperlinks in djvu file, after add/delete pages',
        epilog='--11.03.2023--')
    parser.add_argument('-f', '--file')
    parser.add_argument('-p', '--page', type=int, default=1)
    parser.add_argument('-s', '--shift', type=int, default=0)
    args = parser.parse_args()
    book_name,  start_page, shift = args.file, args.page, args.shift
    djvused = get_djvulibre_path()
    new_book_name = backup_djvu(book_name)
    txt = get_links(new_book_name, djvused)
    tbl = (dict_links(txt))
    LUT = increment_links(tbl, start_page, shift)
    new_links=del_links(txt, LUT)
    new_links = edit_ant(new_links, LUT)
    save_links(new_book_name, new_links, djvused)


if __name__ == "__main__":
    main()
