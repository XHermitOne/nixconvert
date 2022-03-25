# !/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Полная сборка DEB пакета программы nixconvert.
"""
import os
import os.path

__author__ = 'xhermit'
__version__ = (0, 0, 3, 1)

PACKAGENAME = 'nixconvert-exp'
PACKAGE_VERSION = '6.1'
LINUX_VERSION = 'ubuntu20.04'
COPYRIGHT = 'Ayan Company <xhermit@ayan.ru>'
DESCRIPTION = 'The Linux EConvert fork'

#Цвета в консоли
RED_COLOR_TEXT      =   '\x1b[31;1m'    # red
GREEN_COLOR_TEXT    =   '\x1b[32m'      # green
YELLOW_COLOR_TEXT   =   '\x1b[33m'      # yellow
BLUE_COLOR_TEXT     =   '\x1b[34m'      # blue
PURPLE_COLOR_TEXT   =   '\x1b[35m'      # purple
CYAN_COLOR_TEXT     =   '\x1b[36m'      # cyan
WHITE_COLOR_TEXT    =   '\x1b[37m'      # white
NORMAL_COLOR_TEXT   =   '\x1b[0m'       # normal

DEFAULT_ENCODING = 'utf-8'

DEBIAN_CONTROL_FILENAME = './deb/DEBIAN/control'

LIB_XLS_READER_SRC = ('./libxls/.libs/libxlsreader.so.8.0.2',
                      './libxls/.libs/libxlsreader.so',
                      './libxls/.libs/libxlsreader.so.8')

LIB_XLSX_WRITER_SRC = ('./libxlsxwriter/lib/libxlsxwriter.so',
                       './libxlsxwriter/third_party/minizip/ioapi.so',
                       './libxlsxwriter/third_party/minizip/zip.so')

def get_uname(Option_='-a'):
    """
    Результат выполнения комманды uname.
    """
    try:
        cmd = 'uname %s' % Option_
        return os.popen3(cmd)[1].readline()
    except:
        print_color_txt(u'Ошибка Uname <%s>' % cmd, RED_COLOR_TEXT)
        raise
    return None


def is_64_linux():
    """
    Определить разрядность Linux.
    @return: True - 64 разрядная ОС Linux. False - нет.
    """
    uname_result = get_uname()
    return 'x86_64' in uname_result


LINUX_PLATFORM = 'amd64' if is_64_linux() else 'i386'


DEBIAN_CONTROL_BODY = '''Package: %s
Version: %s
Architecture: %s
Maintainer: %s
Depends: 
Section: contrib
Priority: optional
Description: %s 
''' % (PACKAGENAME, PACKAGE_VERSION, LINUX_PLATFORM, COPYRIGHT, DESCRIPTION)


def print_color_txt(sTxt, sColor=NORMAL_COLOR_TEXT):
    if type(sTxt) == unicode:
        sTxt = sTxt.encode(DEFAULT_ENCODING)
    txt = sColor + sTxt + NORMAL_COLOR_TEXT
    print(txt)


def sys_cmd(sCmd):
    """
    Выполнить комманду ОС.
    """
    print_color_txt('System command: <%s>' % sCmd, GREEN_COLOR_TEXT)
    os.system(sCmd)


def compile_and_link():
    """
    Компиляция и сборка.
    """
    check_libraries()
    
    sys_cmd('make clean')
    sys_cmd('make')
    
    if not os.path.exists('nixconvert'):
        print_color_txt(u'Ошибка компиляции', RED_COLOR_TEXT)


def copy_programm_to_deb():
    """
    Копировать программу в папку сборки.
    """
    if not os.path.exists('./deb/usr/share/nixconvert'):
        os.makedirs('./deb/usr/share/nixconvert')

    if not os.path.exists('./deb/usr/bin'):
        os.makedirs('./deb/usr/bin')

    if os.path.exists('default.cfg'):
        sys_cmd('cp ./default.cfg ./deb/usr/share/nixconvert')
        sys_cmd('chmod 755 ./deb/usr/share/nixconvert/default.cfg')
    else:
        print_color_txt(u'Не найден файл <default.cfg> для сборки. Используется старый', YELLOW_COLOR_TEXT)

    if os.path.exists('nixconvert'):
        sys_cmd('chmod 755 ./nixconvert')
        sys_cmd('cp ./nixconvert ./deb/usr/bin')
        sys_cmd('chmod 755 ./deb/usr/bin/nixconvert')
    else:
        print_color_txt(u'Не найден файл <nixconvert> для сборки. Используется старый', YELLOW_COLOR_TEXT)
        
    
def copy_libraries_to_deb():
    """
    Копировать библиотеки в папку сборки.
    """
    if not os.path.exists('./deb/usr/lib'):
        os.makedirs('./deb/usr/lib')

    for lib_filename in LIB_XLS_READER_SRC:
        if os.path.exists(lib_filename):
            sys_cmd('chmod 755 %s' % lib_filename)
            sys_cmd('cp %s ./deb/usr/lib' % lib_filename)
            sys_cmd('chmod 755 ./deb/usr/lib/%s' % lib_filename)
        else:
            print_color_txt(u'Не найден файл <%s> для сборки. Используется старый' % lib_filename, YELLOW_COLOR_TEXT)
    
    for lib_filename in LIB_XLSX_WRITER_SRC:
        if os.path.exists(lib_filename):
            sys_cmd('chmod 755 %s' % lib_filename)
            sys_cmd('cp %s ./deb/usr/lib' % lib_filename)
            sys_cmd('chmod 755 ./deb/usr/lib/%s' % lib_filename)
        else:
            print_color_txt(u'Не найден файл <%s> для сборки. Используется старый' % lib_filename, YELLOW_COLOR_TEXT)

    
def check_libraries():
    """
    Проверка наличия библиотек в папке /usr/lib.
    """
    sys_cmd('chmod 777 ./libxls/configure')
    sys_cmd('./libxls/configure')
    sys_cmd('make --directory=./libxls clean')
    sys_cmd('make --directory=./libxls all')
    
    for lib_filename in LIB_XLS_READER_SRC:
        base_lib_filename = os.path.basename(lib_filename)
        usr_lib_filename = os.path.join('usr', 'lib', base_lib_filename)
        if (not os.path.exists(usr_lib_filename)) and os.path.exists(lib_filename):
            sys_cmd('sudo cp %s /usr/lib' % lib_filename)
            print_color_txt(u'Установка библиотеки <%s> для сборки' % base_lib_filename, YELLOW_COLOR_TEXT)
            sys_cmd('sudo ldconfig')
	elif not os.path.exists(usr_lib_filename):
            print_color_txt(u'Библиотека <%s> не установлена в ОС' % base_lib_filename, YELLOW_COLOR_TEXT)

    sys_cmd('make --directory=./libxlsxwriter clean')
    sys_cmd('make --directory=./libxlsxwriter all')

    for lib_filename in LIB_XLSX_WRITER_SRC:
        base_lib_filename = os.path.basename(lib_filename)
        usr_lib_filename = os.path.join('usr', 'lib', base_lib_filename)
        if (not os.path.exists(usr_lib_filename)) and os.path.exists(lib_filename):
            sys_cmd('sudo cp %s /usr/lib' % lib_filename)
            print_color_txt(u'Установка библиотеки <%s> для сборки' % base_lib_filename, YELLOW_COLOR_TEXT)
            sys_cmd('sudo ldconfig')
	elif not os.path.exists(usr_lib_filename):
            print_color_txt(u'Библиотека <%s> не установлена в ОС' % base_lib_filename, YELLOW_COLOR_TEXT)

            
def build_deb():
    """
    Сборка пакета.
    """
    # Прописать файл control
    try:
        control_file = None 
        control_file = open(DEBIAN_CONTROL_FILENAME, 'w')
        control_file.write(DEBIAN_CONTROL_BODY)
        control_file.close()
        control_file = None        
        print_color_txt('Save file <%s>' % DEBIAN_CONTROL_FILENAME, GREEN_COLOR_TEXT)
    except:
        if control_file:
           control_file.close()
           control_file = None 
        print_color_txt('ERROR! Write control', RED_COLOR_TEXT)
        raise
        
    copy_programm_to_deb()
    copy_libraries_to_deb()
    
    sys_cmd('fakeroot dpkg-deb --build deb')

    if os.path.exists('./deb.deb'):
        deb_filename='%s-%s-%s.%s.deb' % (PACKAGENAME,
                                          PACKAGE_VERSION,
                                          LINUX_VERSION, LINUX_PLATFORM)
        sys_cmd('mv ./deb.deb ./%s' % deb_filename)
    else:
        print_color_txt('ERROR! DEB build error', RED_COLOR_TEXT)


def build():
    """
    Запуск полной сборки.
    """
    import time

    start_time = time.time()
    # print_color_txt(__doc__,CYAN_COLOR_TEXT)
    compile_and_link()
    build_deb()
    sys_cmd('ls *.deb')
    print_color_txt(__doc__, CYAN_COLOR_TEXT)
    print_color_txt('Time: <%d>' % (time.time()-start_time), BLUE_COLOR_TEXT)


if __name__ == '__main__':
    build()
