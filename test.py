# !/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Тестовый модуль.
"""

import os
import os.path
import shutil
import time

import deb_build

LOG_FILENAME = '/home/xhermit/.nixconvert/nixconvert.log'

# Цвета в консоли
RED_COLOR_TEXT = '\x1b[31;1m'       # red
GREEN_COLOR_TEXT = '\x1b[32m'       # green
YELLOW_COLOR_TEXT = '\x1b[33m'      # yellow
BLUE_COLOR_TEXT = '\x1b[34m'        # blue
PURPLE_COLOR_TEXT = '\x1b[35m'      # purple
CYAN_COLOR_TEXT = '\x1b[36m'        # cyan
WHITE_COLOR_TEXT = '\x1b[37m'       # white
NORMAL_COLOR_TEXT = '\x1b[0m'       # normal

DEFAULT_ENCODING = 'utf-8'


def print_color_txt(sTxt, sColor=NORMAL_COLOR_TEXT):
    if type(sTxt) == unicode:
        sTxt = sTxt.encode(DEFAULT_ENCODING)
    txt = sColor+sTxt+NORMAL_COLOR_TEXT
    print(txt)


def do_tst(cmd):
    print_color_txt('Start command: <%s>' % cmd, GREEN_COLOR_TEXT)
    start_time = time.time()
    # cmd_list=cmd.split(' ')
    # os.spawnv(os.P_WAIT,cmd_list[0],cmd_list)
    os.system(cmd)
    print_color_txt('Stop command: <%s> Time = <%s>' % (cmd,
                                                        time.time()-start_time),
                    GREEN_COLOR_TEXT)

commands = ( # './nixconvert --debug',
             #'./nixconvert --debug --log --cfg=./tst/text01/econvert.cfg --rep=./tst/text01/main.rep --out=./tst/text01/main.xlsx; libreoffice ./tst/text01/main.xlsx',
             #'./nixconvert --debug --log --rep=./tst/mega/prtskl.rep --out=./tst/mega/prtskl.xlsx; libreoffice ./tst/mega/prtskl.xlsx',
             #'./nixconvert --debug --log --dbf=./tst/ttn/export.dbf; libreoffice ./tst/ttn/ttn.xlsx',
             #'rm ./tst/bpr735/bpr_735u.xlsx; ./nixconvert --debug --log --cfg=./tst/bpr735/u_oper00.cfg --rep=./tst/bpr735/bookp.txt --template=./tst/bpr735/bpr_735u.xls; libreoffice ./tst/bpr735/bpr_735u.xlsx',
             #'./nixconvert --debug --log --cfg=./tst/ossaldo/u_oper00.cfg --rep=./tst/ossaldo/ossaldo.rep --template=./tst/ossaldo/f720301_.xls; libreoffice ./tst/ossaldo/f720301_.xlsx',
             #'cp ./tst/iznos/iznos_original.xls ./tst/iznos/iznos.xls; rm ./tst/iznos/iznos.xlsx; ./nixconvert --debug --log --cfg=./tst/iznos/u_oper00.cfg --rep=./tst/iznos/iznos.rep --template=./tst/iznos/iznos.xls; libreoffice ./tst/iznos/iznos.xlsx',
             #'cp ./tst/oc1/oc1_original.xls ./tst/oc1/oc1.xls; rm /home/xhermit/.dosemu/drive_c/#bal98/wrk/oc1.xlsx; cp ./tst/oc1/oc1.xls /home/xhermit/.dosemu/drive_c/#bal98/wrk/oc1.xls; ./nixconvert --debug --log --dbf=./tst/oc1/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/#bal98/wrk/oc1.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/#bal98/wrk/oc1.xls; cp ./tst/0c1/oc1_original.xls /home/xhermit/.dosemu/drive_c/#bal98/wrk/oc1.xls; ./nixconvert --debug --log --dbf=./tst/0c1/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/#bal98/wrk/oc1.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/bal98/wrk/19.xlsx; cp ./tst/gruz/19_original.xls /home/xhermit/.dosemu/drive_c/bal98/wrk/19.xls; ./nixconvert --debug --log --dbf=./tst/gruz/export.dbf; libreoffice ./tst/gruz/19.xlsx',
             #'cp ./tst/rep1/main_original.xls ./tst/rep1/main.xls; ./nixconvert --debug --log --cfg=./tst/rep1/econvert.cfg --rep=./tst/rep1/main.rep --template=./tst/rep1/main.xls --sheet=лист1; libreoffice ./tst/rep1/main.xlsx',
             #'cp ./tst/inv/inv_det_orig.xls ./tst/inv/inv_det.xls; ./nixconvert --debug --log --cfg=./tst/inv/inv.cfg --rep=./tst/inv/osnrep.tab --template=./tst/inv/inv_det.xls --sheet=Лист1; libreoffice ./tst/inv/inv_det.xlsx',
             #'./nixconvert --debug --log --dbf=./tst/inv/export.dbf',
             #'rm ./tst/bpr1137/bpr1137.xlsx; ./nixconvert --debug --log --cfg=./tst/bpr1137/u_oper00.cfg --rep=./tst/bpr1137/main.rep --template=./tst/bpr1137/bpr1137.xls --sheet=Лист1; libreoffice ./tst/bpr1137/bpr1137.xlsx',
             #'./nixconvert --debug --log --cfg=./tst/bpr735/u_oper00.cfg --rep=./tst/bpr735/main.rep --template=./tst/bpr735/bpr_735.xls; libreoffice ./tst/bpr735/bpr_735.xlsx',
             #'rm ./tst/realkg/realkg.xlsx; ./nixconvert --debug --log --cfg=./tst/realkg/realkg.cfg --rep=./tst/realkg/main.rep --template=./tst/realkg/realkg.xls --sheet=Лист1; libreoffice ./tst/realkg/realkg.xlsx',
             #'rm ./tst/iznos/f720416_.xlsx; ./nixconvert --debug --log --cfg=./tst/iznos/u_oper00.cfg --rep=./tst/iznos/iznos.rep --template=./tst/iznos/f720416_.xls; libreoffice ./tst/iznos/f720416_.xlsx',
             #'./nixconvert --debug --log --dbf=./tst/specodejda/export.dbf; libreoffice ~/.dosemu/drive_c/bal98/wrk/mb8.xlsx',
             # Проверка выравнивания ИНН
             #'rm ./tst/bpo/bpo_735u.xlsx; ./nixconvert --debug --log --cfg=./tst/bpo/u_oper00.cfg --rep=./tst/bpo/bookp.txt --template=./tst/bpo/bpo_735u.xls; libreoffice ./tst/bpo/bpo_735u.xlsx',
             #'./nixconvert --debug --log --dbf=./tst/oc4/export.dbf',
             # Проверка на отображение копеек
             #'rm ./tst/asver/asver.xlsx; ./nixconvert --debug --log --cfg=./tst/asver/u_oper00.cfg --rep=./tst/asver/udfrep.rep --template=./tst/asver/asver.xls; libreoffice ./tst/asver/asver.xlsx',
             # Проверка блока подписей
             #'rm ./tst/bpo/bpo1137u.xlsx; ./nixconvert --debug --log --cfg=./tst/bpr1137/u_oper00.cfg --rep=./tst/bpo/bookp.txt --template=./tst/bpo/bpo1137u.xls --sheet=Лист1; libreoffice ./tst/bpo/bpo1137u.xlsx',
             #'rm ./tst/1.xlsx; ./nixconvert --debug --log --dbf=./tst/gruz/export.dbf --template=./tst/1.xls; libreoffice ./tst/1.xlsx',
             #'rm ./tst/bpo_735/bpo_735.xlsx; ./nixconvert --debug --log --cfg=./tst/bpo_735/u_oper00.cfg --rep=./tst/bpo_735/main.rep --template=./tst/bpo_735/bpo_735.xls; libreoffice ./tst/bpo_735/bpo_735.xlsx',
             #'rm ./tst/aobv/vob2.xlsx; ./nixconvert --cfg=./tst/aobv/vob.cfg --rep=./tst/aobv/aobv.rep --template=./tst/aobv/vob2.xls --sheet=Лист1 --debug; libreoffice ./tst/aobv/vob2.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/bagln/wrk/4fss_01.xls; cp ./tst/4fss/4fss_01.xls /home/xhermit/.dosemu/drive_c/bagln/wrk/4fss_01.xls; ./nixconvert --debug --log --dbf=./tst/4fss/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/bagln/wrk/4fss_01.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/bagln/wrk/6_ndf00.xls; cp ./tst/6ndf/6_ndf00.xls /home/xhermit/.dosemu/drive_c/bagln/wrk/6_ndf00.xls; ./nixconvert --debug --log --dbf=./tst/6ndf/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/bagln/wrk/6_ndf00.xlsx',
             #'rm ./tst/bpr981/bpr_735u.xlsx; ./nixconvert --debug --log --cfg=./tst/bpr981/u_oper00.cfg --rep=./tst/bpr981/bookp2.txt --template=./tst/bpr981/bpr_735u.xls; libreoffice ./tst/bpr981/bpr_735u.xlsx',
             #'rm ./tst/bpr981/bpr_981u.xlsx; ./nixconvert --debug --log --cfg=./tst/bpr981/u_oper00.cfg --rep=./tst/bpr981/bookp1.txt --template=./tst/bpr981/bpr_981u.xls; libreoffice ./tst/bpr981/bpr_981u.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/bagln/wrk/2ndfl_15.xls; cp ./tst/2ndfl/2ndfl_15_original.xls /home/xhermit/.dosemu/drive_c/bagln/wrk/2ndfl_15.xls; ./nixconvert --debug --log --dbf=./tst/2ndfl/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/bagln/wrk/2ndfl_15.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/baguln/wrk/inv1osn.xls; cp ./tst/inv10/inv1osn_original.xls /home/xhermit/.dosemu/drive_c/baguln/wrk/inv1osn.xls; ./nixconvert --debug --log --dbf=./tst/inv10/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/baguln/wrk/inv1osn.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/pv/bagz/wrk/4fss_01.xls; cp ./tst/fss/4fss_01.xls /home/xhermit/.dosemu/drive_c/pv/bagz/wrk/4fss_01.xls; ./nixconvert --debug --log --dbf=./tst/fss/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/pv/bagz/wrk/4fss_01.xlsx',
             #'rm /home/xhermit/.dosemu/drive_c/pv/gb/wrk/nds1901.xls; cp ./tst/nds/nds1901.xls /home/xhermit/.dosemu/drive_c/pv/gb/wrk/nds1901.xls; ./nixconvert --debug --log --dbf=./tst/nds/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/pv/gb/wrk/nds1901.xlsx',
             'rm /home/xhermit/.dosemu/drive_c/bal98/wrk/akc2001.xls; cp ./tst/akc/akc2001.xls /home/xhermit/.dosemu/drive_c/bal98/wrk/akc2001.xls; ./nixconvert --debug --log --dbf=./tst/akc/export.dbf; libreoffice /home/xhermit/.dosemu/drive_c/bal98/wrk/akc2001.xlsx',
           )


def test():
    print_color_txt('NIXCONVERT test start .. ok', GREEN_COLOR_TEXT)
    
    # deb_build.check_libraries()
    deb_build.compile_and_link()

    for i, command in enumerate(commands):
        if os.path.exists(LOG_FILENAME):
            os.remove(LOG_FILENAME)
        do_tst(command)
        if os.path.exists(LOG_FILENAME):
            # os.system('gedit %s' % LOG_FILENAME)
            pass
        else:
            print_color_txt('File <%s> not found' % LOG_FILENAME, YELLOW_COLOR_TEXT)

    print_color_txt('NIXCONVERT test stop .. ok', GREEN_COLOR_TEXT)

if __name__ == '__main__':
    test()
