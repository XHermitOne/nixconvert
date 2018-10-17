# Makefile for NixConvert project
#

# Basic stuff
SHELL = /bin/sh

top_srcdir = .
srcdir = .
prefix = /usr
exec_prefix = ${prefix}
bindir = $(exec_prefix)/bin
infodir = $(prefix)/info
libdir = $(prefix)/lib
mandir = $(prefix)/man/man1
includedir = $(prefix)/include

#CC = gcc
CC = g++
DEFS = -DHAVE_CONFIG_H
#CFLAGS = -g -O2 -Wall
CFLAGS =
#Флаг -g необходим для отладки!!!
#Флаг -w отключает warning!!!
#CPPFLAGS = -g
CPPFLAGS = -g -w -m64 
LDFLAGS = -v -L/usr/lib
LIBS = -lm 
BASELIBS = -lm 
#THREAD_LIBS = -lpthread
X11_INC = 
X11_LIB = 
#WX_CPPFLAGS=$(shell wx-config --cppflags)
#WX_LIBS=$(shell wx-config --libs)
#CAIRO_LIBS = $(shell pkg-config --cflags --libs cairo)
#GLIB_LIBS = $(shell pkg-config --libs glib-2.0)
QODS_LIBS = -lods -lquazip -lzlib
CXX_FLAGS = -fPIC -std=c++11 
QT_LIBS = -lQt5Core -lQt5Gui

LIBXLSXWRITER = ./libxlsxwriter/src/libxlsxwriter.a
LXW_LIBS = $(LIBXLSXWRITER) -lz
XLS_LIBS = -lxlsreader
#QT_LIBS = -DQT_CORE_LIB -DQT_GUI_LIB
# Directories
#QT_LIBDIR = $(libdir)/i386-linux-gnu/

TOPSRCDIR = .
TOPOBJDIR = .
SRCDIR    = .
#INCLUDEWXDIR = $(includedir)/wx-2.8
#CAIRO_INCLUDEDIR = $(includedir)/cairo
#GLIB_INCLUDEDIR = $(includedir)/glib-2.0
#GLIBCONFIG_INCLUDEDIR = /usr/lib/i386-linux-gnu/glib-2.0/include/
QODS_INCLUDEDIR = ./QOds/
QT_INCLUDEDIR = $(includedir)/qt5
QTCORE_INCLUDEDIR = $(QT_INCLUDEDIR)/QtCore
QTGUI_INCLUDEDIR = $(QT_INCLUDEDIR)/QtGui
LXW_INCLUDEDIR = ./libxlsxwriter/include/
XLS_INCLUDEDIR = ./libxls/include/

#MODULE    = none

CPPFLAGS += $(CXX_FLAGS)
#CPPFLAGS += $(WX_CPPFLAGS)
#CPPFLAGS += -I$(INCLUDEWXDIR)
#CPPFLAGS += -I$(CAIRO_INCLUDEDIR)
#CPPFLAGS += -I$(GLIB_INCLUDEDIR)
#CPPFLAGS += -I$(GLIBCONFIG_INCLUDEDIR)
#CPPFLAGS += -I$(QODS_INCLUDEDIR)
#CPPFLAGS += -I$(QT_INCLUDEDIR)
#CPPFLAGS += -I$(QTCORE_INCLUDEDIR)
#CPPFLAGS += -I$(QTGUI_INCLUDEDIR)
CPPFLAGS += -I$(LXW_INCLUDEDIR)
CPPFLAGS += -I$(XLS_INCLUDEDIR)
#LDFLAGS += $(CAIRO_LIBS)
#LDFLAGS += $(GLIB_LIBS)
#LDFLAGS += $(QODS_LIBS)
#LDFLAGS += $(QT_LIBS)
LDFLAGS += $(LXW_LIBS)
LDFLAGS += $(XLS_LIBS)
#LDFLAGS += -L$(QT_LIBDIR)

nixconvert: main.o version.o tools.o strfunc.o memfunc.o log.o structcfg.o list.o textparser.o spcsimb.o arrayformat.o sheetmgr.o structscript.o glob.o dbfile.o autoproject.o
	$(CC) -o nixconvert ./obj/main.o ./obj/version.o ./obj/tools.o ./obj/strfunc.o ./obj/memfunc.o ./obj/log.o ./obj/structcfg.o ./obj/list.o ./obj/textparser.o ./obj/spcsimb.o ./obj/arrayformat.o ./obj/sheetmgr.o ./obj/structscript.o ./obj/glob.o ./obj/dbfile.o ./obj/autoproject.o $(LDFLAGS)

#test: test.o version.o tools.o strfunc.o log.o structcfg.o list.o textparser.o spcsimb.o arrayformat.o sheetmgr.o structscript.o glob.o dbfile.o autoproject.o
#	$(CC) -o test ./obj/test.o ./obj/version.o ./obj/tools.o ./obj/strfunc.o ./obj/log.o ./obj/structcfg.o ./obj/list.o ./obj/textparser.o ./obj/spcsimb.o ./obj/arrayformat.o ./obj/sheetmgr.o ./obj/structscript.o ./obj/glob.o ./obj/dbfile.o ./obj/autoproject.o $(LDFLAGS)

test: test.o version.o tools.o strfunc.o log.o structcfg.o list.o textparser.o spcsimb.o arrayformat.o glob.o dbfile.o sheetmgr.o structscript.o autoproject.o
	$(CC) -o test ./obj/test.o ./obj/version.o ./obj/tools.o ./obj/strfunc.o ./obj/log.o ./obj/structcfg.o ./obj/list.o ./obj/textparser.o ./obj/spcsimb.o ./obj/arrayformat.o ./obj/glob.o ./obj/dbfile.o ./obj/sheetmgr.o ./obj/structscript.o ./obj/autoproject.o $(LDFLAGS)

tst_xlsx: tst_xlsx.o 
	$(CC) -o tst_xlsx ./obj/tst_xlsx.o $(LDFLAGS)

main.o: ./src/main.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/main.c
	mv main.o ./obj/main.o

test.o: ./src/test.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/test.c
	mv test.o ./obj/test.o

tst_xlsx.o: ./src/tst_xlsx.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/tst_xlsx.c
	mv tst_xlsx.o ./obj/tst_xlsx.o

version.o: ./src/version.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/version.c
	mv version.o ./obj/version.o

log.o: ./src/log.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/log.c
	mv log.o ./obj/log.o

tools.o: ./src/tools.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/tools.c
	mv tools.o ./obj/tools.o

strfunc.o: ./src/strfunc.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/strfunc.c
	mv strfunc.o ./obj/strfunc.o

memfunc.o: ./src/memfunc.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/memfunc.c
	mv memfunc.o ./obj/memfunc.o

list.o: ./src/list.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/list.c
	mv list.o ./obj/list.o

textparser.o: ./src/textparser.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/textparser.c
	mv textparser.o ./obj/textparser.o

spcsimb.o: ./src/spcsimb.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/spcsimb.c
	mv spcsimb.o ./obj/spcsimb.o

arrayformat.o: ./src/arrayformat.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/arrayformat.c
	mv arrayformat.o ./obj/arrayformat.o

glob.o: ./src/glob.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/glob.c
	mv glob.o ./obj/glob.o

dbfile.o: ./src/dbfile.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/dbfile.c
	mv dbfile.o ./obj/dbfile.o

structcfg.o: ./src/structcfg.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/structcfg.c
	mv structcfg.o ./obj/structcfg.o

sheetmgr.o: ./src/sheetmgr.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/sheetmgr.c
	mv sheetmgr.o ./obj/sheetmgr.o

structscript.o: ./src/structscript.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/structscript.c
	mv structscript.o ./obj/structscript.o
    
#econvert.o: ./src/econvert.c
#	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/econvert.c
#	mv econvert.o ./obj/econvert.o

autoproject.o: ./src/autoproject.c
	$(CC) -c  $(CFLAGS) $(CPPFLAGS) ./src/autoproject.c
	mv autoproject.o ./obj/autoproject.o


clean:
	rm -f ./src/*.o ./obj/*.o ./*.o nixconvert test

