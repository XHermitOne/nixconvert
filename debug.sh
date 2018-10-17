#!/bin/sh
echo Отладка
echo

make clean
#make test
make

gdb -x .gdbrc ./nixconvert
#gdb -x .gdbrc ./test
