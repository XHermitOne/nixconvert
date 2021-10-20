/**
* Модуль сервисных функций
* @file
*/

#if !defined( __TOOLS_H )
#define __TOOLS_H

#ifdef __cplusplus
#include <cstdio>			// FILE *
#else
#include <stdio.h>			// FILE *
#endif
//#include <stdio.h>			
#include <sys/stat.h>
#include <stdarg.h>
#include <string.h>
#include <ctype.h>
#include <stdlib.h>
#include <unistd.h>
#include <iconv.h>
#include <math.h>

#include <fcntl.h>              /* O_RDONLY */
#include <sys/types.h>
#include <sys/statvfs.h>
#include <sys/param.h>          /* MAXPATHLEN */
//#include <sys/wait.h>

#include "ictypes.h"
#include "log.h"

#define MAX_NUM_LEN 20

int checkParm(const char *check,int argc, char **argv);

void errBox(char *Fmt,...);

char sayError(char *msg);

char sayMessage(char *Msg, int aOptions);

char *getHomePath(void);       

BOOL dir_exists(char *path);
int mkpath(const char *path, mode_t mode);

/**
* Путь к папке конфига
*/
char *getCfgPath(void);

BOOL file_exists(char *FileName);
int copy_file(char *to, char *from);
BOOL del_file(char *FileName);

/**
*   Прочитать файл и вернуть массив символов
*/
unsigned int load_txt_file_size(char *FileName, char *result);
char *load_txt_file(char *FileName);

/**
*   Выбрать максимальное/минимальное значение из 2-х
*/
double max(double a, double b);
double min(double a, double b);

char *itoa(int value, int base);
char * dtoa(double n);

int bytes_to_int(BYTE *bytes);
double bytes_to_double(BYTE *bytes);

long get_now_time(void);
long time_to_long(char *buffer);
char *long_to_time(long time);

/**
*   Нормализация пути
*/
char *norm_path(char *src, size_t src_len=-1, BOOL bFree=FALSE);

/**
*   Конвертация представления пути из dos(C:\\path\\) в unix(C:/path/)
*/
char *dos_to_unix_path(char *src, BOOL bFree=FALSE);

/**
*   Поменять расширение в имени файла
*/
char *change_filename_ext(char *filename, const char *new_ext, BOOL bFree=FALSE);

/**
*   Проверка на то же самый файл
*/
BOOL is_samefile(const char *filename1, const char *filename2);


/**
*   Цвета
*/
enum icPrintColor {
    IC_NORMAL_COLOR_TEXT = 0,      // normal
    IC_RED_COLOR_TEXT,             // red
    IC_GREEN_COLOR_TEXT,           // green
    IC_YELLOW_COLOR_TEXT,          // yellow
    IC_BLUE_COLOR_TEXT,            // blue
    IC_PURPLE_COLOR_TEXT,          // purple
    IC_CYAN_COLOR_TEXT,            // cyan
    IC_WHITE_COLOR_TEXT,           // white
};
/**
*   Печать цветного текста
*/
void print_color_txt(unsigned int color, char *str, ...);

/**
*   Определение количества разрядов десятичного числа
*/
unsigned int digit_count(int n);

#endif /*__TOOLS_H*/
