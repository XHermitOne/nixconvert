/**
* Модуль функций записи в лог
* @file
*/

#if !defined( __LOG_H )
#define __LOG_H

#define MAX_LOG_MSG 1024

void logAddLine(char *S, ...);
void logErr(char *S, ...);
void logWarning(char *S, ...);
void log_color_line(unsigned int iColor, char *S, ...);

#endif
