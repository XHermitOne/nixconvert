/**
* Модуль функций всё что связано с версией...
* @file
*/

#if !defined( __VERSION_H )
#define __VERSION_H

#ifdef __cplusplus
#include <cstdio>			// FILE *
#else
#include <stdio.h>			// FILE *
#endif
// #include <stdio.h>
#include <string.h>
#include <sys/sysinfo.h>

char *getVersion(void);
void printVersion(void);
void printHelp(void);
void printSysInfo(void);


#endif /*__VERSION_H*/
