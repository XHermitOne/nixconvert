/**
* Модуль определения дополнительных типов
* @file
*/

#if !defined( __ICTYPES_H )
#define __ICTYPES_H

#include <stdint.h>


typedef int                 BOOL;

#if (SIZEOF_LONG_INT == 4)
// Для 32-разрядных систем
typedef uint8_t     BYTE;
typedef uint16_t    WORD;
typedef uint32_t    DWORD;
#else
// Для 64-разрядных систем
typedef uint8_t     BYTE;
typedef uint16_t    WORD;
typedef unsigned int	DWORD;
#endif

#ifndef FALSE
#define FALSE     0
#endif

#ifndef TRUE
#define TRUE      1
#endif

#define MAX_PATH          260

typedef DWORD COLORREF;


#endif /*__ICTYPES_H*/

