// Описание структур заголовка и полей файла DBF
#if !defined( __DBSTRUCT_H )
#define __DBSTRUCT_H

#include "ictypes.h"

//////////////////////////////////////////////////////////////////////
// Характеристики DBF-а
typedef struct 
{
 BYTE  Type;            // Тип файла
 BYTE  Year, Month, Day;
 DWORD nRecords;        // Число записей
 WORD  Offset;          // Сдвиг на начало данных первой записи относительно начала файла
 WORD  RecordLen;       // Длина записи
 BYTE  Reserved1[16];
 BYTE  IndexFile;
 BYTE  Reserved2[3];
} DB_HEADER;

//////////////////////////////////////////////////////////////////////
// Характеристики поля
typedef struct
{
 BYTE  Name[11];        // Название поля
 BYTE  Type;            // Тип поля, из: 'C', 'N', 'D', 'L'
 DWORD Offset;          // Смещение на данные поля относительно начала данных записи
 BYTE  FieldLen;        // Длина поля
 BYTE  Decimals;        // Число знаков после десятичной точки
 BYTE  Reserved[14];
} DB_FIELD;

#endif /* __DBSTRUCT_H */
