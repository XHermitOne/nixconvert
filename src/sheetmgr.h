#if !defined( __SHEETMGR_H )
#define __SHEETMGR_H

#include "econvert.h"

#define XLS_WIDTH_COEF 420.0
#define XLS_WIDTH_CORRECT_COEF (1.0 / 1307.7)
#define XLS_HEIGHT_COEF 20.0
#define XLS_FONT_SIZE_COEF 20.0
#define XLS_ARIAL_SIZE_COEF 27.0

#define MAX_LEN_ADDRESS 7

#define MAX_WORKBOOK_COUNT 10
#define MAX_WORKSHEET_COUNT 256
#define MAX_COLUMN_COUNT 256
#define MAX_ROW_COUNT 65535

#define MAX_PAGE_BREAKS 256

/**
*   Открыть xls шаблон и подготовить его для заполнения
*/
lxw_workbook *open_workbook_xls(char *xls_filename);

/**
*   Закрыть xlsx файл
*/
BOOL close_workbook_xlsx(lxw_workbook *oBook);


class icFont //: public CObject  
{
public:

    char *FontName;
    unsigned int FontSize;
    COLORREF TextColor;
    
    BOOL isBold;
    BOOL isItalic;
    BOOL isUnderline;
    
	icFont() {};
	virtual ~icFont() { FontName = strfree(FontName); };

};

class icAlignment //: public CObject  
{
public:

    BOOL isVerticalAlignment;
    BOOL isHorizontalAlignment;
    
	icAlignment() {};
	virtual ~icAlignment() {};

};


class icNumberFormat //: public CObject  
{
public:

	icNumberFormat() {};
	virtual ~icNumberFormat() {};

};

class icInterior //: public CObject  
{
public:

    COLORREF BackgroundColor;

	icInterior() {};
	virtual ~icInterior() {};

};

class icCellValue //: public CObject  
{
public:

    unsigned int iRow;
    unsigned int iCol;
    
    char *sValue;
    
	icCellValue() { sValue = NULL; };
	virtual ~icCellValue() { sValue = strfree(sValue); };

};


lxw_format *get_cell_format(lxw_worksheet *oSheet, 
                            unsigned int i_row, unsigned int i_col);
lxw_format *clone_cell_format(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                              unsigned int i_row, unsigned int i_col);
/**
*   Клонировать формат ячейки без обрамления и интерьера
*   Т.е. только формат данных
*/
lxw_format *clone_cell_data_format(lxw_workbook *oBook, lxw_worksheet *oSheet, unsigned int i_row, unsigned int i_col);

BOOL set_cell_value(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    unsigned int row, unsigned int col, char *value,
                    lxw_format *default_format=NULL);

/**
*	Функция по имени файла вычисляет название книги.
*	c:\prt\example.xls -> example.xls
*/
char *getWorkBookName(char *xlsName);

/**
*   Очистить лист от всех значений
*/
void cleanSheet(lxw_worksheet *oSheet);

/**
*   Очистить область ячеек от всех значений
*/
void cleanRange(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

/**
*   Заполнить область ячеек массивом формул
*/
void setRangeValueArray(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                        char *sAddressBegin, char *sAddressEnd, icStringCollection *Values);
icStringCollection *getRangeValueArray(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

void setRangeColumnWidth(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, double iWidth);

/**
*	Открывает нужную книгу и лист
*/
BOOL openBookAndList(lxw_workbook *oBook, char *FNameValue, char *ListValue,
                     BOOL bAddList);

/**
*	Принудительн закрывает книгу
*/
BOOL ForceCloseWorkbook(lxw_workbook *oBook, char *xlsName);
char *getWorkBookName(char *xlsName);
void SaveAndQuit(lxw_workbook *oBook);
void SaveActiveBook(lxw_workbook *oBook);

BOOL DelUnchangedList(char *bookName, icStringCollection *m_ChangedListName,
					 lxw_workbook *oBook);


/**
*   Удалить листы из книги
*/
BOOL deleteSheets(lxw_workbook *oBook, icStringCollection *SheetNames);

/**
*   Удалить колонки заданной области ячеек
*/
BOOL deleteRangeColumns(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

/**
*   Переместить листы в укзанную книгу на указанный лист
*/
BOOL moveSheet(icStringCollection *SheetNames, char *xlsName, char *sSheetName);

/**
*   Копировать листы в укзанную книгу на указанный лист
*/
BOOL copySheet(icStringCollection *SheetNames, char *xlsName, char *sSheetName);

/**
*   Установить отображение/скрытие колонок
*/
BOOL setRangeColumnHidden(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, BOOL bEnable);

/**
*   Установить отображение/скрытие строк
*/
BOOL setRangeRowHidden(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, BOOL bEnable);
                      
/**
*   Установить шрифт области ячеек
*/
BOOL setRangeFont(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                  char *sAddressBegin, char *sAddressEnd,
                  char *sFontName, unsigned int iFontSize, COLORREF TextColor,
                  BOOL bBold, BOOL bItalic, BOOL bUnderline);
                    
icFont *getRangeFont(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

/**
*   Установить обрамление области ячеек
*/
BOOL setRangeBorder(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    char *sAddressBegin, char *sAddressEnd,
                    DWORD topStyle, DWORD leftStyle, DWORD bottomStyle, DWORD rightStyle);

BOOL setRangeBorderAround(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    char *sAddressBegin, char *sAddressEnd,
                    DWORD topStyle, DWORD leftStyle, DWORD bottomStyle, DWORD rightStyle);

/**
*   Установить выравнивание области ячеек
*/
BOOL setRangeAlignment(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd,
                       BOOL bVerticalAlignment, BOOL bHorizontalAlignment);

icAlignment *getRangeAlignment(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

/**
*   Установить числовой формат области ячеек
*/
BOOL setRangeNumberFormat(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

icNumberFormat *getRangeNumberFormat(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);

/**
*   Установка цвета текста и цвета фона для области ячеек
*/
BOOL setRangeInterior(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd,
                      COLORREF BackgroundColor);

icInterior *getRangeInterior(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd);


BOOL setCellValue(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                  char *sCellAddress, char *sValue);
char *getCellValue(lxw_worksheet *oSheet, char *sCellAddress);

/**
*   Функция возвращает адрес следующей ячейки со смещением по колонке
*/
char *getNextCellAddress(lxw_worksheet *oSheet, char *sCellAddress, int iColumnOffset=1);
/**
*   Функция возвращает адрес ячейки со смещением
*/
char *getOffsetCellAddress(lxw_worksheet *oSheet, char *sCellAddress, int iRowOffset, int iColumnOffset);

/**
*   Фукнция преабразует номер строки и номер колонки в строковый адрес ячейки
*   В случае ошибки возвращает NULL
*/
char *get_cell_address(unsigned int iRow, unsigned int iColumn);

BOOL setRangeValue(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                   char *sAddressBegin, char *sAddressEnd, char *sValue);

/**
*   Получить лист и сделать его активным
*/
lxw_worksheet *get_worksheet_by_name(lxw_workbook *oBook, char *sSheetName);
lxw_worksheet *get_worksheet_by_idx(lxw_workbook *oBook, unsigned int iIndex);

#endif /*__SHEETMGR_H*/
