// ArrayFormat.h: interface for the CArrayFormat class.
//
//////////////////////////////////////////////////////////////////////

#if !defined( __ARRAYFORMAT_H )
#define __ARRAYFORMAT_H

#ifdef __cplusplus
#include <cstdio>			// FILE *
#else
#include <stdio.h>			// FILE *
#endif
// #include <stdio.h>
#include <stdlib.h>
#include "ictypes.h"

class icArrayFormat //: public CObject  
{
public:

	//	Признак скрытия строки
	BOOL bRowHide;

	//	Признак скрытия колонки
	BOOL bColHide;

	//	Признак скрытия колонки
	BOOL bColShow;

	//	Признак, пустого содержания 
	BOOL bEmptyContent;

	BOOL bBold, bItalic, bUnderline;
	BOOL _bBold, _bItalic, _bUnderline;

	/////////////////////
	//	Название шрифта
	char *nameFont;

	//	размер шрифта
	int sizeFont;

	//	Дополнительная информация

	COLORREF clrText;   //Цвет текста
	COLORREF clrBgr;    //Цвет фона
	char *cellLT;
    char *cellRB;         // координаты региона
    
	icArrayFormat();
	virtual ~icArrayFormat();

    void Print(void);
};

#endif /*__ARRAYFORMAT_H*/
