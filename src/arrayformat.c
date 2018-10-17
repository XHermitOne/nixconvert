// ArrayFormat.cpp: implementation of the CArrayFormat class.
//
//////////////////////////////////////////////////////////////////////

#include "autoproject.h"
#include "arrayformat.h"
#include "strfunc.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

icArrayFormat::icArrayFormat()
{
	bRowHide = FALSE;
	bColHide = FALSE;
	bColShow = FALSE;
	bBold = bItalic = bUnderline = FALSE;
	_bBold = _bItalic = _bUnderline = FALSE;
	clrText = NULL;
    clrBgr = NULL;
	nameFont = NULL;
	sizeFont = 0;
	cellLT = NULL;
    cellRB = NULL;    
}

icArrayFormat::~icArrayFormat()
{

}

void icArrayFormat::Print(void)
{
    printf("[%s : %s]\tbRowHide=%d\tbColHide=%d\tbColShow=%d\tbBold=%d\tbItalic=%d\tbUnderline=%d\n", 
            cellLT, cellRB, 
            bRowHide, bColHide, bColShow, bBold, bItalic, bUnderline);
}

