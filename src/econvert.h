/**
* Модуль главных структур программы и основных запускающих функций
* @file
*/
#if !defined( __ECONVERT_H )
#define __ECONVERT_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <mcheck.h>
#include <locale.h>
#include <malloc.h>
#include <getopt.h>
#include <time.h>
#include <limits.h>
#include <math.h>

#include <xlsxwriter.h>
#include <libxls/xls.h>
//#include <libxls/endian.h>

#include "version.h"
#include "ictypes.h"
#include "log.h"
#include "tools.h"
#include "strfunc.h"
#include "memfunc.h"
#include "list.h"
#include "textparser.h"
#include "structcfg.h"
#include "glob.h"
#include "dbfile.h"
#include "sheetmgr.h"
#include "structscript.h"
#include "spcsimb.h"
#include "arrayformat.h"

#include "autoproject.h"


/**
* Режим отладки
*/
extern BOOL DBG_MODE;


#endif /* __ECONVERT_H */
