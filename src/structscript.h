// StructScript.h: interface for the CStructScript class.
//
//////////////////////////////////////////////////////////////////////

#if !defined( __STRUCTSCRIPT_H )
#define __STRUCTSCRIPT_H

#include "econvert.h"


// Идентификаторы команд удаления листов
#define DEL_NOT_CHANGED_LIST 1
#define DEL_CHANGED_LIST 2

class icStructCFG;

//	Выполняет команды скрипта пост-обработки
BOOL DoPostConvertScript(char *fileScript, char * bookName,
						 char *strCfgName, int idCommand,
						 icStringCollection *m_ChangedListName,
						 lxw_workbook *oBook);

BOOL RunScript(char *fileScript, char *bookName,
               char *strCfgName, int idCommand,
               icStringCollection *m_ChangedListName,
               lxw_workbook *oBook, BOOL bPost);


class icStructScript //: public CObject  
{
public:
	BOOL ParseScriptFile();
	//	Набор команд, выполняемых перед конвертацией
	icStringCollection m_PostCommands = icStringCollection();

	//	Набор команд выполняемых после конвертации
	icStringCollection m_PreCommands = icStringCollection();

	void DoCommands(icStringCollection *mCommands);
	icStructScript();
	virtual ~icStructScript();

};

#endif /*__STRUCTSCRIPT_H*/
