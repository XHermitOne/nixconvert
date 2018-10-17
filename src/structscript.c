// StructScript.cpp: implementation of the icStructScript class.
//
//////////////////////////////////////////////////////////////////////

#include "econvert.h"


/**
* Construction/Destruction
*/
icStructScript::icStructScript()
{
}

icStructScript::~icStructScript()
{
}

void icStructScript::DoCommands(icStringCollection *mCommands)
{
}

BOOL icStructScript::ParseScriptFile()
{
	return TRUE;
}


/**
*	Выполняет команды скрипта пост-обработки
*
*	idCommand: идентификатор команды для работы с измененными листами
*		0 - None: ничего не делать.
*		1 - DEL_NOT_CHANGED_LIST: удалить неизмененные листы
*		2 - DEL_CHANGED_LIST: удалить измененные листы
*/
BOOL DoPostConvertScript(char *fileScript, char * bookName,
						 char *strCfgName, int idCommand,
						 icStringCollection *m_ChangedListName,
						 lxw_workbook *oBook)
{
	BOOL ret = RunScript(fileScript, bookName, strCfgName, idCommand,
                         m_ChangedListName,	oBook, TRUE);
	return ret;
}

/**
*	Выполняет команды скрипта
*
*	fileScript: имя файла скрипта
*	bookName: имя текущей книги
*	strCfgName: имя файла, по которому можно вичислить полный путь до файла скрипта,
*				если он не указан
*	idCommand: идентификатор внешней команды для работы с измененными листами
*		0 - None: ничего не делать.
*		1 - DEL_NOT_CHANGED_LIST: удалить неизмененные листы
*		2 - DEL_CHANGED_LIST: удалить измененные листы
*  m_ChangeListName: список измененных листов
*	pExcel: указатель на 7приложение Excel
*  oBooks: текущий список книг приложения
*	bPost: признак постобработки
*/
BOOL RunScript(char *fileScript, char *bookName,
               char *strCfgName, int idCommand,
               icStringCollection *m_ChangedListName,
               lxw_workbook *oBook, BOOL bPost)
{
    // Проверка на не корректные аргументы
    if (fileScript == NULL)
        return FALSE;
        
	char *str = NULL;
    char *s1 = NULL;
    char *s2 = NULL;
    
	icStringCollection m_buff = icStringCollection();
    icStringCollection m_buffOem = icStringCollection();
    icStringCollection selSheets = icStringCollection();
    icStringCollection delSheets = icStringCollection();
    
    lxw_worksheet *oSheet = NULL;
    
	selSheets.RemoveAll();
	delSheets.RemoveAll();

	//	
	fileScript = MakeFileName(fileScript, strCfgName);

	if (!file_exists(fileScript))
	{
		if (idCommand == DEL_NOT_CHANGED_LIST && m_ChangedListName->GetLength() > 0)
		{
			BOOL bret = DelUnchangedList(bookName, m_ChangedListName, oBook);
			return bret;
		}
		else
			return FALSE;
	}

	if (!LoadTextFile(fileScript, str))
	{
		return FALSE;
	}
	else
	{
		if (strequal(str, ""))
			return TRUE;

		char *prefix = NULL;
		int lprefix = 0;

		if (!bPost)
		{
			prefix = "PRE ";
			lprefix = 4;
		}

		int nins, cursor = 0;
		BOOL bExit = FALSE;

		//	Буфер строки в WIN кодеровке
		char *buffS = NULL;

		//	Буфер строки в досовской кодеровке
		char *buffOem = NULL;

		//	Разбираем файл скрипта
		while (!bExit)
		{
			nins =strfind_offset(str, "\r\n", cursor);
			if (nins == -1)	            				 
			{
				bExit = TRUE;
				nins = strlen(str);
			}

			//	В DOS кодировке
			buffOem = substr(str, cursor, nins-cursor);
			buffS = strcopy(buffOem);

			//	В WIN кодировке
			buffS = cp866_to_utf8(buffS);

			// избавляемся от пустых строчек          
			cursor = nins+2;
			
			// Анализируем очередную строчку
			// - разбиваем на составляющие
			DividePattern(buffOem, &m_buffOem, "\"");
			DividePattern(buffS, &m_buff, "\"");
						
			if (m_buff.GetLength() > 0)
			{
				//		Первое слово переводим в верхний регистр и убираем 
				//	пробелы
				m_buff.SetIndexValue(0, strupr_lat(m_buff.GetIndexValue(0)));
				m_buff.SetIndexValue(0, strtrim(m_buff.GetIndexValue(0)));
				char *sbuff = m_buff.GetIndexValue(0);

				//	Атрибут указывает на книгу, которую надо открыть и какой
				//	лист выбрать (SELECT).
				//		Пример:
				//		PRE OPEN_WORKBOOK: "name.xls", "List1"
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix + 13), strprintf(NULL, "%sOPEN_WORKBOOK", prefix))  && m_buff.GetLength() > 1)
				{
					bookName = MakeFileName(m_buff.GetIndexValue(1), strCfgName);
					char *ListValue = "";

					if (m_buff.GetLength() > 3)
						ListValue = m_buff.GetIndexValue(3);

					openBookAndList(oBook, bookName, ListValue, FALSE);
					selSheets.RemoveAll();
				}
				
				//	Атрибут указывает куда скопировать или переместьть выбранные листы.
				//		Пример:
				//		PRE COPY_SELECTED_LIST: "main.xls", "Лист1"
				//		PRE MOVE_SELECTED_LIST: "main.xls", "Лист1"
				if ((strequal(strleft(m_buff.GetIndexValue(0), lprefix+18), strprintf(NULL, "%sCOPY_SELECTED_LIST", prefix)) ||
					 strequal(strleft(m_buff.GetIndexValue(0), lprefix+18), strprintf(NULL, "%sMOVE_SELECTED_LIST", prefix))) 
					 && m_buff.GetLength() > 1)
				{
					char *msg = NULL;

					bookName = MakeFileName(m_buff.GetIndexValue(1), strCfgName);
					char *ListName = "Лист1";

					if (m_buff.GetLength() > 3)
						ListName = m_buff.GetIndexValue(3);

					if (!openBookAndList(oBook, bookName, ListName, FALSE))
						msg = strprintf(msg, "Не удалось открыть нужную книгу <%s> и лист <%s>", bookName, ListName);

					if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+18), strprintf(NULL, "%sMOVE_SELECTED_LIST", prefix)))
                        moveSheet(&selSheets, m_buff.GetIndexValue(1), ListName);
					else
						copySheet(&selSheets, m_buff.GetIndexValue(1), ListName);
				}

				// <SELECT_LIST:> лист, который надо выделить 
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+11), strprintf(NULL, "%sSELECT_LIST", prefix)) && m_buff.GetLength() > 1)
				{
					char *ListValue = m_buff.GetIndexValue(1);
					
					if (openBookAndList(oBook, bookName, ListValue, FALSE))
						selSheets.AppendValue(ListValue);
				}
				// <DEL_COL:> колонку, которую надо удалить 
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+7), strprintf(NULL, "%sDEL_COL", prefix)) && m_buff.GetLength() > 1)
				{
					char *ColValue = m_buff.GetIndexValue(1);
					char *cell1 = strprintf(NULL, "%s1", ColValue);
					char *cell2 = strprintf(NULL, "%s65535", ColValue);

                    deleteRangeColumns(oSheet, cell1, cell2);
				}

				// Удаляет выделенные листы
				if (strequal(strleft(m_buff.GetIndexValue(0),lprefix+17), strprintf(NULL, "%sDEL_SELECTED_LIST", prefix)) && selSheets.GetLength() > 0)
				{
                    deleteSheets(oBook, &selSheets);
				}
			
				// Удаляет не выделенные листы
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+21), strprintf(NULL, "%sDEL_NOT_SELECTED_LIST", prefix)) && selSheets.GetLength() > 0)
				{
					DelUnchangedList(bookName, &selSheets, oBook);
				}
				
				// Закрывает книгу
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+14), strprintf(NULL, "%sCLOSE_WORKBOOK", prefix)) && m_buff.GetLength() > 1)
				{
					char *fileName = MakeFileName(m_buff.GetIndexValue(1), strCfgName);
					ForceCloseWorkbook(oBook, fileName);
				}
				
				// Удаляет файл
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+8), strprintf(NULL, "%sDEL_FILE", prefix)) && m_buff.GetLength() > 1)
				{
					char *fileName = MakeFileName(m_buff.GetIndexValue(1), strCfgName);
                    del_file(fileName);
				}

				// Копирует файл
				if (strequal(strleft(m_buff.GetIndexValue(0), lprefix+9), strprintf(NULL, "%sCOPY_FILE", prefix)) && m_buff.GetLength() > 3)
				{
					char *fileCopy = MakeFileName(m_buff.GetIndexValue(1), strCfgName);
					char *fileName = MakeFileName(m_buff.GetIndexValue(3), strCfgName);
					copy_file(fileCopy, fileName);
				}
			}
		}

		//	Удаляем файл скрипта
		if (bPost)
			del_file(fileScript);
            
	}

	return TRUE;
}

