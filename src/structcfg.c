// StructCFG.cpp: implementation of the icStructCFG class.
//
//////////////////////////////////////////////////////////////////////

#include "econvert.h"


//////////////////////////////////////////////////////////////////////
// Форматированный тег
//////////////////////////////////////////////////////////////////////
icTagTextFormat::icTagTextFormat() 
{
    Name = NULL;
    clrText = -1;
    clrBgr = -1;
    sBeg = NULL;
    sEnd = NULL;
    bItalic = bBold = bUnderline = FALSE;
    nameFont = NULL;
    sizeFont = 0;
}

icTagTextFormat::~icTagTextFormat() 
{
    //if (DBG_MODE) logAddLine("Destructor ~icTagTextFormat()");
    Name = strfree(Name);
    sBeg = strfree(sBeg);
    sEnd = strfree(sEnd);
    nameFont = strfree(nameFont);            
}
    
BOOL icTagTextFormat::Init(char *tagname, COLORREF clr1, COLORREF clr2)
{
    Name = strinit(Name, strcopy(tagname));
    clrText = clr1;
    clrBgr = clr2;
    sBeg = strprintf(sBeg, "<%s>", Name);
    sEnd = strprintf(sEnd, "</%s>", Name);
    bItalic = bBold = bUnderline = FALSE;
    nameFont = strfree(nameFont);
    sizeFont = 0;
}
    
//////////////////////////////////////////////////////////////////////
// Менеджер конфигурационного файла
//////////////////////////////////////////////////////////////////////

icStructCFG::icStructCFG()
{
	iTab = 8;
	bHead = TRUE;
	bClear = TRUE;
	strTagHideRow = NULL; //strgen_empty();
	strTagHideCol = dupprintf( "<chide/>");
	strTagShowCol  = dupprintf( "<cshow/>");
	bCentredTextBeforeHead = TRUE;
	strListName = NULL; 
	strExcelName = NULL; 
	strEncoding = dupprintf( "DOS");
	strPostScript = NULL; 
	strCfgName = NULL; 
	bDelNotChangedList = FALSE;
    
    strSource = NULL;
	strTemplate = NULL;
	strTemplateList = NULL;
}

icStructCFG::~icStructCFG()
{
	strEncoding = strfree(strEncoding);
	strTagHideRow = strfree(strTagHideRow);
	strTagHideCol = strfree(strTagHideCol);
	strTagShowCol = strfree(strTagShowCol);
	strCfgName = strfree(strCfgName);
	strSource = strfree(strSource);
	strExcelName = strfree(strExcelName);
	strListName = strfree(strListName);
	strTemplate = strfree(strTemplate);
	strTemplateList = strfree(strTemplateList);
    strPostScript = strfree(strPostScript);
    
    m_ListName.RemoveAll();
    m_PageName.RemoveAll();
    m_DBF.RemoveAll();
    m_Div.RemoveAll();
    m_sPrizn.RemoveAll();
    strWord.RemoveAll();
    strReplace.RemoveAll();
    m_ChangedListName.RemoveAll();
    strPathWord.RemoveAll();
    strPathReplace.RemoveAll();
    
    for (int i=0; i < mTagFormat.GetCount(); i++)
    {
        if (mTagFormat.Items[i] != NULL)
        {
            //ВНИМАНИЕ!
            //Здесь необходимо явное приведение к типу класса 
            //для вызова деструтора
            delete ((icTagTextFormat *)mTagFormat.Items[i]);
        }
    }
    mTagFormat.RemoveAll();      
}

/**
*	Разбор файла конфигураций
*/
BOOL icStructCFG::AnalalizeCfgFile(char *CfgName)
{
	char *str = load_txt_file(CfgName);
    char *s1 = NULL;
    char *s2 = NULL;
    char *impStr;
    char *value = NULL;
	icStringCollection m_buffOem = icStringCollection();

	//   Читаем файл конфигураций 
    
	strWord.RemoveAll();
	strReplace.RemoveAll();
	m_sPrizn.RemoveAll();
	
	if (!str)
	{
		if (DBG_MODE) logErr("Нет возможности открыть файл конфигураций: <%s>", CfgName);
		return FALSE;
	}
	else
	{
		if (strempty(str))
        {
            if (DBG_MODE) logWarning("Empty config file <%s>", CfgName);
            str = strfree(str);
			return TRUE;
        }

        
        if (DBG_MODE) logAddLine("Tst1");
		strCfgName = strinit(strCfgName, strcopy(CfgName));
        if (DBG_MODE) logAddLine("Tst2");
        
		int nins;
        int cursor = 0;
		BOOL bExit = FALSE;

		//	Буфер строки в WIN кодеровке
		char *buffS = NULL;

		//	Буфер строки в досовской кодеровке
		char *buffOem = NULL;

		while (!bExit)
		{
			nins = strfind(str+cursor, "\r\n");
			if (nins == -1)	            				 
			{
				bExit = TRUE;
				nins = strlen(str); //.GetLength();
			}

			//	В DOS кодировке
            buffOem = strfree(buffOem);
            buffOem = substr(str, cursor, nins);
			buffOem = strtrim(buffOem, TRUE);
            
			//	В WIN кодировке
            buffS = strfree(buffS);
            buffS = cp866_to_utf8(buffOem, FALSE);
            if (DBG_MODE) logAddLine("Analize string <%s>\tCursor <%d>\tLen <%d>", buffS, cursor, nins-cursor);

			// избавляемся от пустых строчек          
			cursor = cursor + nins + 2;
			
			// Анализируем очередную строчку
			// - разбиваем на составляющие
			DividePattern(buffOem, &m_buffOem, "\"");
            //if (DBG_MODE) m_buffOem.PrintItemValues();
						
			if (m_buffOem.GetLength() > 0)
			{
				//	 Первое слово переводим в верхний регистр и убираем 
				//	пробелы
                value = strinit(value, m_buffOem.GetValue());
                value = strupr_lat(value);
                m_buffOem.SetValue(value);

				// Инструкцию на импортирования внешнего файла канфигурации
                value = strinit(value, m_buffOem.GetValue());
				if (startswith(value, "#INCLUDE") && m_buffOem.GetLength() > 1)
				{
					//
					char *impName = m_buffOem.GetNext()->GetValue();
					char *impPath = getFilePath(impName);
                    if (DBG_MODE) logAddLine("impName <%s>", impName);
                    if (DBG_MODE) logAddLine("impPath <%s>", impPath);

					if (impPath == "")
					{
						impPath = getFilePath(CfgName);
						impName = strprintf(impName, "%s/%s" , impPath, impName);
					}

					
					//	Если указан относительный путь, то абсолютный путь определяем
					//	по пути до файла конфигурации
					if (!LoadTextFile(impName, impStr))
					{
						logErr("Нет возможности открыть файл конфигураций: <%s>", impName);
                        str = strfree(str);
                        impName = strfree(impName);
                        impPath = strfree(impPath);
						return FALSE;
					}

					strconcatenate(str, impStr);
                    impName = strfree(impName);
                    impPath = strfree(impPath);
				}

				// Если нашли ключевое слово для расширений dbf  
				if (startswith(value, "DBF"))
				{
					int n = 1;
                    int bound = m_buffOem.GetLength()/2;
                    int indx;
					
					for(n; n<=bound; n++)
					{
						indx = n * 2 - 1;
						s1 = strinit(s1, m_buffOem.GetIndexValue(indx));
						//	Убираем пробелы
                        s1 = strtrim(s1, TRUE);

                        //  если расширение указано как *.ras
						if (strfind(s1, "*.") >= 0)
                        {
                            char *v1 = strright_pos(s1,1);
							m_DBF.AppendValue(v1);        
                            v1 = strfree(v1);
                        }
						else 
                        {
                            char *v1 = strconcatenate(".",s1);
                            //  указаны только расширения
							m_DBF.AppendValue(v1);
                            v1 = strfree(v1);
                        }
					}
				}

				// Если нашли ключевое слово для описания листов
				// LISTNAME:  
				if (startswith(value, "LISTNAME"))
				{
					int n = 1;
                    int bound = m_buffOem.GetLength() / 2;
                    int indx;
					
					for(n; n <= bound; n++)
					{
						indx = n * 2 - 1;
						s1 = strinit(s1, m_buffOem.GetIndexValue(indx));

						//	Убираем пробелы
                        s1 = strtrim(s1, TRUE);
                        //Имена листов используются в кодировке UTF8
                        s1 = cp866_to_utf8(s1, TRUE);
					}

				}

				// Если нашли ключевое слово для описания листов
				// PAGENAME:
				if (startswith(value, "PAGENAME"))
				{
					int n = 1;
                    int bound = m_buffOem.GetLength()/2;
                    int indx;
					
					for(n; n <= bound; n++)
					{
						indx = n*2-1;
						s1 = m_buffOem.GetIndexValue(indx);

						//	Убираем пробелы
                        s1 = strtrim(s1);
						m_PageName.AppendValue(s1);
					}
				}

				//   Если нашли ключевое слово для разделителей, формируем 
				// строчку разделителей
				
				if (startswith(value, "DIV"))
				{
					int n = 1;
                    int bound = m_buffOem.GetLength()/2;
                    int indx;
					
					for(n; n <= bound; n++)
					{
						indx = n * 2 - 1;
						s1 = strinit(s1, m_buffOem.GetIndexValue(indx));
						s2 = strinit(s2, strcopy(s1));
                        s1 = strlwr_lat(s1);
						
						if (startswith(s1, "space"))
						{
							// Определяем количество пробелов
							int nspace = atoi(strright(s1, 6));
							char *strSpace=strgen('0x20', nspace);
							m_Div.AppendValue(strSpace);
						}
						else
							m_Div.AppendValue(s2); 
					}
				}

				//		Если находим ключевое слово для табуляции, то
				//	изменяем размер табуляции. По умолчанию один символ табуляции
				//	соответствует 8 пробелам
				if (startswith(value, "TAB"))
				{
                    icStringCollection m_Word = icStringCollection();

					DividePattern(buffS, &m_Word, ":");
					s1 = strinit(s1, m_Word.GetIndexValue(1));
                    s1 = strlwr_lat(s1);
						
					int nspace = atoi(s1);
					iTab = nspace;
				}

				s1 = strinit(s1, m_buffOem.GetValue());
				s1 = strlwr_lat(s1);

				//	По ключевому слову  TEMPLATE определяется имя файла шаблона 
				//if (startswith(s1, "template"))
				//	strTemplate = strprintf(strTemplate, "CFG");

				//	По ключевому слову  NOHEAD, флаг шапки устанавливаем в FALSE
				if (startswith(s1, "nohead"))
					bHead = FALSE;
			
				//	По ключевому слову  NOCLEAR, флаг шапки устанавливаем в FALSE
				if (startswith(s1, "noclear"))
					bClear = FALSE;

				//	По ключевому слову LEFTTEXT определяется тэг, прижимающий текст
				//	перед шапкой к левому краю
				if (startswith(s1, "lefttext"))
					bCentredTextBeforeHead = FALSE;

				//	По ключевому слову DEL_NOT_CHANGED_LIST определяется тэг, прижимающий текст
				//	перед шапкой к левому краю
				if (startswith(s1, "del_not_changed_list"))
					bDelNotChangedList = TRUE;

				//	По ключевому слову REPLACE: формируем список замен
				if (m_buffOem.GetLength() >= 3 && startswith(s1, "replace"))
				{
					//	Вычленяем пары замен
					
					//	Если размер буфера < 4, означает ошибку в синтаксисе
					//	replace: "A1" "A2" -> {"replace:", "A1", " ", "A2"}
					if (m_buffOem.GetLength() >= 4)
					{
                        char *v1 = m_buffOem.GetIndexValue(1);
                        char *v2 = m_buffOem.GetIndexValue(3);
						strWord.AppendValue(v1);
						strReplace.AppendValue(v2);
                        v1 = strfree(v1);
                        v2 = strfree(v2);
					}
				}

                //	По ключевому слову PATH_REPLACE: формируем список замен
				if (m_buffOem.GetLength() >= 3 && startswith(s1, "path_replace"))
				{
					//	Вычленяем пары замен
                    					
					//	Если размер буфера < 4, означает ошибку в синтаксисе
					//	path_replace: "A1" "A2" -> {"path_replace:", "A1", " ", "A2"}
					if (m_buffOem.GetLength() >= 4)
					{
                        char *v1 = m_buffOem.GetIndexValue(1);
                        char *v2 = m_buffOem.GetIndexValue(3);
						strPathWord.AppendValue(v1);
						strPathReplace.AppendValue(v2);
                        v1 = strfree(v1);
                        v2 = strfree(v2);
					}
                    else
                        if (DBG_MODE) errBox("Ошибка синтаксиса секции <path_replace>");
				}
				
				//
				if (m_buffOem.GetLength() >= 1)
				{
				
					s2 = strinit(s2, m_buffOem.GetIndexValue(1));
                    s2 = (s2) ? strtrim(s2, TRUE) : strgen_empty();

					//	По ключевому слову  DYNHEADLINE определяются спец. слова 
					//	- признаки обновляемых строк в шапке отчета
					if (startswith(s1, "post_script"))
						strPostScript = strinit(strPostScript, s2);

					//	По ключевому слову  DYNHEADLINE определяются спец. слова 
					//	- признаки обновляемых строк в шапке отчета
					if (startswith(s1, "dynheadline"))
						s2 = strinit(s2, m_buffOem.GetIndexValue(1));
					m_sPrizn.AppendValue(s2);
			
					//		По ключевому слову SOURCE определяется имя отчета или
					//	буфферного файла
					if (startswith(s1, "source"))
						strSource = strinit(strSource, strcopy(s2));

					//		По ключевому слову ENCODING определяется имя кодировки отчета
					//	Возможные кодировки "WIN"(CP1251), "DOS"(IBM-866/866)
					if (startswith(s1, "encoding"))
						strEncoding = strinit(strEncoding, strcopy(s2));

					//		По ключевому слову EXCELNAME определяется имя ехсеl файла,
					//	куда конвертируется отчет
					if (startswith(s1, "excelname"))
						strExcelName = strinit(strExcelName, strcopy(s2));
				
					//	По ключевому слову  LISTNAME определяется имя листа 

					//	По ключевому слову  TEMPLATE определяется имя файла шаблона 
					if (startswith(s1, "template:"))
					{
						//	Если указан путь без имени шаблона, то в качестве
						//	имени берем имя отчета из <EXCELNAME>
						if (!find_word(s2, ".xls"))
						{
							char *nm = getFileName(strExcelName);

                            if (strempty(s2))
                                strTemplate = strcopy(strExcelName);
                            else
                            {
                                char *srev = strcopy(s2);
                                srev = strreverse(srev);
                            
                                if (strfind(srev, "/") == 0)
                                    strTemplate = strconcatenate(s2, nm);
                                else
                                    strTemplate = strprintf(strTemplate, "%s/%s", s2, nm);
                                srev = strfree(srev);
                            }
                            nm = strfree(nm);
						}
						else
							strTemplate = strinit(strTemplate, strcopy(s2));
                        if (DBG_MODE) logAddLine("Template: <%s>", strTemplate);
					}

					//	По ключевому слову  TEMPLATELIST определяется имя листа шаблона 
					if (startswith(s1, "templatelist"))
						strTemplateList = strinit(strTemplateList, strcopy(s2));

					//	По ключевому слову ROWHIDE определяется тэг, скрывающий строку
					if (startswith(s1, "rowhide"))
						strTagHideRow = strinit(strTagHideRow, strcopy(s2));

					//	По ключевому слову COLHIDE определяется тэг, скрывающий столбец
					if (startswith(s1, "colhide"))
						strTagHideCol = strprintf(strTagHideCol, "<%s>", s2);

					//	По ключевому слову COLSHOW определяется тэг, открывающий столбец
					if (startswith(s1, "colshow"))
						strTagShowCol = strprintf(strTagShowCol, "<%s>", s2);

					//	По ключевому слову  TEXTFORMAT определяются тэги, управляющие форматом текста  
					if (startswith(s1, "textformat"))
					{
						icTagTextFormat *pClr = _newTagTextFormat(s2, -1, -1, &m_buffOem);
						mTagFormat.AppendItem(pClr);
					}
				}
			}

            s1 = strfree(s1);
            s2 = strfree(s2);
            value = strfree(value);
            buffS = strfree(buffS);
            buffOem = strfree(buffOem);
		}
	}

    s1 = strfree(s1);
    s2 = strfree(s2);
    value = strfree(value);
    str = strfree(str);
	return TRUE;
}

icTagTextFormat *icStructCFG::_newTagTextFormat(char *sTagName, 
                                                COLORREF Color1, COLORREF Color2,
                                                icStringCollection *ParseString)
{
    char *value = NULL;
    icTagTextFormat *pClr = new icTagTextFormat(); 
    pClr->Init(sTagName, Color1, Color2);

    if (ParseString != NULL)
    {
        
        //////////////////////////////
        //	Читаем стили
        if (ParseString->GetLength() > 7)
        {
            value = strinit(value, ParseString->GetIndexValue(7));
            pClr->bItalic = (strequal(value, "t") || strequal(value, "T"));
        }

        if (ParseString->GetLength() > 9)
        {
            value = strinit(value, ParseString->GetIndexValue(9));
            pClr->bBold = (strequal(value, "t") || strequal(value, "T"));
        }

        if (ParseString->GetLength() > 11)
        {
            value = strinit(value, ParseString->GetIndexValue(11));
            pClr->bUnderline = (strequal(value, "t") || strequal(value, "T"));
        }

        //	Читаем название и размер шрифта
        if (ParseString->GetLength() > 13)
            pClr->nameFont = strinit(pClr->nameFont, ParseString->GetIndexValue(13));

        if (ParseString->GetLength() > 15)
            pClr->sizeFont = atoi(ParseString->GetIndexValue(15));
    }
        
    value = strfree(value);
    
    //if (DBG_MODE) logAddLine("TAG <%s> [%s] - [%s]", pClr->Name, pClr->sBeg, pClr->sEnd);
    
    return pClr;
}


icTagTextFormat *icStructCFG::ParseStyle(char *s1)
{
	COLORREF clr1 = -1;
    COLORREF clr2 = -1;
	icStringCollection m_buffOem = icStringCollection();
	char *tagname = NULL;
	//
	DividePattern(s1, &m_buffOem, "\"");
    
	if (m_buffOem.GetLength() > 1)
		tagname = m_buffOem.GetIndexValue(1);

	icTagTextFormat *pClr = _newTagTextFormat(tagname, clr1, clr2, &m_buffOem);


    tagname = strfree(tagname);
    
	return pClr;
}


void icStructCFG::PrintCFG(void)
{
    unsigned int i = 0;
    char *value = NULL;
    
    printf("Configuration file <%s> parameters:\r\n", strCfgName);
    printf("\tENCODING: %s\r\n", strEncoding);
	printf("\tEXCEL NAME: %s\r\n", strExcelName);
	printf("\tCFG NAME: %s\r\n", strCfgName);
	printf("\tSOURCE: %s\r\n", strSource);
	printf("\tLIST NAME: %s\r\n", strListName);

	printf("\tTEMPLATE: %s\r\n", strTemplate);
    printf("\tTEMPLATE LIST: %s\r\n", strTemplateList);

  	printf("\tTAG HIDE ROW: %s\r\n", strTagHideRow);
	printf("\tTAG HIDE COL: %s\r\n", strTagHideCol);
	printf("\tTAG SHOW COL: %s\r\n", strTagShowCol);
    
    printf("\tPOST SCRIPT: %s\r\n", strPostScript);

    printf("\tLIST NAME: \r\n");
    for (i=0; i < m_ListName.GetLength(); i++)
    {
        value = strinit(value, m_ListName.GetIndexValue(i));
        printf("\t\t<%s>\r\n", value);    
    }

    printf("\tPAGE NAME: \r\n");
    for (i=0; i < m_PageName.GetLength(); i++)
    {
        value = strinit(value, m_PageName.GetIndexValue(i));
        printf("\t\t<%s>\r\n", value);    
    }
        
    printf("\tDBF: \r\n");
    for (i=0; i < m_DBF.GetLength(); i++)
    {
        value = strinit(value, m_DBF.GetIndexValue(i));
        printf("\t\t<%s>\r\n", value);    
    }

    printf("\tDIV: \r\n");
    for (i=0; i < m_Div.GetLength(); i++)
    {
        value = strinit(value, m_Div.GetIndexValue(i));
        printf("\t\t<%s>\r\n", value);    
    }

    printf("\tPRIZN: \r\n");
    for (i=0; i < m_sPrizn.GetLength(); i++)
    {
        value = strinit(value, m_sPrizn.GetIndexValue(i));
        printf("\t\t<%s>\r\n", value);    
    }

    printf("\tCHANGED LIST NAME: \r\n");
    for (i=0; i < m_ChangedListName.GetLength(); i++)
    {
        value = strinit(value, m_ChangedListName.GetIndexValue(i));
        printf("\t\t<%s>\r\n", value);    
    }
        
    printf("\tREPLACE: \r\n");
    for (i=0; i < strWord.GetLength(); i++)
    {
        value = strinit(value, strWord.GetIndexValue(i));
        char *value2 = strReplace.GetIndexValue(i);
        printf("\t\t<%s : %s>\r\n", value, value2);
        value2 = strfree(value2);
    }

    printf("\tPATH REPLACE: \r\n");
    for (i=0; i < strPathWord.GetLength(); i++)
    {
        value = strinit(value, strPathWord.GetIndexValue(i));
        char *value2 = strPathReplace.GetIndexValue(i);
        printf("\t\t<%s : %s>\r\n", value, value2);
        value2 = strfree(value2);
    }
        
    value = strfree(value);
}
