
#include "econvert.h"

/**
* Функция создает архив
*/
BOOL CreateArchive(char *pFileName)
{
    int code;
    char bak_filename[MAX_PATH];
    sprintf(bak_filename, "%s.bak", pFileName);
    
    if (file_exists(bak_filename))
        del_file(bak_filename);
        
    code = copy_file(bak_filename, pFileName);
	if (code == 0)
    {
        if (DBG_MODE) logAddLine("Create archive file <%s>\tCode: %d", bak_filename, code);
		return TRUE;
    }
    else
        if (DBG_MODE) logAddLine("Don't create archive file <%s>\tCode: %d", bak_filename, code);

	return FALSE;
}


/**
* Загружаем текст, который будем конвертировать
*/
BOOL LoadTextFile(char *Name, char *ret)
{
	BOOL bret = FALSE;
    char buff[1024];
	
    if (CreateArchive(Name))
    {
        ret = load_txt_file(Name);
        //if (DBG_MODE) logAddLine("LoadTextFile. ret <%s>", ret);
        bret = TRUE;
	}
	
    if (DBG_MODE) logAddLine("LoadTextFile. bret <%d>", bret);
	return bret;
}

/**
*	Аналог питоновского split
*/
BOOL DividePattern(char *str, icStringCollection *m_buff, char *sdiv)
{
	char *buffS = NULL;
    BOOL bExit = FALSE;
	int cursor = 0;
	int nins;
    int old_nins = -1;

  	m_buff->RemoveAll();
    
	if (strempty(str))
		return TRUE;

    while (!bExit)
    {
		nins = strfind_offset(str, sdiv, cursor);
		if (nins == -1)	 
        {
            //не наиден разделитель
			bExit = TRUE;
			nins = strlen(str);
		}
			
		buffS = strinit(buffS, substr(str, cursor, nins-cursor));  // очередной сигнал

		if (nins > old_nins + 1)	// от промежуточных пробелов избавились
		{
			m_buff->AppendValue(buffS);	  
			cursor = nins + 1;
		}
		else
			cursor++; 

		old_nins = nins;
        
        buffS = strfree(buffS);
	}
	return TRUE;
}

/**
*   Сложный разбор
*/
BOOL DividePatternAlgbr(char *str, icStringCollection *m_buff, 
                        icStringCollection *mDiv)
{
	char *buffS = NULL;
    char *s1 = NULL;
	int cursor = 0;
    int cursor1 = 0;
    int n;
    int bound = mDiv->GetLength();
    int ndiv;
    
    //m_buff->PrintItemValues();
	m_buff->RemoveAll();
    
	int len = strlen(str);
	BOOL bFind;
    BOOL bLast = FALSE;
    BOOL bAllSpace = TRUE;
    BOOL bDivSpace;
    BOOL bOldDivSpace=FALSE;
	BOOL bEmptyCell = FALSE;
	
	if (cursor1 == len-1)
		bLast = TRUE;

    while (cursor1 < len)
    {
		bFind = FALSE;
		ndiv =-1;
		for (n = 0; n < bound; n++)
		{
            char *v = mDiv->GetIndexValue(n);
			s1 = strinit(s1, substr(str, cursor1, strlen(v)));
            
			if (strequal(s1, v))
			{
				bFind = TRUE;
				ndiv = n;
                v = strfree(v);
				break;
			}
            v = strfree(v);
		}

		if (bFind || bLast)
		{	
			bDivSpace = FALSE;
			if (ndiv >= 0)
			{
				s1 = strinit(s1, mDiv->GetIndexValue(ndiv));
				s1 = strtrim(s1, TRUE);

				if (strempty(s1))
					bDivSpace = TRUE;
			}

			if (bLast && !bFind)
			{
				buffS = strinit(buffS, substr(str, cursor, strlen(str)-cursor));
				cursor1 = len;
			}
			else
			{
				buffS = strinit(buffS, substr(str, cursor, cursor1-cursor));
				cursor = cursor1 + strlen(mDiv->GetIndexValue(ndiv));
				cursor1 = cursor;
			}
		   
			if (!strempty(buffS))
			{
				m_buff->AppendValue(buffS);	  
				bEmptyCell = FALSE;
			}
			else if (bEmptyCell && !bDivSpace)
            {
                char *v = strgen(' ', 1);
                m_buff->AppendValue(v);
                v = strfree(v);
                
                bEmptyCell = FALSE;
            }

			if (!bDivSpace)
				bEmptyCell = TRUE;
		}
		else
		{
			
			if (str[cursor1] != 0x20)
				bEmptyCell = FALSE;

			cursor1++;
		}
		
		//
		if (cursor1 == len-1)
			bLast = TRUE;
	}
    
    s1 = strfree(s1);
    buffS = strfree(buffS);
	return TRUE;
}


/**
*   ф-ия по шаблону формирует массив параметров столбцов
*/
BOOL SetWidthCol(char *strLine, int **parColl, int &numColl)
{
	icStringCollection m_buff = icStringCollection();
	
	if ((*parColl) != NULL)
		delete (*parColl);

	if (strlen(strLine) > 0)
	{
		DividePattern(strLine, &m_buff,  " ");
		numColl = m_buff.GetLength();
		if (numColl > 0)
		{
			//(*parColl) = new int[numColl];
			(*parColl) = (int*) calloc(numColl, sizeof(int));
			for (int n=0; n < numColl; n++)
            {
                char *v = m_buff.GetIndexValue(n);
				(*parColl)[n] = strlen(v);
                v = strfree(v);
            }
		}
	}
    
    //Чистим память
    m_buff.RemoveAll();
    
	return TRUE;
}


/**
*	Разбивает строчку по столбцам 
*	str - строка разбора
*	m_buff - массив котровый содержит разобранную строку
*	par - массив, котроый содержит ширины столбцов
*	numPar - количество столбцов
*/
BOOL DivideFixPattern(char *str, icStringCollection *m_buff, int *par, int numPar)
{
	if (par == NULL || numPar < 0 )
		return FALSE;

	if (numPar == 0 )
		return TRUE;
        
    m_buff->RemoveAll();
    // if (DBG_MODE) m_buff->PrintItemValues();

	int cursor = 0;
	for (int n=0; n < numPar; n++)
	{
		if (cursor <= strlen(str))
		{
			char *s = substr(str, cursor, par[n]);
			m_buff->AppendValue(s);
            s = strfree(s);
			cursor = cursor + par[n] + 1;  // +1 на пробел
		}
		else
        {
            char *s = strgen_empty();
			m_buff->AppendValue(s);
            s = strfree(s);
        }
	}
	return TRUE;
}



/**
*	Ver 2.54 [27-03-2002] 14:37 
*
*	Разбивает строчку по столбцам с учетом того, что тэги форматирования
*	также должны попасть в массив разбора
*	str    - строка разбора
*	m_buff - массив котровый содержит разобранную строку
*	par    - массив, котроый содержит ширины столбцов
*	numPar - количество столбцов
*	pcfg   - указатель на структуру настроек разбора
*/
BOOL DivideFixTagPattern(char *str, icStringCollection *m_buff, 
						 int *par, int numPar, icStructCFG *pcfg)
{
	if (par == NULL || numPar < 0)
		return FALSE;

	if (numPar == 0 )
		return TRUE;

    m_buff->RemoveAll();

	int cursor = 0;
	int width, j, beg, end;
	char *strbuff = NULL;
    char *tagbeg = NULL;
    char *tagend = NULL;

	for (int n=0; n < numPar; n++)
	{
		if (cursor <= strlen(str))
		{
			//	Ширина очередного столбца
			width = par[n];

			//	Получаем не разобранную часть
			strbuff = strinit(strbuff, substr(str, cursor, strlen(str)-cursor));
						
			// Все стандартные тэги форматирования
			j=0;

			while (AllTags[j].Seq != NULL)
			{
				tagbeg = strinit(tagbeg, strcopy(AllTags[j].Seq));
				tagend = strinit(tagend, strcopy(AllTags[j].SeqEnd));

				beg = strfind(strbuff, AllTags[j].Seq);
				end = strfind(strbuff, AllTags[j].SeqEnd);

				//	Проверяем есть ли в разбираемом столбце стандартные
				//	тэг форматирования
				if (beg < width && beg > -1)
					width += strlen(tagbeg);
				
				if (end < width && end > -1)
					width += strlen(tagend);

				j++;
			}
			
			//	Удаляем из текста внешние тэги, описанные в *.cfg			
			if (pcfg != NULL)
			{
				for (j=0; j < pcfg->mTagFormat.GetCount(); j++)
				{
					tagbeg = strinit(tagbeg, strcopy(((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(j))->sBeg));
					tagend = strinit(tagend, strcopy(((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(j))->sEnd));

					beg = strfind(strbuff, tagbeg);
					end = strfind(strbuff, tagend);

					//	Проверяем есть ли в разбираемом столбце внешние
					//	тэг форматирования
					if (beg < width && beg > -1)
						width += strlen(tagbeg);
					
					if (end < width && end > -1)
						width += strlen(tagend);                    
				}

				//////////////////////////////////////////////
				//	Проверяем тэг скрытия строк
				tagbeg = replaceprintf(tagbeg, "<%s>", pcfg->strTagHideRow);
				tagend = replaceprintf(tagend, "</%s>", pcfg->strTagHideRow);

				beg = strfind(strbuff, tagbeg);
				end = strfind(strbuff, tagend);

				//	Проверяем есть ли в разбираемом столбце стандартный
				//	тэг форматирования
				if (beg < width && beg > -1)
					width += strlen(tagbeg);
				
				if (end < width && end > -1)
					width += strlen(tagend);
			}

			char *s = substr(str, cursor, width);
			m_buff->AppendValue(s);
            s = strfree(s);
			cursor = cursor + width + 1;  // +1 на пробел
		}
		else
        {
            char *v = strgen_empty();
			m_buff->AppendValue(v);
            v = strfree(v);
        }
	}
    
    //Чистим
    tagbeg = strfree(tagbeg);
    tagend = strfree(tagend);
    strbuff = strfree(strbuff);
    
	return TRUE;
}


/**
*	Является fileName файлом?
*/
BOOL isFile(char *fileName)
{
  if (strlen(fileName) == 0)
    return FALSE;
  return file_exists(fileName);
}


/**
*	Является dirName директорией?
*/
BOOL isDir(char *dirName)
{
  if (strlen(dirName) == 0)
    return FALSE;
  return dir_exists(dirName);
}


/**
*	Определяет путь до файла
*/
char *getFilePath(char *fileName)
{
    if (!fileName)
    {
        if (DBG_MODE) errBox("getFilePath. Argument <fileName> is NULL");
        return NULL;
    }
        
	char *fileDir = NULL;
    fileName = strreverse(fileName);
	int index = strfind(fileName, "/");

	if (index >= 0)
	{
        fileName = strreverse(fileName);
        fileDir = strleft(fileName, strlen(fileName) - index - 1, FALSE);
	}

	return fileDir;
}


/**
*	Определяет имя файла с расширением
*/
char *getFileName(char *fileName)
{
    fileName = strreverse(fileName);
    int index = strfind(fileName, "/");
    fileName = strreverse(fileName);

	if (index >= 0)
		fileName = strright(fileName, index);
	
	return fileName;
}


/**
*	Определяет полный путь до файла, по имени
*	imageFileName - имя файла по которому определяется абсолютный путь
*/
char *MakeFileName(char *fileName, char *imageFileName)
{
    char *new_filename = NULL;
	if (!file_exists(fileName))
    {
        char *path_name = getFilePath(imageFileName);
        char *file_name = getFileName(fileName);
		new_filename = dupprintf( "%s/%s", path_name, file_name);
        path_name = strfree(path_name);
        file_name = strfree(file_name);
    }

	return new_filename;
}

/**
*   Произвести замены в пути до файла
*/
char *replacePath(char *Path, icStringCollection *strWord, icStringCollection *strReplace, int iCount)
{
    char *result_path = NULL;
    
    //Проверка входных данных
    if (!Path || !strWord || !strReplace)
        return NULL;

    if (iCount < 0)
        iCount = strWord->GetLength();
        
    if (!iCount)
        return Path;    
    
    //Прежде чем начать замены необходимо привести путь к  unix виду
    result_path = dos_to_unix_path(Path, TRUE);
    
    for (unsigned int i=0; i < iCount; i++)
    {
        //Получить дупликаты строк
        char *word = strWord->GetIndexValue(i);
        char *replace_word = strReplace->GetIndexValue(i);
        
        //if (DBG_MODE) logAddLine("<%s> [%s] : [%s] %d", result_path, word, replace_word, i);
        result_path = strreplace(result_path, word, replace_word, TRUE);
        
        //Освободить память
        word = strfree(word);
        replace_word = strfree(replace_word);
    }
    return result_path;
}

