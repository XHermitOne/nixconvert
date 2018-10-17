
#include "econvert.h"


char errorDescription[MAX_LOG_MSG];


/**
*	Функция разделяет сточку типа "abc:def" -> на две - "abc" и "def"
*/
char *divideToCellBeg(char *str)
{
    return strleft_to(str, ':');
}

char *divideToCellEnd(char *str)
{
    return strright_to(str, ':');
}

/**
*	Объект - синоним названия столбца или строчки
*/
class icSCell 
{
public:
	int shift;					//	сдвиг от верхнего левого угла таблицы
	char *sCell = NULL;         //  реальное название ячейки, 
	icSCell(int x) {shift = x;}
	~icSCell() { sCell = strfree(sCell);}
};

/**
*	Таблицы синонимов названий столбцов и строк 
*/
icDict m_RowMap = icDict();
icDict m_ColMap = icDict();

/**
*
*/
BOOL LookupMap(icDict *Map, char *sKey, BOOL Return)
{
    if (Map->HasKey(sKey))
        return Return;        
    return FALSE;
}

/**
*	Функция чистит таблицы синонимов
*/
void ClearSynonimsMap()
{
	m_RowMap.RemoveAll();
	m_ColMap.RemoveAll();
}

/**
*	Функция инициализирует таблицу синонимов 
*	sLTCell - название левой верхней ячейки таблицы
*/
void  InitSynonimsMap(lxw_worksheet *oSheet, char *sLTCell)
{
	char *wordCrd = "+";
    BOOL covTrue = TRUE;
    BOOL covFalse = FALSE;
	icSCell *cX;
	int countEmptyCell;
    char *cell_address;
    char *old_cell_address = NULL;

	int step = 0;

	//	Чистим таблицу синонимов
	ClearSynonimsMap();
	
	/////////////////////////////////////////////////////////////////////////////////////
	//	Устанавливаем начало таблицы
    if (!strempty(sLTCell))
        cell_address = sLTCell;
    else
        cell_address = "A1";

	//	Счетчик пустых ячеек
	countEmptyCell = 0;

	// Ищем нужный столбец (X) в самой таблице
	while (countEmptyCell < 2)
	{
		step++;
			
		//	При попытке передвинутся с последней правой ячейки вправо генерируется
		//	исключение. Если ловим его, то выходим из цикла
        old_cell_address = cell_address;
        cell_address = getNextCellAddress(oSheet, old_cell_address, 1);
        if (cell_address == NULL)
            break;
        strfree(old_cell_address);

        wordCrd = getCellValue(oSheet, cell_address);
		wordCrd = strtrim(wordCrd);
				
		// добавляем новый столбец
		if (!strequal(wordCrd, ""))
		{
			//		Заносим найденный сдвиг от начала Excel-таблицы
			//	в таблицу синонимов названий столбцов
			cX = new icSCell(step);
            cX->sCell = cell_address;
			m_ColMap.Set(wordCrd, cX);
			countEmptyCell = 0;
		}
		else 
			countEmptyCell++;
	}

	// Ищем нужную строку (Y)
	step = 0;
    cell_address = sLTCell;
	wordCrd = "+";
	
	//	Счетчик пустых ячеек
	countEmptyCell = 0;

	// Ищем нужный столбец (Y) в самой таблице
	while (countEmptyCell < 2)
	{
		step++;

		//	При попытке передвинутся с последней нижней ячейки вниз генерируется
		//	исключение. Если ловим его, то выходим из цикла
        old_cell_address = cell_address;
        cell_address = getNextCellAddress(oSheet, old_cell_address, 1);
        if (cell_address == NULL)
            break;
        strfree(old_cell_address);

		wordCrd = getCellValue(oSheet, cell_address);
		wordCrd = strtrim(wordCrd);
				
		// добавляем новый столбец
		if (!strequal(wordCrd, ""))
		{
			//		Заносим найденный сдвиг от начала Excel-таблицы
			//	в таблицу синонимов названий столбцов
			cX = new icSCell(step);
            cX->sCell = cell_address;
			m_RowMap.Set(wordCrd, cX);
			countEmptyCell = 0;
		}
		else
			countEmptyCell++;
	}
}


/**
*		Функция преобразует логические коородинаты в
*	координаты Excel листа
*		Если в имени ячейки находим ":", то надо использовать логическую систему 
*	координат, задаваемую таблицей. В поле Prim должна быть указана координата
*   верхнего левого угла таблицы. Особенности:
*		1.  Если координата угла в примечании не указана, то считаем, что верхний
*			 левых угол находится в ячейки A1
*		2.	Поиск нужной ячейки в таблице происходит следующим образом:
*			-  первое слово имени ячейки кодирует название столбца, вторая - строки
*		3.	Если название столбца в таблице отсутствует, то данный столбец в таблицу 
*			 будет добавлен. Аналогично со строками. 
*      (16:05:2001 10:25)
*      4.  В фигурных скобках {} указывается смещение. {+n}, {n} - в положительном 
*           направлении {-n} - в отрицательном
*/
int TransformLogCrd(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    char *sCell, char *sLTCell, char *realRangeAddress)
{
	int numF;
	int shiftX = 0;
    int shiftY = 0;
	BOOL bFindShift = FALSE;    // флаг того, что указано смещение

	char *XWord;
    char *YWord;

	//	В ассоцитивных координатах определяем названия 
	//	столбца и строки

	if ((numF = strfind(sCell, simbAssocCrd)) >= 0)
	{
		//		Выделяем название столбца, строчки и смещения
		if (numF > 0 && numF+1 < strlen(sCell))
		{
			XWord = strleft(sCell, numF);
			YWord = strright_pos(sCell, numF+2);

			//	Пытаемся вычленить смещения записанные в "{}"
			int numS;
		
			//	Определяем сдвиг по X
			if ((numS = strfind(XWord, "{")) >= 0)
			{
				char *sShift = strright_pos(XWord, numS+1);
				shiftX = atoi(sShift);
				XWord = strleft(XWord, numS);
				bFindShift = TRUE;
			}

			//	Определяем сдвиг по Y
			if ((numS = strfind(YWord, "{")) >= 0)
			{
				char *sShift = strright_pos(YWord, numS+1);
				shiftY = atoi(sShift);
				YWord = strleft(YWord, numS);
				bFindShift = TRUE;
			}
		}	
		else
			return -1;

		/////////////////////////////////////////////////////////////////////////////////////
		//	Устанавливаем начало таблицы
	
		char *wordCrd = "+";
		BOOL bFindColl = FALSE, bFindRow = FALSE, bExit = FALSE;
		int step=0, stepX = 0, stepY = 0;
		icSCell *cX, *cY;
        char *cell_address = NULL;

		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//	Пытаемся найти Excel названия столбца по таблице синонимов

		if (!LookupMap(&m_ColMap, XWord, (cX && !strempty(XWord))))
		{
			
			//	Устанавливаем начало таблицы
            if (!strempty(sLTCell))
                cell_address = sLTCell;
            else
                cell_address = "A1";
                
			step = m_ColMap.GetCount();
            cell_address = getNextCellAddress(oSheet, cell_address, ++step);
            if (cell_address == NULL)
                return -1;
			
			// добавляем новый столбец
            setCellValue(oBook, oSheet, cell_address, XWord);
			stepX = step;
			
			//		Заносим найденный сдвиг от начала Excel-таблицы
			//	в таблицу синонимов названий столбцов

			cX = new icSCell(stepX);
			cX->sCell = cell_address;
			m_ColMap.Set(XWord, cX);
		}

		// Ищем нужную строку (Y)
		wordCrd = "+";
		bExit = FALSE;
		step = 0;

		if (!LookupMap(&m_RowMap, YWord, (cY && !strempty(YWord))))
		{	
			step = m_RowMap.GetCount();
            cell_address = getNextCellAddress(oSheet, sLTCell, ++step);
            if (cell_address == NULL)
                return -1;
            
            setCellValue(oBook, oSheet, cell_address, YWord);
			stepY = step;
		
			//		Заносим найденный сдвиг от начала Excel-таблицы
			//	в таблицу синонимов названий строк
			cY = new icSCell(stepY);
            cY->sCell = cell_address;
			m_RowMap.Set(YWord, cY);
		}

		//	Для однообразия использования GetOffset
		if (strempty(XWord))
			cX = new icSCell(0);

		if (strempty(YWord))
			cY = new icSCell(0);

		if ((cX->shift > 0 && cY->shift >0) || bFindShift)
		{
            //	Определяем адресс ячейки
            int numX = strfind(cX->sCell, "$");
            int numY = strfind(cY->sCell, "$");
				
            if (numX >= 0 && numY >= 0)
            {
                char *sColX;
                char *sColY;
                char *sRowX;
                char *sRowY;
                char *sCell;

                //	Координата нужной ячейки вычисляется как sColX + sRowY
                //	Определяем координаты ячейки, задающей колонку
                sColX = strleft(cX->sCell, numX);

                //	Определяем координаты ячейки, задающей ряд
                sRowY = strright_pos(cY->sCell, numY);
				
                //	Ороедееляем базовую ячейку
                sCell = strprintf(sCell, "%s%s", sColX, sRowY);
					
                //	Делаем смещение от базовой ячейки, которое задается в {}
                realRangeAddress = getOffsetCellAddress(oSheet, sCell, shiftY, shiftX);
            }
            else
            {
                errBox("ERROR: Внутренняя ошибка адрессации (не найден разделитель '$').");
                return -1;
            }
            //
            //	Данный способ работает не правильно, если есть объединенные ячейки
            //
		}
		else  // ошибка адрессации
			return -1;

	}
	else	// обычная адрессация 
		return 0;

	return 1;
}

/**
*   Заполнение массива значениями
*/
icStringCollection *FillSafeArray(char *line, int iRow, int iCol)
{
    icStringCollection *collection = new icStringCollection();

    collection->AppendValue(line);
    return collection;
} 

/**
*   Функция конвертации из буфферного файла в Excel
*/
void convert(icStructCFG *pcfg, char *nmFile, char *fileScript)
{
	//	Открываем dbf файл
	dbFile oFile;
	char *buff = NULL;
	oFile.Open(nmFile);
	
	if (!oFile.isOpen())
	{
		errBox("Нет возможности открыть файл: <%s>", nmFile);
		return;
	}

    if (DBG_MODE) logAddLine("Запуск конвертации по файлу <%s>", nmFile);

	//	Читаем буффер dbf файла
	oFile.readRecords(0, -1);

    lxw_workbook *oBook = NULL;
    lxw_worksheet *oSheet = NULL;

    char *FNameValue = NULL;
    char *ListValue = NULL;
    char *CellValue = NULL;
    char *AccountValue = NULL;
    char *OldName = NULL;
    char *sPrim = NULL;
    char *old_sPrim = NULL;
	char *oldSheet = NULL;
    char *s1 = NULL;
    icStringCollection m_buff = icStringCollection();

	short nFields = oFile.nFields;  
    
	long count_rec = 0;
    BOOL FL_Create = TRUE;
    BOOL bOk = TRUE;
    BOOL covTrue = TRUE;
    BOOL covFalse = FALSE;

    //---------------------------------------------------------------------
    //	Запускаем скрипт предобработки
    if (!strempty(fileScript))
    {
        icStringCollection mChangedList = icStringCollection();
        RunScript(fileScript, "", nmFile, 0, &mChangedList, oBook, FALSE);
    }
    //---------------------------------------------------------------------
	
    oFile.gotop();
    FNameValue = strinit(FNameValue, oFile.getField("FNAME"));
    ListValue = strinit(ListValue, oFile.getField("TABLE"));
    ListValue = cp866_to_utf8(ListValue, TRUE);
    CellValue = strinit(CellValue, oFile.getField("CELL"));
    AccountValue = strinit(AccountValue, oFile.getField("ACCOUNT"));
    AccountValue = cp866_to_utf8(AccountValue, TRUE);
    sPrim = strinit(sPrim, oFile.getField("PRIM"));

    //Произвести замены в пути
    FNameValue = replacePath(FNameValue, &(pcfg->strPathWord), &(pcfg->strPathReplace), -1);
    
	char *Cell1 = NULL;
    char *Cell2 = NULL;
    char *firstWord = NULL;
    char *cmdWord = NULL;

    //	Если имя листа не определено, то открываем книгу и выходим из функции
    if (strempty(ListValue) && strempty(AccountValue) && strempty(sPrim) &&	oFile.getRecCount() == 1)
        bOk = FALSE;
	
    //	Если Excel файл существует, открываем его
    if (file_exists(FNameValue))
    {   
        if (DBG_MODE) logAddLine("DBF convert. XLS File <%s>", FNameValue);
        oBook = open_workbook_xls(FNameValue);
        oSheet = get_worksheet_by_idx(oBook, 0);
        bOk = TRUE;
    }
    else
    {
        if (DBG_MODE) logAddLine("WARNING. XLS File <%s> not exists!", FNameValue);
        bOk = FALSE;
    }

		
    if (bOk)
    {
        //	Чистим лист, если находим специальное слово "clean"
        sPrim = strlwr_lat(sPrim);
        if (strequal(sPrim, "clean"))
        {
            cleanSheet(oSheet);            
        }

        old_sPrim = NULL;
        char *strError = NULL;
        int num;

        //	Признак того, что изменилось название листа или книги
        BOOL bChangeList = FALSE;		

        //	Запоминаем имя открытой книги и листа
        OldName = strinit(OldName, strcopy(FNameValue));
        oldSheet = strinit(oldSheet, strcopy(ListValue));
        char *buffCellVal = NULL;

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // цикл заполнения
        while (!oFile.IsEOF())
        {
				
			//	Читаем данные из DBF таблицы
			FNameValue = strinit(FNameValue, oFile.getField("FNAME"));
			ListValue = strinit(ListValue, oFile.getField("TABLE"));
            ListValue = cp866_to_utf8(ListValue, TRUE);
			CellValue = strinit(CellValue, oFile.getField("CELL"));
			AccountValue = strinit(AccountValue, oFile.getField("ACCOUNT"));
			AccountValue = cp866_to_utf8(AccountValue, TRUE);
			sPrim = strinit(sPrim, oFile.getField("PRIM"));

            //Произвести замены в пути
            FNameValue = replacePath(FNameValue, &(pcfg->strPathWord), &(pcfg->strPathReplace), -1);

			//	Если изменилось имя файла, то закрываем текущую 
			//	книгу и открываем нужную
			bChangeList = FALSE;

			if (!strequal(OldName, FNameValue))
			{
                close_workbook_xlsx(oBook);
                if (!file_exists(FNameValue))
                {
                    FNameValue = strfree(FNameValue);
                    ListValue = strfree(ListValue);
                    CellValue = strfree(CellValue);
                    AccountValue = strfree(AccountValue);
                    sPrim = strfree(sPrim);
					return;
                }
                else
                {
                    oBook = open_workbook_xls(FNameValue);
                    oSheet = get_worksheet_by_idx(oBook, 0);
                    bOk = TRUE;
                }

				bChangeList = TRUE;
			}

			//	Если изменилось имя листа, то открываем нужный лист
			if (!strequal(oldSheet, ListValue))
			{
				// проверяем существует ли такой лист
                oSheet = get_worksheet_by_name(oBook, ListValue);
                if (!oSheet)
                {
                    // добавляем новый с таким именем
                    oSheet = workbook_add_worksheet(oBook, ListValue);
                }
					
				bChangeList = TRUE;
				oldSheet = strinit(oldSheet, strcopy(ListValue));

				//	Для того, чтобы перегрузился ассоциативный массив, если
				// координаты левого верхнего угла случайно не изменятся
				old_sPrim = strinit(old_sPrim, strgen_empty());
            }

			OldName = strinit(OldName, strcopy(FNameValue));
			////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//		  Примечание в режиме ассоциативной адресации
			//	содержит адрес левого верхнего угла и может содержать (после ":")
			//	определенную команду для действий над ячейкой, строкой и столбцом
			//	Пример: "A1 : RowHide"

			//	Выделяем две части: примечание и команду
			if ((num = strfind(sPrim, simbCmd)) >= 0)
			{
				firstWord = strinit(firstWord, strleft(sPrim, num));
				cmdWord = strinit(cmdWord, strright_pos(sPrim, num+1));
				cmdWord = strtrim(cmdWord, TRUE);
			}
			else
			{
				firstWord = strinit(firstWord, strcopy(sPrim));
				cmdWord = strinit(cmdWord, strgen_empty());
			}

			//	Удаляем пробелы и переводим в верхний регистр
			firstWord = strtrim(firstWord, TRUE);
			firstWord = strupr_lat(firstWord);

			//	Если находим, что:
			//	1.	координаты верхнего левого угла (sPrim) 
			//		Excel таблицы изменились и координаты ячейки записаны в ассоциативной
			//		адрессации;
			//	2.	изменилось название книги или листа (в аналоговой адрессации),
			//	то чистим буфферные таблицы названий синонимов.
			//	Если sPrim == _T(""), то считаем, что координаты не изменились
				
			if ((!strequal(old_sPrim, firstWord) && !strequal(firstWord, "") &&
				strfind(CellValue, simbAssocCrd) >= 0) || 
				(strfind(CellValue, simbAssocCrd) >= 0 && bChangeList))
			{
				//	Перегружаем таблицу синонимов
				InitSynonimsMap(oSheet, firstWord);    
				old_sPrim = strinit(old_sPrim, strcopy(firstWord));
			}

	//		Если в имени ячейки находим "::", то надо использовать логическую систему 
	//	координат, задаваемую таблицей. В поле Prim должна быть указана координата
	//  верхнего левого угла таблицы. Особенности:
	//		1.  Если координата угла в примечании не указана, то считаем, что верхний
	//			 левых угол находится в ячейки A1
	//		2.	Поиск нужной ячейки в таблице происходит следующим образом:
	//			-  первое слово имени ячейки кодирует название столбца, вторая - строки
	//		3.	Если название столбца в таблице отсутствует, то данный столбец в таблицу 
	//			 будет добавлен. Аналогично со строками. 

			int result = TransformLogCrd(oBook, oSheet, CellValue, firstWord, NULL);

			switch (result)
			{
				case  0:
				{	
					//	Если в имени ячейки находим ":" - это означает, что задается регион
					Cell1 = divideToCellBeg(CellValue);
					Cell2 = divideToCellEnd(CellValue);
                    AccountValue = convertDate(AccountValue);
                    AccountValue = convertNumber(AccountValue);
                    setRangeValue(oBook, oSheet, Cell1, Cell2, AccountValue);
                    Cell1 = strfree(Cell1);
                    Cell2 = strfree(Cell2);

					break;
				}
				case 1:	
				{
                    AccountValue = convertDate(AccountValue);
                    AccountValue = convertNumber(AccountValue);
                    setRangeValue(oBook, oSheet, Cell1, Cell2, AccountValue);
					break;
				}
				case -1:
					errBox("Ошибка адрессации: <%s> \r\n", CellValue);
					break;
			}

			////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//		Обрабатываем специальные команды, которые передаются в поле Prim
			//	после специального символа  simbCmd, определенного в spcSimb.h
			if (num >= 0 && result >= 0)
			{
				//	Пока считаем, что параметры не передаются
                cmdWord = strtrim(cmdWord, TRUE);
                cmdWord = strlwr_lat(cmdWord);
					
				//	Скрыть ряд
				if (strequal(cmdWord, cmdRowHide))
                    setRangeRowHidden(oSheet, Cell1, Cell2, covTrue);

				//	Открыть ряд
				if (strequal(cmdWord, cmdRowShow))
                    setRangeRowHidden(oSheet, Cell1, Cell2, covFalse);

				//	Скрыть колонку
				if (strequal(cmdWord, cmdColHide))
                    setRangeColumnHidden(oSheet, Cell1, Cell2, covTrue);

				//	Показать колонку
				if (strequal(cmdWord, cmdColShow))
                    setRangeColumnHidden(oSheet, Cell1, Cell2, covFalse);

				// Установка стиля
				s1 = strinit(s1, strcopy(cmdWord));
				if (strequal(strleft(cmdWord, 10), "textformat") || 
                    strequal(strleft(cmdWord, 2), "tf")) 
                {
					DividePattern(s1, &m_buff, "\"");
					icTagTextFormat *pClr = pcfg->ParseStyle(s1);

					// Устанавливаем стили
					//	Устанавливаем цвет шрифта
					//	Устанавливаем название шрифта
					//	Устанавливаем размер шрифта
                    setRangeFont(oBook, oSheet, 
                                 AccountValue, AccountValue, pClr->nameFont,
                                 pClr->sizeFont, pClr->clrText,
                                 pClr->bBold, pClr->bItalic, pClr->bUnderline);
                    
					//	Устанавливаем цвет фона
					if (pClr->clrBgr != -1)
					{
                        icInterior *oInterior = getRangeInterior(oSheet, AccountValue, AccountValue);
                        setRangeInterior(oSheet, AccountValue, AccountValue, oInterior->BackgroundColor);
                        delete oInterior;
                        oInterior = NULL;
					}

					delete pClr;
				}
			}

			oFile.Next();
			count_rec++;
		}

		//	Пишем в лог имя последней книги
		logAddLine("DBF convert. EXCELNAME: <%s>", FNameValue);

	} // bOk

    //	Сохраняем изменения
    if (oBook)
    {
        close_workbook_xlsx(oBook);
        oBook = NULL;
    }

    //---------------------------------------------------------------------
    //	Запускаем скрипт постобработки
    if (fileScript != "")
    {
        icStringCollection mChangedList = icStringCollection();
        RunScript(fileScript, "", nmFile, 0, &mChangedList, oBook, TRUE);
    }
    //---------------------------------------------------------------------

    //	Закрываем и чистим 
	FNameValue = strfree(FNameValue);
	ListValue = strfree(ListValue);
	CellValue = strfree(CellValue);
	AccountValue = strfree(AccountValue);
	sPrim = strfree(sPrim);
    
	Cell1 = strfree(Cell1);
    Cell2 = strfree(Cell2);
    firstWord = strfree(firstWord);
    cmdWord = strfree(cmdWord);    
    
    old_sPrim = strfree(old_sPrim);
	oldSheet = strfree(oldSheet);
    OldName = strfree(OldName);
    s1 = strfree(s1);
}


/**
*  Функция ищет тэги выделения текста и записывает их в специальный
*  буффер mfCell 
*/
char *VerSpcTag(char *buffStr, char *strcell,
			   icArray *mfCell, BOOL &bTagOpen, icStructCFG *pcfg)
{
	int n=0;
	int bound = mfCell->GetCount() - 1;
    char *str1 = NULL;
	//    Каждый начальный тэг создает допролнительную запись в массиве форматов
	//  и закрывает предудущий тэг 

    if (strempty(buffStr))
        return NULL;
        
	/////////////////////////////////
	// Италик
	////////////////////////////////

    // начальный тэг
    if (find_word(buffStr, ItalicLP[0].Seq))
    {
        if (!bTagOpen)
        {
            icArrayFormat *p = new icArrayFormat;
            mfCell->AppendItem(p);
            p->cellLT = strinit(p->cellLT, strcopy(strcell));
            bound++;
        }
        bTagOpen = TRUE;
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->bItalic = TRUE;
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->_bItalic = TRUE;
        buffStr = strreplace(buffStr, ItalicLP[0].Seq, "", TRUE);
    }

    // конечный тэг
    if (find_word(buffStr, ItalicLP[0].SeqEnd) && bTagOpen)
    {
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->_bItalic = FALSE;
        if (!((icArrayFormat *) mfCell->GetIndexItem(bound))->_bUnderline && !((icArrayFormat *) mfCell->GetIndexItem(bound))->_bBold)
        {
            ((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB = strinit(((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB, strcopy(strcell));
            bTagOpen = FALSE;
        }
        buffStr = strreplace(buffStr, ItalicLP[0].SeqEnd, "", TRUE);
    }

	////////////////////////////////////////////
	// Подчеркнутый
	////////////////////////////////////////////

    // начальный тэг
    if (find_word(buffStr, UnderlineLP[0].Seq))
    {
        if (!bTagOpen)
        {
            icArrayFormat *p = new icArrayFormat;
            mfCell->AppendItem(p);
            p->cellLT = strinit(p->cellLT, strcopy(strcell));
            bound++;
        }
        bTagOpen = TRUE;
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->bUnderline = TRUE;
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->_bUnderline = TRUE;
        buffStr = strreplace(buffStr, UnderlineLP[0].Seq, "", TRUE);
    }

    // конечный тэг
    if (find_word(buffStr, UnderlineLP[0].SeqEnd) && bTagOpen)
    {
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->_bUnderline = FALSE;
        if (!((icArrayFormat *) mfCell->GetIndexItem(bound))->_bItalic && !((icArrayFormat *) mfCell->GetIndexItem(bound))->_bBold)
        {
            ((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB = strinit(((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB, strcopy(strcell));
            bTagOpen = FALSE;
        }
        buffStr = strreplace(buffStr, UnderlineLP[0].SeqEnd, "", TRUE);
    }
	
	///////////////////////////////////////////////////////////////////
	//  Жирный шрифт и двойной удар
	///////////////////////////////////////////////////////////////////

    // начальный тэг
    if (find_word(buffStr, BoldLP[0].Seq) || find_word(buffStr, BoldLP[1].Seq))
    {
        if (!bTagOpen)
        {
            icArrayFormat *p = new icArrayFormat;
            mfCell->AppendItem(p);
            p->cellLT = strinit(p->cellLT, strcopy(strcell));
            bound++;
        }
        bTagOpen = TRUE;
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->bBold = TRUE;
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->_bBold = TRUE;
        buffStr = strreplace(buffStr, BoldLP[0].Seq, "", TRUE);
        buffStr = strreplace(buffStr, BoldLP[1].Seq, "", TRUE);
    }

    // конечный тэг
    if ((find_word(buffStr, BoldLP[0].SeqEnd) || find_word(buffStr, BoldLP[1].SeqEnd))
        && bTagOpen)
    {
        ((icArrayFormat *) mfCell->GetIndexItem(bound))->_bBold = FALSE;
        if (!((icArrayFormat *) mfCell->GetIndexItem(bound))->_bItalic && !((icArrayFormat *) mfCell->GetIndexItem(bound))->_bUnderline)
        {
            ((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB = strinit(((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB, strcopy(strcell));
            bTagOpen = FALSE;
        }
        buffStr = strreplace(buffStr, BoldLP[0].SeqEnd, "", TRUE);
        buffStr = strreplace(buffStr, BoldLP[1].SeqEnd, "", TRUE);
    }

	/////////////////////////////////////////////////
	//	Тэги настройки цветов текста
	/////////////////////////////////////////////////

	BOOL bCreate = FALSE;

	for (int i=0; i < pcfg->mTagFormat.GetCount(); i++)
	{
		// начальный тэг
		if (find_word(buffStr, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->sBeg))
		{
			if (!bCreate)
			{
				//	Закрываем предыдущий тэг
				if (bTagOpen)
					((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB = strinit(((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB, strcopy(strcell));

				//	Создаем новый
				icArrayFormat *p = new icArrayFormat;
				mfCell->AppendItem(p);
				p->cellLT = strinit(p->cellLT, strcopy(strcell));
				bound++;
				bCreate = TRUE;

                str1 = strprintf(str1, "<%s>", pcfg->strTagHideRow);
				if (strequal(str1, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->sBeg))
					p->bRowHide = TRUE;
			}
			
			bTagOpen = TRUE;
			((icArrayFormat *) mfCell->GetIndexItem(bound))->clrText = ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->clrText;
			((icArrayFormat *) mfCell->GetIndexItem(bound))->clrBgr = ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->clrBgr;
			((icArrayFormat *) mfCell->GetIndexItem(bound))->nameFont = strcopy(((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->nameFont);
			((icArrayFormat *) mfCell->GetIndexItem(bound))->sizeFont = ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->sizeFont;

			if (((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->bItalic)
				((icArrayFormat *) mfCell->GetIndexItem(bound))->bItalic = ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->bItalic;

			if (((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->bBold)
				((icArrayFormat *) mfCell->GetIndexItem(bound))->bBold = ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->bBold;

			if (((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->bUnderline)
				((icArrayFormat *) mfCell->GetIndexItem(bound))->bUnderline = ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->bUnderline;

			buffStr = strreplace(buffStr, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->sBeg, "", TRUE);
		}

		//	Конечный тэг
		if (find_word(buffStr, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->sEnd) && bTagOpen)
		{
			bTagOpen = FALSE;
			((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB = strinit(((icArrayFormat *) mfCell->GetIndexItem(bound))->cellRB, strcopy(strcell));
			buffStr = strreplace(buffStr, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(i))->sEnd, "", TRUE);
		}
	}

	//**********************************************************
	//	Обрабатываем тэг, который скрывает строчку

	BOOL bEmpty = strempty(pcfg->strTagHideRow);
	
	//	Начальный тэг	
    str1 = replaceprintf(str1, "<%s", pcfg->strTagHideRow);
	if (!bEmpty && find_word(buffStr, str1))
	{
		icArrayFormat *p = new icArrayFormat();
		mfCell->AppendItem(p);
		p->cellLT = strinit(p->cellLT, strcopy(strcell));
		p->cellRB = strinit(p->cellRB, strcopy(strcell));
		p->bRowHide = TRUE;
		bound++;
        
        str1 = replaceprintf(str1, "<%s>", pcfg->strTagHideRow);
		buffStr = strreplace(buffStr, str1, "", TRUE);
		bTagOpen = FALSE;
	}

	//	Конечный тэг	
    str1 = replaceprintf(str1, "</%s>", pcfg->strTagHideRow);
	if (!bEmpty && find_word(buffStr, str1))
		buffStr = strreplace(buffStr, str1, "", TRUE);

    str1 = strfree(str1);
	return buffStr;
}


/**
*		Функция удаляет из строчки все символы форматирования. Это делается 
*	для того, чтобы данные символы  не учавствовали при разделении текста на  
*  столбцы таблицы.
*/
char *FreeAllTag(char *str, icStructCFG *pcfg)
{
	char *strFree = strcopy(str);
    char *str1 = NULL;
	int n = 0;
	
	// Все стандартные тэги форматирования	
	while (AllTags[n].Seq != NULL)
	{
		strFree = strreplace(strFree, AllTags[n].Seq, "", TRUE);
		strFree = strreplace(strFree, AllTags[n].SeqEnd, "", TRUE);
		n++;
	}
	
	//	Удаляем из текста внешние тэги, описанные в *.cfg	
	if (pcfg != NULL)
	{
		char *strTag;

		for (n=0; n < pcfg->mTagFormat.GetCount(); n++)
		{
			//	Удаляем открывающие и закрывающие тэги			
			strFree = strreplace(strFree, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(n))->sBeg, "", TRUE);
			strFree = strreplace(strFree, ((icTagTextFormat *) pcfg->mTagFormat.GetIndexItem(n))->sEnd, "", TRUE);
		}

		//////////////////////////////////////////////
		//	Удаляем тэги скрытия строк
        str1 = dupprintf("<%s>", pcfg->strTagHideRow);
		strFree = strreplace(strFree, str1, "", TRUE);
        str1 = strfree(str1);
        
        str1 = dupprintf("<%s\\>", pcfg->strTagHideRow);
		strFree = strreplace(strFree, str1, "", TRUE);
        str1 = strfree(str1);
	}
    else  
    {  
        str1 = strfree(str1);
    }

	return strFree;
}

/**
*	Функция в строчке заменяет все символы табуляции на пробелы
*/
char *replaceTab(char *str, int iTab)
{
	int num = 0;
    int cursor = 0;
	char *strTab = strgen(0x20, iTab);

	while (num >= 0)
	{
		num = strfind_char(str, '\t', cursor);

		if (num >= 0)
		{
			// Определям количество пробелов, 
			int ost =  iTab - num + (num / iTab) * iTab;

            char *str1 = strleft(strTab, ost);
            str = strreplace_pos(str, num, str1, TRUE);
            str1 = strfree(str1);
		}
	}
    //Очистить память
    strTab = strfree(strTab);
    
    return str;
}

/**
*	Функция ищет тэги скрывающие колонки
*/
void FindHideShowColTag(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                        char *sLTCell, icArray *mfCell, icStructCFG *pcfg)
{
	char *cellText = NULL; 
    char *vrCell = NULL;
    BOOL covTrue = TRUE;
    BOOL covFalse = FALSE;
	icArrayFormat *cX;
	int countEmptyCell = 0, przEnd=256;
	int step = 0;

	/////////////////////////////////////////////////////////////////////////////////////
	//	Устанавливаем начало таблицы
    cellText = getCellValue(oSheet, sLTCell);
    if (cellText)
        cellText = strtrim(cellText, TRUE);
				
    // добавляем новый столбец
    if (!strempty(cellText))
    {
        //	Если находим тэг скрытия колонки, то в буфер форматированных 
        //	ячеек добавляем соответствующий объект формата 
        if (find_word(cellText, pcfg->strTagHideCol))
        {
            cX = new icArrayFormat();
            cX->bColHide = TRUE;
            cX->cellLT = sLTCell;
            cX->cellRB = cX->cellLT;
            mfCell->AppendItem(cX);
            cellText = strreplace(cellText, pcfg->strTagHideCol, "", TRUE);
            setCellValue(oBook, oSheet, sLTCell, cellText);
        }
        //	Если находим тэг открытия колонки, то в буфер форматированных 
        //	ячеек добавляем соответствующий объект формата 
        else if (find_word(cellText, pcfg->strTagShowCol))
        {
            cX = new icArrayFormat();
            cX->bColShow = TRUE;
            cX->cellLT = sLTCell;
            cX->cellRB = cX->cellLT;
            mfCell->AppendItem(cX);
            cellText = strreplace(cellText, pcfg->strTagShowCol, "", TRUE);
            setCellValue(oBook, oSheet, sLTCell, cellText);
        }
    }
    else 
        countEmptyCell++;

	// Ищем нужный столбец (X) в самой таблице
    char *next_addr = sLTCell;
	while (countEmptyCell < przEnd)
	{
		step++;
			
		//	При попытке передвинутся с последней правой ячейки вправо генерируется
		//	исключение. Если ловим его, то выходим из цикла
        next_addr = getNextCellAddress(oSheet, next_addr);
        if (next_addr == NULL)
            break;

		vrCell = getCellValue(oSheet, next_addr); //oRange.GetText();
		cellText = strcopy(vrCell); //.bstrVal);
		cellText = strtrim(cellText);
				
		// добавляем новый столбец
		if (!strempty(cellText))
		{
			//	Если находим тэг скрытия колонки, то в буфер форматированных 
			//	ячеек добавляем соответствующий объект формата 
			if (find_word(cellText, pcfg->strTagHideCol))
			{
				cX = new icArrayFormat();
				cX->bColHide = TRUE;
                cX->cellLT = next_addr;
				cX->cellRB = cX->cellLT;
				mfCell->AppendItem(cX);
				cellText = strreplace(cellText, pcfg->strTagHideCol, "", TRUE);
                setCellValue(oBook, oSheet, next_addr, cellText);
			}
			//	Если находим тэг открытия колонки, то в буфер форматированных 
			//	ячеек добавляем соответствующий объект формата 
			else if (find_word(cellText, pcfg->strTagShowCol))
			{
				cX = new icArrayFormat();
				cX->bColShow = TRUE;
                cX->cellLT = next_addr;
				cX->cellRB = cX->cellLT;
                mfCell->AppendItem(cX);
                cellText = strreplace(cellText, pcfg->strTagShowCol, "", TRUE);
                setCellValue(oBook, oSheet, next_addr, cellText);
			}
			countEmptyCell = 0;
		}
		else 
			countEmptyCell++;
	}

}

/**
* Функция ожидает пока не будет разблокирована ячейка
*/
BOOL WaitRangeLock(void)
{
    return FALSE;
}


/**
*   Функция генерирует список имен колонок
*/
static char *_gen_column_names_str(void)
{
	char *sAlf = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    unsigned int alf_len = strlen(sAlf);
    char *strCell = (char *) calloc(513, sizeof(char));
	int firstLetter = 0;
    int secondLetter = -1;
	char sL1 = ' ';
    char sL2 = ' ';

	////////////////////////////////////////////////////////////////////////////////////////////////////////
	//	Заполняем строчку возможными названиями столбцов
	for (int i=0; i < 256; i++)
	{
		sL1 = sAlf[firstLetter];

		if (secondLetter >= 0)
            sL2 = sAlf[secondLetter];
        strCell[i * 2] = sL2;
        strCell[i * 2 + 1] = sL1;

		firstLetter++;

		if (firstLetter == alf_len)
		{
			secondLetter++;
			firstLetter = 0;
		}
	}
    
    return strCell;
}


/**
*   Заменить подстроки на пробелы
*/
static char *_replaceSpace(char *str, icStringCollection *lReplace)
{
    char *strEmpty = "                       ";
    for(int i=0; i < lReplace->GetLength(); i++)
    {
        char *str1 = lReplace->GetIndexValue(i);
        char *str2 = strleft(strEmpty, strlen(str1));
        str = strreplace(str, str1, str2, TRUE);
        str1 = strfree(str1);
        str2 = strfree(str2);
    }
    return str;
}

static BOOL _clearAfterHead(lxw_worksheet *oSheet, unsigned int startLine, int numCol, int step, char *strCell)
{
    char *cell1 = NULL;
    char *cell2 = NULL;
    unsigned int start_line = startLine+2;
    unsigned int stop_line = 0;

    cell1 = strprintf(cell1, "A%d", start_line);

    //	Если строка форматирования есть, то чистим только 
    //	область таблицы
    if (numCol > 0)
    {
        char *sRightCell = substr(strCell, (numCol + step - 1)*2, 2);
        sRightCell = strtrim(sRightCell, TRUE);
					
        stop_line = (oSheet->dim_rowmax > start_line) ? oSheet->dim_rowmax : 0 ;
        if (stop_line > 0)
            cell2 = strprintf(cell2, "%s%d", sRightCell, stop_line);
        else
            //Не надо чистить 
            cell2 = strfree(cell2);
            
        sRightCell = strfree(sRightCell);
    }
    //	В противном случае чистим всю область после шапки
    else
        cell2 = strprintf(cell2, "IV%d", startLine+1);
			
    //	Ищем тэги скрытия колонок
            
    //	Чистим первую строчку так, чтобы не повредить числовые
    //	форматы столбцов
				
    //	Остальную область полностью чистим
    if (cell2)
        cleanRange(oSheet, cell1, cell2);
        
    //Чистим память
    cell1 = strfree(cell1);
    cell2 = strfree(cell2);
    
    return TRUE;
}

/**
*	Ищим специальные слова признаки
*/
static BOOL _is_prizn_word(char *sLine, icStructCFG *pcfg)
{
    BOOL ret = FALSE;
    for (int i=0; i < pcfg->m_sPrizn.GetLength(); i++)
    {
        char *v = pcfg->m_sPrizn.GetIndexValue(i);
        if (find_word(sLine, v))
        {
            ret = TRUE;
            v = strfree(v);
            break;
        }
        v = strfree(v);
    }
    return ret;
}

/**
*   Определяем количество символов, на которое сдвинута таблица в право.
*       если это не учесть, то произойдет перемешивание столбцов за счет того,
*       что первые пробелы будут учитываться при разборе.
*       пример: "-" обозначены пробелы
*           - - @====== == ==============
*           - -  строчка1   ...
*/
static int _find_endline_pos(const char *line, const int start_pos)
{
    int pos;
    char *strReverse = strleft((char *) line, start_pos);
    strReverse = strreverse(strReverse);
		
    if ((pos = strfind(strReverse, "\n\r")) < 0)
        pos = start_pos;
    strReverse = strfree(strReverse);    
    return pos;
}

/**
*   Очистить массив заполнения ячеек
*/
static BOOL _free_cell_val_array(icArray *array)
{
    for (unsigned int i=0; i < array->GetCount(); i++)
    {
        if (array->Items[i] != NULL)
        {
            //ВНИМАНИЕ!
            //Здесь необходимо явное приведение к типу класса 
            //для вызова деструтора
            delete ((icCellValue *)(array->Items[i]));
        }
    }
    array->RemoveAll();      
}

static BOOL _drawTabHead(lxw_workbook *oBook, lxw_worksheet *oSheet,
                        icStringCollection *IndexFormat, 
                        char *strCell, int step, int maxCell, BOOL bHead)
{
    char *v = IndexFormat->GetIndexValue(0);
                
    char *col1 = substr(strCell, step * 2, 2);
    col1 = strtrim(col1, TRUE);
    char *cell1 = dupprintf("%s%s", col1, v);

	char *col2 = substr(strCell, (maxCell + step) * 2, 2);
    col2 = strtrim(col2, TRUE);
    char *cell2 = NULL;

    char *row = IndexFormat->GetIndexValue(1);
    if (!strempty(row))
    {
        int beg = atoi(row);
        beg--;
        cell2 = strprintf(cell2, "%s%d", col2, beg);
	}
    
    // Отрисовка шапки
    if (DBG_MODE) logAddLine("Head draw [%s : %s]", cell1, cell2);
    if (bHead)
    {
        // выделяем шапку
        setRangeBorder(oBook, oSheet, cell1, cell2, 0, LXW_BORDER_THIN, 0, LXW_BORDER_THIN);
        setRangeBorderAround(oBook, oSheet, cell1, cell2, LXW_BORDER_MEDIUM, LXW_BORDER_MEDIUM, LXW_BORDER_MEDIUM, LXW_BORDER_MEDIUM);
    }
        
    // Отрисовка дополнительных линий
    for (unsigned int n = 2; n < IndexFormat->GetLength(); n++)
    {
        row = strinit(row, IndexFormat->GetIndexValue(n));
        cell1 = replaceprintf(cell1, "%s%s", col1, row);
        cell2 = replaceprintf(cell2, "%s%s", col2, row);
        if (DBG_MODE) logAddLine("Draw line [%s : %s]", cell1, cell2);
        setRangeBorder(oBook, oSheet, cell1, cell2, LXW_BORDER_THIN, 0, 0, 0);
    }
    
    // Чистим память
    row = strfree(row);
    v = strfree(v);
    cell1 = strfree(cell1);
    cell2 = strfree(cell2);
    col1 = strfree(col1);
    col2 = strfree(col2);
}

/**
* Функция конвертирует текст на выбранный лист
*/
BOOL convertList(char *buff, char *ListValue, lxw_workbook *oBook, 
							icStructCFG *pcfg, BOOL bClear)
{
    int n;
	BOOL bTagOpen  = FALSE;
	icArray mfCell = icArray();
	lxw_worksheet *oSheet = NULL;
    
    BOOL covTrue = TRUE;
    BOOL covFalse = FALSE;
    
	//	Проверяем существует ли такой лист
    oSheet = get_worksheet_by_name(oBook, ListValue);
    if (!oSheet)
		//	Добавляем новый с таким именем
        oSheet = workbook_add_worksheet(oBook, ListValue);
    else
        worksheet_select(oSheet);

	// грузим файл
    
	char *strLine = NULL;
    char *strCell = _gen_column_names_str();
	int lenBuffStr = strlen(strCell);

	///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	char *NameCell = NULL;
    char *cell_TopLeft = NULL;
    char *cell_BottomRight = NULL;
    char *buffStr = NULL;
	char *strLineFreeTag = NULL;
    char *buffStrFreeTag = NULL;
    char *strDLine = NULL;

    char *str1 = NULL;
    
	char bch[20];
	int cursor = 0;
    int countLine = 0;
    int old_len = -1;
    int num, numColl = 0;
    
	icStringCollection m_buff = icStringCollection();
    icStringCollection m_buff2 = icStringCollection();
    icStringCollection m_buffFreeTag = icStringCollection();

	BOOL IsEOF = FALSE;
	BOOL bEndTable = FALSE, bFormatLine = FALSE, oldFormatLine = FALSE;
	BOOL bClearAfterHead = FALSE;
	BOOL bSetBuff = pcfg->bHead;
	BOOL bSpecialWord = FALSE;
	BOOL bPrevSpecialWord = FALSE;

	int *parColl = NULL;
	icStringCollection indexFormat = icStringCollection();
	int maxCell = 0, step, num_buffLine = 0, heightT = 50, old_step;
	int old_countLine;
	int numSpc, countSpc = 0;
	int iShiftCell = 0, nLineAfterHead = -1;
    
    //Признак заголовка табличной части
    BOOL bTabHead = FALSE;
    BOOL bTabHeadTrig = FALSE; 
    
    // Форматы таблицы данных
    lxw_format *tab_formats[MAX_COLUMN_COUNT] = {NULL};

	if (pcfg->bCentredTextBeforeHead)
		step = 2;
	else
		step = 1;

	old_step = step;

	//		Количество пробелов перед таблицей - определяется по 
	//	строчке форматирования	
	int numSpaceFormat = 0, __numSpaceFormat = 0;

	//	Количество полных разделительных линий в шапке, нужно для того, чтобы в режиме 
	//	с шаблоном шапка не переформатировалась
	int numDivLineInHead = 0;

	//	Счетчик игнорируемых строк в тексте отчета.
	//	Признаком конца строки является "\r\n"	
	int numIgnoreLine = 0;
	
	//	Ищем описания в шаблоне для определения, количества игнорируемых строк (numIgnoreLine)
	//  в отчете и номер строки в шаблоне (countLine), куда выгружается таблица уже без шапки
	if (!pcfg->bHead)		
	{
		char *coord = getCellValue(oSheet, "A1");
        if (!coord) 
            coord = strgen_empty();

		if (find_word(coord, "="))
		{
			countLine = atoi(substr(coord, 0, n)) - 1;
			numIgnoreLine = atoi(strright_pos(coord, n+1));
			countSpc++;

			if (countLine <= 0)
				countLine = 0;
            char *v = itoa(countLine, 10);
            indexFormat.AppendValue(v);
            v = strfree(v);

			nLineAfterHead = countLine;
		}
        coord = strfree(coord);
	}

	//	Ищем строку с которой будем вести конвертацию
	int countTxtRow = 0;

	while (!IsEOF && countTxtRow < numIgnoreLine)
	{
		num = strfind_offset(buff, "\r\n", cursor);
		countTxtRow++;
		
		if (num == -1)
		{
			IsEOF = TRUE;
			num = strlen(buff);
		}

		cursor = num + 2;
	}

	//	Обрезаем нужное количество строк
	if (IsEOF)
		buff = "";
	else if (cursor > 0)
	{
		buff = strright_pos(buff, cursor);
		bSetBuff = TRUE;
	}

	cursor = 0;
	old_countLine = countLine;

	//	Если строчка начинается с @, то форматируем по заданному шаблону	
	int pos = strfind(buff, "@-");

	if ( pos > 0)
	{
		//		Определяем количество символов, на которое сдвинута таблица в право.
		//	если это не учесть, то произойдет перемешивание столбцов за счет того,
		//	что первые пробелы будут учитываться при разборе.
		//	пример: "-" обозначены пробелы
		//  - - @====== == ==============
		//  - -  строчка1   ...
        __numSpaceFormat = _find_endline_pos(buff, pos);

		num = strfind_offset(buff, "\r\n", pos);
		if (num > 0)
		{
			// выделяем строчку, по которой разбиваем на столбцы
			strLine = strinit(strLine, substr(buff, pos, num-pos));
			
			//	В этой строке заменяем все разделители на пробелы, а затем
			//	уже разбираем
            strLine = _replaceSpace(strLine, &(pcfg->m_Div));

			SetWidthCol(strLine, &parColl, numColl);
		}
	}

	//	Чистим лист перед конвертацией		
    //	Если bHead = FALSE, чистим после определении  шапки таблицы
    if (bClear && pcfg->bHead)
        cleanSheet(oSheet);

	//////////////////////////////////////////////////////////////////////////////////////
	//
	//  Разбираем лист. Разбиваем текст на строчки и заполняем данными эксельный лист.
	//
	//////////////////////////////////////////////////////////////////////////////////////
	while (!IsEOF)
	{
		num = strfind_offset(buff, "\r\n", cursor);
			
		if (num == -1)
		{
			IsEOF = TRUE;
			num = strlen(buff);
		}

		//  Выделяем очередную строчку
		strLine = strinit(strLine, substr(buff, cursor, num-cursor));
		cursor = num + 2;
		
		///////////////////////////////////////////////
		//	Обрабатываем символы табуляции
		strLine = replaceTab(strLine, pcfg->iTab);
		
		//	Ищем вертикальные разделители. Т. к. в шапке разделители могут использоваться
		//	как замены на пробелы, поэтому вот так извращаемся
		if (numColl > 0)
		{
			char *strVer = strcopy(strLine);
			strVer = strtrim(strVer, TRUE);

			if (!strempty(strVer))
			{
                strVer = _replaceSpace(strVer, &(pcfg->m_Div));

				strVer = strreplace(strVer, "-", "", TRUE);
				strVer = strtrim(strVer, TRUE);
				
				//	Если пусто, то считаем строку вертикальным разделителем
				//	Символы '----' используем для того, чтобы ниже программа 
				//	распознала эту строку как разделитель. 
				if (strempty(strVer))
                {
                    strLine = strfree(strLine);
					strLine = strgen('-', 47);
                }
			}
            strVer = strfree(strVer);
		}
		
		//   Символ @- является признаком форматирования, а также конца форматирования
		//  @----   ----  ---  -----------   - таким образом задается форматирование.
		numSpc = strfind_offset(strLine, "@-", 0);
        
        //Установить признак заголовочной части таблицы
        if (!bTabHeadTrig)
        {
            bTabHead = find_word(strLine, "--");
            if (bTabHead)
                bTabHeadTrig = TRUE;        
        }

		if (numSpc >= 0)
		{
			//	Заменяем разделители
            strLine = _replaceSpace(strLine, &(pcfg->m_Div));

			if (countSpc == 0)
				numDivLineInHead = indexFormat.GetLength();
			
			countSpc++;
            bTabHead = FALSE;
		}

		//	Производим замены символов в основном тексте отчета, за исключением	шапки.
		//	Конец шапки определятся по следующим признакам:
		//	1.  По символу "@"
		//	2.	Если встретили две разделительные линии и символа "@" в отчете нет
		if (((indexFormat.GetLength() >= 2) && (numColl == 0)) || (countSpc > 0))
		{
			//	Чистим область листа после шапки
			//	
			if (!bClearAfterHead && !pcfg->bHead)
			{
                bClearAfterHead = _clearAfterHead(oSheet, countLine, numColl, step, strCell);

				//	Запоминаем номер линии где кончилась шапка
				if (nLineAfterHead <= 0)
					nLineAfterHead = countLine;
			}
		}
		
		//	По старому соглашению второй символы "@-" является признаком конца таблицы 
		//	Его обработка оставлена для совместимости.
		//	Для совместимости с программой печати - если строчка содержит
		//	только символ "{", то это также обозначает конец таблицы

		strDLine = strinit(strDLine, strcopy(strLine));
		strDLine = strtrim(strDLine, TRUE);

		if ((strfind(strDLine, simbEndTable) == 0) || ((numSpc >= 0) && (countSpc >= 2)) )
		{
			numSpaceFormat = 0;	//	 отменяем сдвиг 
			bEndTable = TRUE;   //  признак того, что далее идет разбор по разделительным символам
			strLine = strreplace(strLine, simbEndTable, " ", TRUE);
		}
        strDLine = strfree(strDLine);

		// Cтрочку разбиваем на ячейки в зависимости от типа разбивки
		strLine = strreplace(strLine, simbFrmStr, "--", TRUE);

		//	Учитываем сдвиг символа форматирования
		strLine = strinit(strLine, strright_pos(strLine, numSpaceFormat));

		//	Создаем строчку, в которой вырезаны все символы
		//	форматированиия		
		strLineFreeTag = strinit(strLineFreeTag, FreeAllTag(strLine, pcfg));
        //if (DBG_MODE) logAddLine("Line <%s>", cp866_to_utf8(strLineFreeTag));

		//////////////////////////////////////////////////////////////////
		//	Ищим специальные слова признаки
		//	!pcfg->bHead  - признак режима с шаблоном
		//	countSpc == 0 - означает, что шапка еще не началась 
		bSpecialWord = FALSE;
		
		if (!pcfg->bHead && countSpc == 0)
            bSpecialWord = _is_prizn_word(strLine, pcfg);

		//	Разбиваем строчку на ячейки, используя разделители, если 
		//  indexFormat.GetSize() == 0  - шапка еще не начилась либо  
		//  numColl = 0 - количество колонок таблицы не задано, либо 
		//  numSpc >= 0 - в строчке найден символ "@", признак строчки, которая задает форматирование таблицы
		//  bEndTable == TRUE - признак выхода за пределы таблицы.
		if (numColl == 0 || indexFormat.GetLength() == 0 || numSpc >= 0 || bEndTable)
		{
            char *strL = strcopy(strLine);
			DividePatternAlgbr(strLine, &m_buff, &(pcfg->m_Div));
			DividePatternAlgbr(strLineFreeTag, &m_buffFreeTag, &(pcfg->m_Div));
            // if (DBG_MODE) m_buffFreeTag.PrintItemValues();

			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//	В случае если ширина таблицы известна пытаемся правильно
			//	расположить строчку - центрировать относительно шапки
			if (numColl > 0 && pcfg->bCentredTextBeforeHead)
			{
				//	По количеству пробелов определяем сдвиг певого слова от начала строки 
				strL = strtrim(strL, TRUE);
				int shift = strlen(strLine) - strlen(strL);
                int shiftCell = 0;
				iShiftCell = 0;

				//	Определяем ячейку куда попадает первое слово
				for (int i=0; i < numColl; i++)
				{
					shiftCell += parColl[i] + 1;
					if (shiftCell > shift)
					{
						iShiftCell = i;
						break;
					}
				}

				//	В начало буффера вставляем нужное количеств пустых ячеек
                if (DBG_MODE) logAddLine("Shift cell <%d>", iShiftCell);
				for (int i=0; i < iShiftCell-1; i++)
				{
                    char *v = strgen(' ', 1);
					m_buff.InsertValue(0, v);
					m_buffFreeTag.InsertValue(0, v);
                    v = strfree(v);
				}
			}
            strL = strfree(strL);
		}
		else
		{
			// Заменяем разделяющие символы на пробелы
			int i=0;
            int bound = pcfg->m_Div.GetLength();
			
            strLine = _replaceSpace(strLine, &(pcfg->m_Div));
            strLineFreeTag = _replaceSpace(strLineFreeTag, &(pcfg->m_Div));
			
			//	Разбиваем на ячейки определенной ширины, заданной в строчке
			//	форматирования.  			
			DivideFixTagPattern(strLine, &m_buff, parColl, numColl, pcfg);
			DivideFixPattern(strLineFreeTag, &m_buffFreeTag, parColl, numColl);
            //if (DBG_MODE) logAddLine("Cur line <%s>", strLineFreeTag);
            //if (DBG_MODE) m_buff.PrintItemValues();
            //if (DBG_MODE) m_buffFreeTag.PrintItemValues();
		}

		strLine = strtrim(strLine, TRUE);
		strLineFreeTag = strtrim(strLineFreeTag, TRUE);
		
		if (strlen(strLine) == 0)
		{
            char *space = strgen(' ', 1);
			m_buff.RemoveAll();
			m_buff.AppendValue(space);
			m_buffFreeTag.RemoveAll();
			m_buffFreeTag.AppendValue(space);
            space = strfree(space);
		}
		
		//	Отлавливаем разделяющие линии		
		oldFormatLine = bFormatLine;
		bFormatLine = FALSE;
		int nrepl = strfind_count(strLine, "-");
        strLine = strreplace(strLine, "-", "", TRUE);
        
		strLineFreeTag = strreplace(strLineFreeTag , "-", "", TRUE);
        strLine = strtrim(strLine, TRUE);
        strLineFreeTag = strtrim(strLineFreeTag, TRUE);

		//	Запоминаем ряды, где находятся разделительные линии
		if (strlen(strLine) == 0 && nrepl > 0)
		{
            char *v = dupprintf("%d", countLine + 1);
			indexFormat.AppendValue(v);
            v = strfree(v);
			step = 1;
			bFormatLine = TRUE;
			numSpaceFormat = __numSpaceFormat;
		}
		else
			countLine++;

		//	Условие необходимо для того, чтобы несколько горизонтальных
		//	линий воспринимались как она линия
		//	В режиме с шаблоном, подсчет и заполнения буфферного массива 
		//	(см. ниже после формирования буфф. таблицы цикл заполнения из
		//	буффера строки m_buff) 	начинается после обнаружения шапки		
		if (!oldFormatLine && ((bClearAfterHead && !pcfg->bHead) ||  pcfg->bHead))
			num_buffLine++;

		//	Формируем или переформировываем массив буфферной таблицы в 
		//	следующих случаях:
		//	1. old_len != m_buff.GetSize() && m_buff.GetSize()> 0 - изменилась количество столбцов
		//	2. num_buffLine == heightT - буфферная таблица заполнена и ее надо выгрузить в Excel 
		//	3. bSpecialWord - обнаружено слово признак и надо отчистить старый буффер
		//	4. bPrevSpecialWord - надо отчистить буфер обновляемой строки со словом признаком
		if (((old_len != m_buff.GetLength()) && (m_buff.GetLength() > 0) )
			|| (num_buffLine == heightT) || (bSpecialWord || bPrevSpecialWord))
		{
			//	Если буффер заполнен пишем его в Excel таблицу			
			if (old_len >= 0)
			{
				cell_BottomRight = strinit(cell_BottomRight, substr(strCell, (old_len -1+old_step)*2, 2));
				cell_BottomRight = strtrim(cell_BottomRight, TRUE);

				if (num_buffLine > 0)
					cell_BottomRight = replaceprintf(cell_BottomRight, "%s%d", cell_BottomRight, old_countLine + (num_buffLine - 1));
				else
					cell_BottomRight = replaceprintf(cell_BottomRight, "%s%d", cell_BottomRight, old_countLine);

				if (bClearAfterHead && !pcfg->bHead)
					bSetBuff = TRUE;
				
			}

			old_len = m_buff.GetLength();
			num_buffLine = 0;

			cell_TopLeft = strinit(cell_TopLeft, strgen_empty());
			cell_BottomRight = strinit(cell_BottomRight, strgen_empty());
		}

		//	Устанавливаем признак того, что предыдущая строчка была обновляемой
		//	Для того, чтобы обновить буфферную таблицу на следующем кругу		
		bPrevSpecialWord = bSpecialWord;

		if (strempty(cell_TopLeft))
		{
			cell_TopLeft = strinit(cell_TopLeft, substr(strCell, step*2, 2));
			cell_TopLeft = strtrim(cell_TopLeft, TRUE);

			if (bFormatLine)
				old_countLine = countLine+1;
			else
				old_countLine = countLine;

			cell_TopLeft = replaceprintf(cell_TopLeft, "%s%d", cell_TopLeft, old_countLine);
            
			old_step = step;
		}

		// n = 0;
        int boundBuff = m_buff.GetLength();

        if (DBG_MODE) logAddLine("\t %d. Step[%d]", countLine, step);
        
		//	Цикл заполнения буфферной таблицы из буффера строки
		for (n = 0; (n < boundBuff) && (!bFormatLine) && (((n + step) * 2 + 1) < lenBuffStr); n++)
		{
			// Определяем максимальную ячейку по ширине	
			if (n > maxCell)
				maxCell = n;

			// Вычисляем название ячейки				
			NameCell = strinit(NameCell, substr(strCell, (n + step) * 2, 2));
			NameCell = strtrim(NameCell, TRUE);
			NameCell = replaceprintf(NameCell, "%s%d", NameCell, countLine);

			if (m_buffFreeTag.GetLength() >= n)
				buffStrFreeTag = strinit(buffStrFreeTag, m_buffFreeTag.GetIndexValue(n)); 
			else
				buffStrFreeTag = strinit(buffStrFreeTag, strgen_empty());
                
            // Пропустить обработку?
            if ((m_buffFreeTag.GetLength() <= 1) && !strcontent(buffStrFreeTag))
                continue;
				
			//	1.	Преобразуем символы типа 10,234,300.50  -> 10234300,50
			//		Признак такого преобразования 
			//		- существует "."
			//		- существует ","
			//		- все осальные символы являются числами
			//	2.	Преобразует "балансовские счета"  типа "03-47" -> " 03-47". Это
			//	связано с тем, что Excel подобную запись преобразует в даты
            
            //if (DBG_MODE) logAddLine("Convert <%s>", buffStrFreeTag);
			buffStrFreeTag = convertNumber(buffStrFreeTag);
            //if (DBG_MODE) logAddLine("\tnumber [%s]", buffStrFreeTag);
			buffStrFreeTag = convertDate(buffStrFreeTag);
            //if (DBG_MODE) logAddLine("\tdate [%s]", buffStrFreeTag);
			
			//	По необходимости перекодируем в windows кодировку
			if (strequal(pcfg->strEncoding, "dos") || strequal(pcfg->strEncoding, "ibm866") ||
				strequal(pcfg->strEncoding, "866"))
                buffStrFreeTag = cp866_to_utf8(buffStrFreeTag, TRUE);
            //if (DBG_MODE) logAddLine("\tto utf8 [%s]", buffStrFreeTag);

            buffStr = m_buff.GetIndexValue(n); 
			//	Закрываем все открытые стандартные теги
			if (n == boundBuff)
			{
				buffStr = strconcatenate(buffStr, ItalicLP[0].SeqEnd, TRUE);
				buffStr = strconcatenate(buffStr, UnderlineLP[0].SeqEnd, TRUE);
				buffStr = strconcatenate(buffStr, BoldLP[0].SeqEnd, TRUE);
				buffStr = strconcatenate(buffStr, BoldLP[1].SeqEnd, TRUE);

				//	Закрываем все открытые внешние (описанные в .cfg) теги
				int nt = pcfg->mTagFormat.GetCount();

				if (nt >= 0)
					buffStr = strconcatenate(buffStr, ((icTagTextFormat *)pcfg->mTagFormat.GetIndexItem(0))->sEnd, TRUE);
			}
            
			//	Ищем специальные тэги форматирования текста
			buffStr = VerSpcTag(buffStr, NameCell, &mfCell, bTagOpen, pcfg);
				
            int row = lxw_name_to_row(NameCell);
            int col = lxw_name_to_col(NameCell);
            
            lxw_format *format = NULL;
            
            //	Предварительно поставим формат данных столбцов, 
            //	который определен в первой строке таблицы после шапки. 
            //	Т.к. общий формат может изменять данные не нужным нам
            //	образом. Пример 0123 -> 123.
            if (row == nLineAfterHead)
                // ВНИМАНИЕ! Здесь клонируется формат ячейки без обрамления
                // чтобы не было разлиновки табличной части 
                tab_formats[col] = clone_cell_data_format(oBook, oSheet, nLineAfterHead, col);
                
            // ВНИМАНИЕ! формат колонки применять только для ячеек после
            // шапки и до того пока не кончиться табличная часть
            if ((row > nLineAfterHead) && (!bEndTable))
                format = tab_formats[col];
            
            //if (DBG_MODE) logAddLine("convertList. Set cell [%d : %d] value <%s> Head: (%d : %d)", row, col, buffStrFreeTag, pcfg->bHead, bTabHead);
            if ((!pcfg->bHead) && (!bTabHead))
            {
                if (DBG_MODE) logAddLine("Set cell (%s) [%d : %d] value <%s>", NameCell, row, col, buffStrFreeTag);
                //	В режиме с шаблоном шапку не выводим				
                set_cell_value(oBook, oSheet, row, col, buffStrFreeTag, format);
            }
            else if (pcfg->bHead)
            {
                if (DBG_MODE) logAddLine("Set head cell %s [%d : %d] value <%s>", NameCell, row, col, buffStrFreeTag);
                set_cell_value(oBook, oSheet, row, col, buffStrFreeTag, format);
            }
            NameCell = strfree(NameCell);
		}
	}

	///////////////////////////////////////////////////////////////////
	//  Конец цикла
	///////////////////////////////////////////////////////////////////
        
	char *cell1 = NULL;
    char *cell2 = NULL;
    char *cellSpc = NULL;
    char *cellSpcNext = NULL;
	
	//////////////////////////////////////////////////////////////
	//  Выделяем и форматируем столбцы
	//////////////////////////////////////////////////////////////

	if (indexFormat.GetLength() == 1 && countLine > 2)
    {
        char *v = itoa(countLine-1, 10);
		indexFormat.AppendValue(v);
        v = strfree(v);
    }

	//	Ищем тэги скрытых строк (в конфиг. спец. слово COLHIDE: "tag") в первой строке после шапки
	//	Найденные тэги удаляются
	cell1 = strinit(cell1, substr(strCell, step * 2, 2));
	cell1 = strtrim(cell1, TRUE);
	cellSpc = strprintf(cellSpc, "%s%d", cell1, nLineAfterHead + 1);

	//	Бывают и такие ситуации, когда стоят разделители страниц где попало

	for (n=0; (n <= maxCell) && (indexFormat.GetLength() > 0); n++)
	{
		cell1 = strinit(cell1, substr(strCell, (n + step) * 2, 2));
		cell1 = strtrim(cell1, TRUE);
		
		//////////////////////////////////////////////////////
		//	Определяем адресса специальных ячеек:
		//	cellSpc - где заканчивается шапка
		//	cellSpcNext - где начинается таблица

		if (nLineAfterHead >= 0)
		{
			cellSpc = strprintf(cellSpc, "%s%d", cell1, nLineAfterHead);
			cellSpcNext = strprintf(cellSpcNext, "%s%d", cell1, nLineAfterHead + 1);
		}
		else
        {
            char *v = indexFormat.GetIndexValue(0);
			cellSpc = strprintf(cellSpc, "%s%s", cell1, v);
            cellSpcNext = strinit(cellSpcNext, strcopy(cellSpc));
            v = strfree(v);
        }

		//	Усли используется шаблон, то форматирование начинаем со строки с данными
		if (!pcfg->bHead)
			cell1 = strinit(cell1, strcopy(cellSpcNext));
		else
        {
            char *v = indexFormat.GetIndexValue(0);
			cell1 = replaceprintf(cell1, "%s%s", cell1, v);
            v = strfree(v);
        }

		cell2 = strinit(cell2, substr(strCell, (n + step) * 2, 2));
		cell2 = strtrim(cell2, TRUE);
        char *v = indexFormat.GetLast()->GetValue();
        int last_row = atoi(v) - 1;
        v = strfree(v);
		cell2 = replaceprintf(cell2, "%s%d", cell2, last_row);
		
		////////////////////////////////////////////////////////
		//	Выставляем ширину столбцов
		if (numColl > n && pcfg->bHead)
            setRangeColumnWidth(oSheet, cell1, cell2, (long)(parColl[n] * COLUMN_WIDTH_COEFF));

		//////////////////////////////////////////////////////
		//	Рисуем границы столбца
        if (DBG_MODE) logAddLine("Draw column [%s : %s]", cell1, cell2);
        setRangeBorder(oBook, oSheet, cell1, cell2, 0, LXW_BORDER_THIN, 0, LXW_BORDER_THIN);
        //setRangeBorder(oBook, oSheet, cell2, cell2,  0, 0, LXW_BORDER_THIN, 0);

		///////////////////////////////////////////////////////
		//	Раскрашиваем по шаблону
        icInterior *oInterior = getRangeInterior(oSheet, cellSpc, cellSpc);
        if (oInterior)
        {
            setRangeInterior(oSheet, cellSpc, cellSpc, oInterior->BackgroundColor);
            delete oInterior;
            oInterior = NULL;
        }
		
        oInterior = getRangeInterior(oSheet, cellSpc, cellSpc);
        if (oInterior)
        {
            setRangeInterior(oSheet, cellSpcNext, cell2, oInterior->BackgroundColor);
            delete oInterior;
            oInterior = NULL;
        }        

		//1.	Устанавливаем по шаблону числовой формат столбца
        icNumberFormat *oNumberFormat = getRangeNumberFormat(oSheet, cellSpcNext, cellSpcNext);
        if (oNumberFormat)
        {
            setRangeNumberFormat(oSheet, cellSpcNext, cellSpcNext);
            delete oNumberFormat;
        }
		//2.	Устанавливаем перенос по словам
        icAlignment *oAlignment = getRangeAlignment(oSheet, cellSpcNext, cellSpcNext);
        if (oAlignment)
        {
            setRangeAlignment(oSheet, cellSpcNext, cellSpcNext, 
                              oAlignment->isVerticalAlignment, oAlignment->isHorizontalAlignment);
            delete oAlignment;
            oAlignment = NULL;
        }
    
        icFont *oFont = getRangeFont(oSheet, cellSpcNext, cellSpcNext);
        if (oFont)
        {
            //3.	Устанавливаем цвет и вид заливки столбца		
            //4.	Устанавливаем по шаблону шрифт столбца		
            //	Копируем атрибуты
            setRangeFont(oBook, oSheet, cellSpcNext, cellSpcNext, 
                         oFont->FontName, oFont->FontSize, oFont->TextColor,
                         oFont->isBold, oFont->isItalic, oFont->isUnderline);
            delete oFont;
            oFont = NULL;
        }
	}

	////////////////////////////////////////////////
	// Форматируем ячейки - формируем таблицу
	////////////////////////////////////////////////
    if (DBG_MODE) logAddLine("--- Step --- [%d]", step);    
    _drawTabHead(oBook, oSheet, &indexFormat, strCell, step, maxCell, pcfg->bHead);

	////////////////////////////////////////////
	// Форматируем текст 

	int bound = mfCell.GetCount();
	n = 0;
    BOOL boolVar;

	for (n; n < bound; n++)
		if (!strempty(((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT) && 
            !strempty(((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB))
		{
			char *s1 = ((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT;
			char *s2 = ((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB;

			// Устанавливаем стили
			//	Устанавливаем цвет шрифта
			//	Устанавливаем название шрифта
			//	Устанавливаем размер шрифта
            setRangeFont(oBook, oSheet, 
                         ((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT, 
                         ((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB, 
                         ((icArrayFormat*) mfCell.GetIndexItem(n))->nameFont, 
                         (long) ((icArrayFormat*) mfCell.GetIndexItem(n))->sizeFont, 
                         (long) ((icArrayFormat*) mfCell.GetIndexItem(n))->clrText, 
                         ((icArrayFormat*) mfCell.GetIndexItem(n))->bBold, 
                         ((icArrayFormat*) mfCell.GetIndexItem(n))->bItalic,
                         ((icArrayFormat*) mfCell.GetIndexItem(n))->bUnderline);

			//	При необходимости скрываем строку
			if (((icArrayFormat*) mfCell.GetIndexItem(n))->bRowHide)
                setRangeRowHidden(oSheet, ((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT, 
                                          ((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB, 
                                          covTrue);

			//	При необходимости скрываем колонку
			if (((icArrayFormat*) mfCell.GetIndexItem(n))->bColHide)
                setRangeColumnHidden(oSheet, ((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT, 
                                             ((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB, 
                                             covTrue);

			//	При необходимости раскрываем колонку
			if (((icArrayFormat*) mfCell.GetIndexItem(n))->bColShow)
                setRangeColumnHidden(oSheet, ((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT, 
                                             ((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB, 
                                             covFalse);

			//	Устанавливаем цвет фона			
			if (((icArrayFormat*) mfCell.GetIndexItem(n))->clrBgr != -1)
                setRangeInterior(oSheet, ((icArrayFormat*) mfCell.GetIndexItem(n))->cellLT, 
                                          ((icArrayFormat*) mfCell.GetIndexItem(n))->cellRB, 
                                          (long) ((icArrayFormat*) mfCell.GetIndexItem(n))->clrBgr);

		}

    // Чистим память
    for (int i=0; i < mfCell.GetCount(); i++)
    {
        if (mfCell.Items[i] != NULL)
        {
            //ВНИМАНИЕ!
            //Здесь необходимо явное приведение к типу класса 
            //для вызова деструтора
            delete ((icArrayFormat *)mfCell.Items[i]);
        }
    }
    mfCell.RemoveAll();          
    
    strCell = strfree(strCell);
    strLine = strfree(strLine);
    strLineFreeTag = strfree(strLineFreeTag);
    cell_TopLeft = strfree(cell_TopLeft);
    cell_BottomRight = strfree(cell_BottomRight);

    buffStr =strfree(buffStr);
    buffStrFreeTag = strfree(buffStrFreeTag);
    
    cell1 = strfree(cell1);
    cell2 = strfree(cell2);
    cellSpc = strfree(cellSpc);
    cellSpcNext = strfree(cellSpcNext);
    
    indexFormat.RemoveAll();
    
    //free(parColl);
    memfree(parColl);
    parColl = NULL;
    
	return TRUE;
}

/**
*	Функция преобразования из Balans-ого формата чисел в Excel-ный:
*	1. числа преобразуются как xxx,xxx,xxx.xx -> xxxxxxxxx.xx
*	2. даты остаются без изменения
*/
char *convertNumber(char *str)
{
	int n = 0;
    int count = 0;
    int countMinus = 0;
    char *bch = NULL;
    char *pstr = strtrim(str, FALSE); 
    int size = strlen(pstr);

    bch = (char *) calloc(size + 1, sizeof(char));

	//	Признак балансовского счета	
	BOOL bBalansCount = FALSE;
    BOOL bPoint = FALSE;

	//	Проверим является ли строка датой по заранее заданному шаблону
	//	- шаблон задается templDate[] = "ii.ii.iiii"
	char templDate[] = "ii.ii.iiii";
	BOOL bDate = FALSE;

    // Проверка на дату
	for (n=0; (n < 10) && (size == 10); n++)
	{
		bDate = TRUE;
		
		//  Проверяем на числа и символы  "." (2E),  ","(2C) "-"
		//	"-" рассматриваю для оптимизации разбора коротких строк,
		//  чтобы потом второй раз не разбирать строку
		if ((pstr[n] < 0x30 || pstr[n] > 0x39)  && pstr[n] != 0x2E && pstr[n] != 0x2C && pstr[n] != '-')
		{
			bch = strfree(bch);
            pstr = strfree(pstr);
			return str;
		}

		//	Признак того, что строка не является датой
		if ((templDate[n] == '.' && pstr[n] != '.') || pstr[n] == '-')
		{
			bDate = FALSE;
			break;
		}

		bch[n] = pstr[n];
	}

    // Проверяем, если первый символ 0 (учет ведущего нуля) значит это строка
    if ((!bDate) && is_leading0(pstr))
    {
        //Это строка
        if (DBG_MODE) logAddLine("<%s> is string", pstr);    
        bch = strfree(bch);
        pstr = strfree(pstr);
        return str;
    }
        
	//Проверка на число
	for (n=0; (n < size) && (!bDate); n++)
	{
       
		//  Проверяем на числа и символы  "." (2E),  ","(2C)
 		if ((pstr[n] < 0x30 || pstr[n] > 0x39)  && pstr[n] != 0x2E && pstr[n] != 0x2C && pstr[n] != '-')
		{
            //Это строка
			bch = strfree(bch);
            pstr = strfree(pstr);
			return str;
		}

        //	Устанавливаем признак точек и запятых
        if (pstr[n] == 0x2E || pstr[n] == 0x2C)
            bPoint = TRUE;

        if (pstr[n] == 0x2E)
            bch[count++] = 0x2E;
        else 
        {
            if (pstr[n] != 0x2C)
                bch[count++] = pstr[n];
				
            // считаем количество '-'
            if (pstr[n] == '-')
                countMinus++;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////
            //	Ищем ситуации, когда счет может преобразоваться в дату
				
            //	1. XX-XX
            if (pstr[n] == '-' && n > 1 && n < size-2 && size == 5)
                bBalansCount = TRUE;		
				
            //	2. X-XX
            if (pstr[n] == '-' && n > 0 && n < size-2 && size == 4)
                bBalansCount = TRUE;		
				
            //	3. XX-X
            if (pstr[n] == '-' && n > 1 && n < size-1 && size == 4)
                bBalansCount = TRUE;		

            //	4. X-XXXX 
            if (pstr[n] == '-' && n > 0 && n < size-4 && size == 6)
                bBalansCount = TRUE;		

            //	5. XX-XXXX
            if (pstr[n] == '-' && n > 1 && n < size-4 && size == 7)
                bBalansCount = TRUE;		
        }
	}
	
    // Заменяем "-----" на 0
    // ВНИМАНИЕ! Бывают случаи когда нам небходимо поставить прочерк в ячейку.
    // В этом случае используем 1 минус. Например в налоговых декларациях.    
	if (countMinus > 1 && countMinus == size)
    {
        str = strfree(str);
		str = strgen('0', 1);
    }
	else
	{
		//
		if (bDate)
			count = 10;

		bch[count] = 0;            // символ конца строки
        //printf("Str: %s Bch: %s", str, bch);
        if (str)
            strcpy(str, bch);
        else
        {
            str = strcopy(bch);
            bch = NULL;
        }

		//	К балансовскому счету добавляем в начале "|", чтобы не преобразовывались к датам
	}

	bch = strfree(bch);
	pstr = strfree(pstr);
	return str;
}


/**
*	Преобразует дату в строку - добавляет впереди строки пробел
*	Признаком даты являются наличие цыфр и символов ".", "/", "\"
*/
char *convertDate(char *str)
{
	int n=0;
    int count=0;
    int countMinus=0;
    char *pstr = strtrim(str, FALSE); 
    int size = strlen(pstr);
	BOOL bDate = FALSE;

	for(n; n < size; n++)
	{
		//  Проверяем на числа и символы  "." (2E),  "/"
 		if ((pstr[n] < 0x30 || pstr[n] > 0x39)  && pstr[n] != 0x2E && pstr[n] != '/' && pstr[n] != '-')
		{
            pstr = strfree(pstr);
			return str;
		}

		if(pstr[n] == 0x2E || pstr[n] == '/' || pstr[n] == '-')
			bDate = TRUE;
	}

    pstr = strfree(pstr);
	return str;
}

/**
*   Произвести необходимые замены в буфере файла
*/
static char *_replacements(char * sBuff, icStringCollection *Word, icStringCollection *Replace)
{
    //	Производим необходимые замены (описаны в *.cfg, REPLACE) в строчке
    for (int nrpl = 0; nrpl < Word->GetLength(); nrpl++)
    {
        char *v1 = Word->GetIndexValue(nrpl);
        char *v2 = Replace->GetIndexValue(nrpl);
        sBuff = strreplace(sBuff, v1, v2, TRUE);
        v1 = strfree(v1);
        v2 = strfree(v2);
    }
    
    //	Заменяем последовательности специальных символов
    //	на более приемлемые
    int n = 0;
    while (ChangeTag[n].Seq != NULL)
    {
        sBuff = strreplace(sBuff, ChangeTag[n].Seq, ChangeTag[n].SeqEnd, TRUE);
        n++;
    }
    
    return sBuff;
}


/**
*	Функция конвертации текстового отчета
*/
BOOL convertTXT(icStructCFG *pcfg)
{
	char *ListValue = NULL;
	BOOL bClear = pcfg->bClear;

	// получаем доступ к Excel
    lxw_workbook *oBook = NULL;
    BOOL covTrue = TRUE;
    BOOL covFalse = FALSE;
    
	//	Считываем текстовый файл
	char *buff = load_txt_file(pcfg->strSource);
	if (!buff)
	{
		errBox("ERROR: Нет возможности открыть файл: <%s>", pcfg->strSource);
		return FALSE;
	}

	pcfg->strEncoding = strlwr_lat(pcfg->strEncoding);

    char *strN = NULL;
    char *s1 = NULL;
    char *s2 = NULL;

    //	Проверяем существует ли такой файл
    if (file_exists(pcfg->strExcelName))
        oBook = open_workbook_xls(pcfg->strExcelName);
    else 
    {
        // Пытаемся создать файл с указаным именем
        char *ExcelDir = NULL;
        char * ExcelName = strcopy(pcfg->strExcelName);
        ExcelName = strreverse(ExcelName);
        int index = strfind(ExcelName, "/");
        ExcelName = strreverse(ExcelName);
        ExcelDir = strleft(ExcelName, strlen(ExcelName) - index - 1);
        if (DBG_MODE) logAddLine("New workbook <%s>", ExcelName);
        if (dir_exists(ExcelDir))
            oBook = new_workbook(ExcelName);
        else
        {
            errBox("Неверно указана директория : <%s>", ExcelName);
            ExcelName = strfree(ExcelName);
            ExcelName = strfree(ExcelDir);
            return FALSE;
        }
        ExcelName = strfree(ExcelDir);
    }
		
    //	Производим необходимые замены (описаны в *.cfg, REPLACE) в строчке
    buff = _replacements(buff, &(pcfg->strWord), &(pcfg->strReplace));
		
    //    Делим текст на страницы. Каждая страница конвертируется на свой лист.
    //  0x1B и ~
    icStringCollection m_buff = icStringCollection();
    icStringCollection mDiv = icStringCollection();
    char *strError = NULL;
    char *ListName = NULL;
    char bch[20];

    if (pcfg->m_Div.GetLength() == 0)
    {
        pcfg->m_Div.AppendValue("  ");
        pcfg->m_Div.AppendValue("|");
    }
		
    //	Собираем строчку разделителей
    int n = 0;
    while (DivPage[n].Seq != NULL)
    {
        mDiv.AppendValue(DivPage[n].Seq);
        n++;
    }
    
    DividePatternAlgbr(buff, &m_buff, &mDiv);
    int bound = m_buff.GetLength();
    char *pageName = NULL;

    //	Цикл по листам
    for (n=0; n < bound; n++)
    {
        if (pcfg->m_PageName.GetLength() > n)
        {
            char *v = pcfg->m_PageName.GetIndexValue(n);
            pageName = strprintf(pageName, "#%s", v);
            v = strfree(v);
        }
        else
            pageName = strprintf(pageName, "#%d", n);
		
        //	Цикл по листам, указанным в тэге LISTNAME
        for(int nList = 0; nList < pcfg->m_ListName.GetLength(); nList++)
        {
            ListValue = strinit(ListValue, pcfg->m_ListName.GetIndexValue(nList));

            if (pcfg->m_ListName.GetLength() > 1)
                ListName = strprintf(ListName, "%s%s", ListValue, pageName);
            else
                if (pcfg->m_PageName.GetLength() > 0)
                    ListName = strinit(ListName, substr(pageName, 0, 1));
                else
                    if(n>0)
                        ListName = strprintf(ListName, "%s%s", ListValue, pageName);
                    else
                        ListName = strinit(ListName, strcopy(ListValue));

            if (DBG_MODE) logAddLine("Worksheet name <%s>", ListName);
            buff = strinit(buff, m_buff.GetIndexValue(n));
            //		Ошибка в разборе последней строчки, если она не заканчивается \r\n
            //	Лень разбираться, поэтому в конце всегда добавляем \r\n - это ничего
            //	не меняет в отчете и уберегет от ошибки разбора
            buff = strconcatenate(buff, "\r\n", TRUE);
    
            pcfg->m_ChangedListName.AppendValue(ListName);
            convertList(buff, ListName, oBook, pcfg, bClear);
        }
    }

    // Отрабатывает скрипт прописанный в атрибуте <POST_SCRIPT> файла конфигурации
    DoPostConvertScript(pcfg->strPostScript, pcfg->strExcelName,
                        pcfg->strCfgName, (int) pcfg->bDelNotChangedList,
                        &pcfg->m_ChangedListName, oBook);

    // Сохраняемся 
    close_workbook_xlsx(oBook);

    //Чистим память
    buff = strfree(buff);
    pageName = strfree(pageName);
    
    ListValue = strfree(ListValue); 
    ListName = strfree(ListName); 
    
    return TRUE;
}


static char DEFAULT_CFG[] = "/usr/share/nixconvert/default.cfg";

BOOL run(int argc, char *argv[])
{  
	if (argc <= 1)
	{
        printHelp();        
		return FALSE;
	}
    
    
    int opt;

    char *strFile;
    char *buff;
	icStructCFG cfg;
	
	long theTimeBeg = get_now_time();

	//////////////////////////////////////////////////////////////////////////////////////////////
    //  Разбор коммандной строки
    char *dbf_filename = NULL;
    char *scr_filename = NULL;
    char *cfg_filename = NULL;
    char *rep_filename = NULL;
    char *tmpl_filename = NULL;
    char *sheet_name = NULL;
    BOOL bClear = FALSE;
    BOOL bSrvMode = FALSE;    
    
    const struct option long_opts[] = {
              { "debug", no_argument, NULL, 'd' },
              { "log", no_argument, NULL, 'j' },
              { "version", no_argument, NULL, 'v' },
              { "help", no_argument, NULL, 'h' },
              { "dbf", required_argument, NULL, 'D' },
              { "script", required_argument, NULL, 'S' },
              { "cfg", required_argument, NULL, 'C' },
              { "rep", required_argument, NULL, 'R' },
              { "template", required_argument, NULL, 'T' },
              { "out", required_argument, NULL, 'O' },
              { "sheet", required_argument, NULL, 'L' },
              { "clear", no_argument, NULL, 'F' },
              { "srv", no_argument, NULL, 'M' },
              { NULL, 0, NULL, 0 }
       };
  
    if (DBG_MODE) logAddLine("OPTIONS:");
    while ((opt = getopt_long(argc, argv, "djvhDSCRTLFM:",long_opts, NULL)) != -1)
    {
        switch (opt) 
        {
            case 'd':
                DBG_MODE = TRUE;
                if (DBG_MODE) logAddLine("\t--debug");
                break;

            case 'j':
                DBG_MODE = TRUE;
                if (DBG_MODE) logAddLine("\t--log");
                break;
                
            case 'h':
                printHelp();
                if (DBG_MODE) logAddLine("\t--help");
                break;
                
            case 'v':
                printVersion();
                if (DBG_MODE) logAddLine("\t--version");
                break;
                
            case '?':
                printHelp();
                return TRUE;

            case 'M':
                bSrvMode = TRUE;
                if (DBG_MODE) logAddLine("\t--srv");
                break;
                
            case 'F':
                bClear = TRUE;
                if (DBG_MODE) logAddLine("\t--clear");
                break;
                
            case 'D':
                dbf_filename = optarg;
                
                if (!file_exists(dbf_filename))
                {
                    logAddLine("File <%s> not exists!", dbf_filename);
                    return FALSE;
                }
                
                if (DBG_MODE) logAddLine("\t--dbf = %s", dbf_filename);
                break;
                
            case 'S':
                scr_filename = optarg;
                
                if (!file_exists(scr_filename))
                {
                    logAddLine("File <%s> not exists!", scr_filename);
                    return FALSE;
                }
                
                if (DBG_MODE) logAddLine("\t--script = %s", scr_filename);
                break;

            case 'C':
                cfg_filename = optarg;
                
                if (!file_exists(cfg_filename))
                {
                    logAddLine("File <%s> not exists!", cfg_filename);
                    return FALSE;
                }
                
                if (DBG_MODE) logAddLine("\t--cfg = %s", cfg_filename);
                break;
                                
            case 'R':
                rep_filename = optarg;
                
                if (!file_exists(rep_filename))
                {
                    logAddLine("File <%s> not exists!", rep_filename);
                    return FALSE;
                }
                
                if (DBG_MODE) logAddLine("\t--rep = %s", rep_filename);
                break;
                
            case 'T':
                tmpl_filename = optarg;
                
                if (!file_exists(tmpl_filename))
                {
                    logAddLine("File <%s> not exists!", tmpl_filename);
                    return FALSE;
                }
                
                if (DBG_MODE) logAddLine("\t--tmpl = %s", tmpl_filename);
                break;
                
            case 'O':
                tmpl_filename = optarg;
                
                if (file_exists(tmpl_filename))
                    del_file(tmpl_filename);
                
                if (DBG_MODE) logAddLine("\t--out = %s", tmpl_filename);
                break;
                
            case 'L':
                sheet_name = optarg;
                
                if (DBG_MODE) logAddLine("\t--sheet = %s", sheet_name);
                break;
                
            default:
                fprintf(stderr, "Unknown parameter: \'%c\'", opt);
                return FALSE;
        }
    }

    // Стереть сообщение об ошибке в начале работы
    strcpy(errorDescription, "");
    
	// Режим работы конвертора:
	if (bSrvMode)
	{
		strFile = argv[0];
		SetOutPutMode(0);
	}
	else
		SetOutPutMode(1);
		
	//	В зависимости от тапа файла указанного в первом параметре
	//	командной строчки выбираются разные механизмы конвертации:
	//	1. Если dbf, то запускаем ф-ию convert()
	//	2. Если текстовый файл, то запускаеь convertTXT
	//     - командная строчка при этом разбирается следующим образом:
	//     *.txt  *.xls List1 bClear
	//	3. Если cfg, то сначало анализируем файл конфигураций
		
    char *CfgName = NULL;
    char *CfgPath = NULL;
	int n_ptrn = 0;
	BOOL bTXT = TRUE;
    BOOL bAnalize = FALSE;

	//	Если есть файл конфигураций
    if (!cfg.AnalalizeCfgFile(DEFAULT_CFG))
    {
        logErr("Default config file <%s> not exists!", DEFAULT_CFG);
        return FALSE;
    }        

	if (cfg_filename)
	{
		//	Разбираем файл конфигураций
		CfgName = cfg_filename;
		CfgPath = getFilePath(CfgName);
		n_ptrn = 1;

		// Необходимо для того, чтобы вычислить имя шаблона, если указан только
		// путь до шаблона
		if (tmpl_filename)
			cfg.strExcelName = strinit(cfg.strExcelName, strcopy(tmpl_filename));

		bAnalize = cfg.AnalalizeCfgFile(CfgName);

		//	Заполняем источник, даже если он описан в файле конфигурации.
		//	Больший приоритет у название, которое в параметрах запуска.	
		if (rep_filename)
			cfg.strSource = strinit(cfg.strSource, strcopy(rep_filename));
		
		if (tmpl_filename)
			cfg.strExcelName = strinit(cfg.strExcelName, strcopy(tmpl_filename));
		
		if (bAnalize)
		{
			//	Определяем является буфферный файл dbf или текстом
			int n=0;
            int bound = cfg.m_DBF.GetLength();
		
			for (n; n < bound; n++)
			{
				//   Ищем по расширению; файл является dbf, если его расширение
				//  есть в списке dbf расширений в файле настроек
                char *v = cfg.m_DBF.GetIndexValue(n);
				if (find_word(cfg.strSource, v))
				{
					bTXT = FALSE;
                    v = strfree(v);
					break;
				}
                v = strfree(v);
			}
		}
	}
	else 
		cfg.strSource = strinit(cfg.strSource, strcopy(rep_filename));
	
	//	Если файла настроек нет или не смогли его проанализировать 
	if (!bAnalize)
	{
		if (dbf_filename)
			bTXT = FALSE;
	}

	//	Разбираем параметры для конвертации отчетов  
	if (bTXT)
	{
		char *XLSName = NULL;
        char *ListName = NULL;
		BOOL bClear = TRUE;
		
		// Проверяем наличие имени xls таблицы
		if (tmpl_filename)
		{
			cfg.strExcelName = strinit(cfg.strExcelName, strcopy(tmpl_filename));
				
			// Проверяем наличие имени листа
			if (sheet_name)
			{
				cfg.m_ListName.RemoveAll();
				cfg.m_ListName.AppendValue(sheet_name);
					
                //	Устанавливаем флаг отчистки содержимого.
				//  - если данный параметр = "f", "false", то флаг устанавливается 
				//  в FALSE, если параметр отсутствует или какая либо другая
				//  строка, то флаг устанавливается в TRUE
				cfg.bClear = bClear;
			}
			// по умолчанию
			else if (cfg.m_ListName.GetLength() == 0)
				cfg.m_ListName.AppendValue("Лист1");
		}
		else if (strempty(cfg.strExcelName))
		{
			errBox("В командной строке не указано имя Excel файла!");
			return FALSE;
		}

        //////////////////////////////////////////////////////////////////
		char *ExcelDir = getFilePath(cfg.strExcelName);
		
		//	Если задан относительный путь, то определяем полный путь
		if (strempty(ExcelDir) && !strempty(CfgPath))
			cfg.strExcelName = strprintf(cfg.strExcelName, "%s/%s", CfgPath, cfg.strExcelName);
        ExcelDir = strinit(ExcelDir, getFilePath(cfg.strExcelName));

		//	Проверяем существует ли директория для файла отчета
		if (!dir_exists(ExcelDir))
		{
			// !!!!! Пока убрал, т.к. в серверном варианте работать не будет
			int result = TRUE;
			BOOL bCreate = FALSE;
				
			//	Пытаемся создать каталог		
			bCreate = mkpath(ExcelDir, S_IRWXU | S_IRGRP | S_IXGRP | S_IROTH | S_IXOTH);

			//	Если каталог создать не удается, завершаем работу			
			if (!bCreate)
				return FALSE;
		}

        
		//	Договорились, что если флаг шапки = FALSE,
		//	и имя шаблона не задано, то в качестве шаблона
		//	используется основной Excel файл.
		if (strempty(cfg.strTemplate) && !cfg.bHead)
			cfg.strTemplate = cfg.strExcelName;
		else 
            if (strequal(cfg.strTemplate, "CFG") && !cfg.bHead)
                cfg.strTemplate = dupprintf("%s/%s", CfgPath, getFileName(cfg.strExcelName));

		//	Если задан относительный путь, то определяем полный путь
        ExcelDir = strinit(ExcelDir, getFilePath(cfg.strExcelName));
		if (strempty(ExcelDir) && !strempty(CfgPath))
			cfg.strTemplate = dupprintf("%s/%s", CfgPath, cfg.strTemplate);

		//	Проверяем описан ли шаблон
		if (cfg.bHead == FALSE)
		{
			//	Если возможно копируем из шаблон
			if (file_exists(cfg.strTemplate))
			{
				//	Если не удается копировать файл шаблона, то пытаемя закрыть Excel
				//	и скопировать заново
				copy_file(cfg.strTemplate, cfg.strExcelName);
			}
			else
			{
				ExcelDir = getFilePath(cfg.strTemplate);

				//	Проверяем существует ли директория для файла шаблона
				if (!dir_exists(ExcelDir))
				{
					// !!!!! Пока убрал, т.к. в серверном варианте работать не будет
					int result = TRUE;
					BOOL bCreate = FALSE;
						
					//	Пытаемся создать каталог
					bCreate = mkpath(ExcelDir, S_IRWXU | S_IRGRP | S_IXGRP | S_IROTH | S_IXOTH);

					//	Если каталог создали, заменяем имя файла отчета 
					//	на имя шаблона. После первой загрузки будет создан шаблон,
					//	в следующий раз шаблон будет копироваться в файл отчета, 
					//  таким образом будут сохраняться настройки форматирования
					//	шапки отчета.
					if (bCreate)
					{
						cfg.strExcelName = cfg.strTemplate;
						cfg.bHead = TRUE;
					}
				}
				else
				{
					cfg.strExcelName = cfg.strTemplate;
					cfg.bHead = TRUE;
				}
			}
		}

		//////////////////////////////////////////////////////////////////
		if(!convertTXT(&cfg))
			return FALSE;

        
        ExcelDir = strfree(ExcelDir);
        XLSName = strfree(XLSName);
        ListName = strfree(ListName);
		logAddLine("TXT convert. EXCELNAME: <%s>", cfg.strExcelName);
	}
	// Конвертация через буферный файл
	else 
	{	
		char *scriptFile = NULL;

		//	Усли указан третий параметр, то трактуем его как имя файла
		//	скрипта
		if (scr_filename)
			scriptFile = scr_filename;
		
        if (DBG_MODE) cfg.PrintCFG();
		convert(&cfg, dbf_filename, scriptFile);
	}
	
    
	if (strempty(errorDescription))
    {
        long theTimeEnd = get_now_time();
        char *str_time = long_to_time(theTimeEnd - theTimeBeg);
        char *sTime = dupprintf("Время конвертации: <%s>", str_time);
        str_time = strfree(str_time);
        
		logAddLine("REPORT: Обновление данных завершено. %s", sTime);
        sTime = strfree(sTime);
    }
    else
		errBox("Ошибка выполнения: <%s>", errorDescription);
    
    //Очитска памяти
    CfgPath = strfree(CfgPath);
    
   	return FALSE;
}