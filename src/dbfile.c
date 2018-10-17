// dbFile.cpp: implementation of the dbFile class.
//
//////////////////////////////////////////////////////////////////////

#include "econvert.h"

/**
*   Construction/Destruction
*/
dbFile::dbFile()
{
  init();
}

dbFile::dbFile(char *FileName)
{
  init();
  Open(FileName);
}

dbFile::~dbFile()
{
	Close();
}

/**
* Инициализация
*/
void dbFile::init()
{
  opened = FALSE;
  records = NULL;
  file = NULL;
  types.RemoveAll();
  fields.RemoveAll();
  fld = NULL;
  LField = NULL;
  Decimal = NULL;
  cursor = 0;
}

/**
* Открыть dbf
*/
void dbFile::Open(char *FileName)
{
    if (opened) 
        Close();
    opened = TRUE;

    file = fopen(FileName, "rwb");
    if (file == NULL)
    {
        errBox("Ошибка открытия файла: <%s>\n  ", FileName);
        Close();
        return;
    }
    
    //fread((void*)(&fHdr), 1, sizeof(DB_HEADER), file);
    fread(&fHdr, sizeof(DB_HEADER), 1, file);
    if (DBG_MODE) logAddLine("DBF file <%s> header:", FileName);
    //fread((void*)&fHdr.Type, sizeof(BYTE), 1, file);
    if (DBG_MODE) logAddLine("\tType: <%x>", fHdr.Type);
    //fread((void*)&fHdr.Year, sizeof(BYTE), 1, file);
    //fread((void*)&fHdr.Month, sizeof(BYTE), 1, file);
    //fread((void*)&fHdr.Day, sizeof(BYTE), 1, file);
    if (DBG_MODE) logAddLine("\tDate: %d-%d-%d", fHdr.Year, fHdr.Month, fHdr.Day);
    //fread((void*)&fHdr.nRecords, sizeof(DWORD), 1, file);
    if (DBG_MODE) logAddLine("\tRecord count: <%d>", fHdr.nRecords);
    //fread((void*)&fHdr.Offset, sizeof(WORD), 1, file);
    if (DBG_MODE) logAddLine("\tOffset: <%d>", fHdr.Offset);
    //fread((void*)&fHdr.RecordLen, sizeof(WORD), 1, file);
    if (DBG_MODE) logAddLine("\tRecord length: <%d>", fHdr.RecordLen);
    
    BYTE t = fHdr.Type;
    if ((t != 0x03) && (t != 0x83) && (t != 0xf5) && (t != 0x8b))
    {
        Close();
        errBox("File <%s> dosn't have a DBF-structure ! File is closed !", FileName);
        return;
    }
    nFields = (fHdr.Offset - (int)(sizeof(DB_HEADER))) / (int)(sizeof(DB_FIELD));
    
    if (nFields <= 0)
        errBox("Field count error: <%d> Offset: <%d>", nFields, fHdr.Offset);
        
	firstRec = 0;
	lastRec = getRecCount();
    fld = (DB_FIELD *) calloc(nFields, sizeof(DB_FIELD));
	LField = (int *) calloc(nFields, sizeof(int));
	Decimal = (int *) calloc(nFields, sizeof(int));
    int offset = 1;
    for (int i=0; i<nFields; i++)
    {
        fread((void*)&fld[i], 1, sizeof(DB_FIELD), file);
		fld[i].Offset = offset;
		offset += fld[i].FieldLen;
        fields.AppendValue((char *) fld[i].Name);
        //Прочитать тип поля
        char *sql_type = strSQLType(&fld[i]);
        types.AppendValue(sql_type);
        sql_type = strfree(sql_type);
        
		LField[i] = fld[i].FieldLen;
		Decimal[i] = fld[i].Decimals;
    }
}

/**
*   Закрыть базу
*/
void dbFile::Close()
{
    fields.RemoveAll();
    types.RemoveAll();
    
    if (file != NULL) 
    {
        fclose(file);
        file = NULL;
    }
    opened = FALSE;
    if (records != NULL) 
    {
        memfree(records);  
        records = NULL;
    }
    if (fld != NULL) 
    {
        memfree(fld); 
        fld = NULL;
    }
    if (LField != NULL) 
    {
        memfree(LField); 
        LField = NULL;
    }
    if (Decimal != NULL) 
    {
        memfree(Decimal); 
        Decimal = NULL;
    }
}

/**
*   Получить тип данных в формате SQL
*/
char *dbFile::strSQLType(DB_FIELD * fld)
{
    char *fldType;
    switch (fld->Type) 
    {
    case 'C':
        fldType = dupprintf( "varchar(%d)", fld->FieldLen);
        break;
    case 'N':
        fldType = dupprintf( "numeric(%d, %d)", fld->FieldLen, fld->Decimals);
        break;
    case 'L':
        fldType = dupprintf( "bit");
        break;
    case 'D':
        fldType = dupprintf( "date");
        break;
    case 'F':
        fldType = dupprintf( "float");
        break;
    default :
        fldType = dupprintf( "not found");
    }
    
    return fldType;
}

/**
*   Открыт ли файл
*/
BOOL dbFile::isOpen()
{
    return opened;
}

/**
*   Сколько в DBF-е записей
*/
int dbFile::getRecCount()
{
    return fHdr.nRecords;
}

/**
*   Прочитать записи
*/
void dbFile::readRecords(DWORD nBeg = 0, DWORD nEnd = -1)
{
    if (nEnd > fHdr.nRecords || nEnd == -1) 
        nEnd = fHdr.nRecords-1;
    if (nBeg < 0) 
        nBeg = 0;
    firstRec = nBeg; 
    lastRec = nEnd;
    if (records != NULL) 
    {
        memfree(records);
        records = NULL;
    }

    int allBytes = fHdr.RecordLen * (nEnd - nBeg + 1);
    records = (BYTE *) calloc(allBytes, sizeof(BYTE));
    fseek(file, fHdr.Offset + nBeg * fHdr.RecordLen, SEEK_SET);
    fread((void*)records, 1, allBytes, file);
}

/**
*   Сохранить прочитанные записи
*/
void dbFile::writeRecords()
{
    int allBytes = fHdr.RecordLen * (lastRec - firstRec + 1);
    fseek(file, fHdr.Offset + firstRec * fHdr.RecordLen, SEEK_SET);
    fwrite((void*)records, allBytes, sizeof(BYTE), file);
}

/**
*   Прочитать запись, если ее не существует, то в базу добавляется нужное количество записей
*/
void dbFile::readRecord(DWORD n)
{
    if (n >= fHdr.nRecords)
        addRecords(n-fHdr.nRecords+1);
    readRecords(n, n);
}

/**
*   Очистить запись (одну из прочитанных)
*/
void dbFile::clearRecord(DWORD n)
{
    if (n >= firstRec && n <= lastRec)
    {
        memset((void*)&records[(n - firstRec) * fHdr.RecordLen], 0x20, fHdr.RecordLen);

        for (int i=0; i<nFields; i++)
        {
            if (fld[i].Type == 'N')
            {
                int offset = (n - firstRec) * fHdr.RecordLen + fld[i].Offset;
                if (fld[i].Decimals != 0)
                {
                    for (int j=fld[i].FieldLen-fld[i].Decimals-2; j<fld[i].FieldLen; j++) 
                        records[offset+j] = '0';
                    records[offset+fld[i].FieldLen-fld[i].Decimals-1] = '.';
                }
                else
                {
                    records[offset+fld[i].FieldLen-1] = '0';
                }
            }
        }
    }
}

/**
*   Добавляет n записей
*/
void dbFile::addRecords(int n)
{
    int allBytes = fHdr.RecordLen * n;
    BYTE* buffer = (BYTE *) calloc(allBytes, sizeof(BYTE));
    memset((void*)buffer, 0x20, allBytes);

    fseek(file, fHdr.Offset + fHdr.nRecords * fHdr.RecordLen, SEEK_SET);
    fwrite((void*)buffer, allBytes, sizeof(BYTE), file);
    fseek(file, 0L, SEEK_SET);
    
    int from = fHdr.nRecords;
    fHdr.nRecords += n;
    fwrite((void*)&fHdr, 1, sizeof(fHdr), file);
    memfree(buffer);
}

/**
*   Взять значение поля из записи в виде строки
*/
char *dbFile::getRecField(DWORD nRec, int nFld)
{
    readRecords(nRec, nRec);
	if (nRec >= firstRec && nRec <= lastRec)
	{
        int offset = (nRec - firstRec) * fHdr.RecordLen + fld[nFld].Offset;
		BYTE tmpB = records[offset+fld[nFld].FieldLen];
		records[offset+fld[nFld].FieldLen] = 0;
		char *ret = strcopy((char*) &records[offset]);
		ret = strtrim(ret, TRUE);
		records[offset+fld[nFld].FieldLen] = tmpB;
        if (ret == NULL)
            if (DBG_MODE) errBox("DBF. Not correct field value <%d>", nFld);
		return ret;
	}
	//return strcopy("err");
    if (DBG_MODE) errBox("DBF. Not correct record index <%d>", nRec);
    return NULL;
}

/**
*   Установить значение поля в записи
*/
void dbFile::setRecField(DWORD nRec, int nFld, BYTE* data)
{
    if (nRec >= firstRec && nRec <= lastRec)
    {
        int offset = (nRec - firstRec) * fHdr.RecordLen + fld[nFld].Offset;
        memset((void*)&records[offset], 0x20, fld[nFld].FieldLen);
        for (int i=0; i<fld[nFld].FieldLen; i++)
        {
            if (data[i]==0) break;
            records[offset+i]=data[i];
        }
    }
}

/**
*   Сбросить флаг удаления, (т.е. запись становится не удаленной)
*/
void dbFile::setNoDel(DWORD nRec)
{
    if (nRec >= firstRec && nRec <= lastRec)
    {
        records[(nRec - firstRec) * fHdr.RecordLen] = ' ';
    }
}

/**
*   Пометить запись на удаление, поля заполняются по типу клипперовской функции
*/
void dbFile::delRecord(DWORD nRec)
{
    if (nRec >= firstRec && nRec <= lastRec)
    {
        records[(nRec - firstRec) * fHdr.RecordLen] = '*';
        for (int i=0; i<nFields; i++)
        {
            char *fieldname = NULL;
            fieldname = fields.GetIndexValue(i);
           
            if (strequal(fieldname, "ARM_REC"))
            {
                int offset = (nRec - firstRec) * fHdr.RecordLen + fld[i].Offset;
                if (fld[i].Type == 'D')
                    strncpy((char*)&records[offset],"29991231",8);
                else
                    if (fld[i].Type == 'C')
                        for (int j=0; j<fld[i].FieldLen; j++) 
                            records[offset+j] = 223;
                    else
                    {
                        for (int j=0; j<fld[i].FieldLen; j++) records[offset+j] = '9';
                        if (fld[i].Decimals != 0)
                        {
                            records[offset+fld[i].FieldLen-fld[i].Decimals-1] = '.';
                            for (int j=fld[i].FieldLen-fld[i].Decimals; j<fld[i].FieldLen; j++) 
                                records[offset+j] = '0';
                        }
                    }
            }
            fieldname = strfree(fieldname);
        }
    }
}

BOOL dbFile::Next()
{
    BOOL bret = FALSE;

	if (fHdr.nRecords > cursor)
	{
		cursor++;
		bret = TRUE;
	}

	return bret;
}

BOOL dbFile::Prev()
{
	BOOL bret = FALSE;

	if (cursor > 0)
	{
		cursor--;
		bret = TRUE;
	}

	return bret;
}

void dbFile::gotop()
{
	cursor = 0;
}

void dbFile::gobottom()
{
	cursor = fHdr.nRecords;
}

char *dbFile::getField(char *sName)
{
    //	по названию поля определяем его номер 
	int nFld = -1;
    int n = 0;

    // if (DBG_MODE) logAddLine("DBF. Field count <%d>", nFields);
    // Поиск поля
	for(n; n < nFields; n++)
	{
        char *fieldname = NULL;
        fieldname = fields.GetIndexValue(n);
		if (strequal(fieldname, sName))
		{
			nFld = n;
            fieldname = strfree(fieldname);
			break;
		}
        fieldname = strfree(fieldname);
	}

	if (nFld >= 0)
	{
		return getRecField(cursor, nFld);
	}

    if (DBG_MODE) errBox("DBF. Field <%s> not found", sName);
	return NULL;
}

BOOL dbFile::setField(char *sName, char *sValue)
{
	int nf = -1;
    sName = strupr_lat(sName);

	//	Ищем нужное поле
	for (int n=0; n<nFields; n++)
	{
        char *fieldname = NULL;
        fieldname = fields.GetIndexValue(n);
		if (strequal(fieldname, sName))
		{
			nf = n;
            fieldname = strfree(fieldname);
			break;
		}
        fieldname = strfree(fieldname);
	}

	//	Пишем в поле 
	if (nf >= 0)
	{
		if (records == NULL)
			readRecords();

		setRecField(cursor, nf, (BYTE *) sValue);	
		return TRUE;
	}

	return FALSE;
}

BOOL dbFile::IsEOF()
{
	if (cursor >= fHdr.nRecords)
		return TRUE;

	return FALSE;
}

