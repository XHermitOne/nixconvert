// dbFile.h: interface for the dbFile class.
//
//////////////////////////////////////////////////////////////////////

#if !defined( __DBFILE_H )
#define __DBFILE_H

#include "ictypes.h"
#include "glob.h"
#include "dbstruct.h"

class dbFile  
{
public:

	//	Проверка на конец файла
	BOOL IsEOF();

	//	Пищет в поле текущей записи
	BOOL setField(char *sName, char *sValue);

	//	Читает значение поля из текущей записи
	char *getField(char *sName);

	//	Функции навигации
	void gobottom();
	void gotop();
	BOOL Prev();
	BOOL Next();

	//	Указатель на текущую запись
	int cursor;

	void init();

	void readRecords(DWORD nBeg, DWORD nEnd);
	void readRecord(DWORD n);
	void addRecords(int n);
	void clearRecord(DWORD n);
	void writeRecords();
	void setNoDel(DWORD nRec);
	void delRecord(DWORD nRec);
	char *getRecField(DWORD nRec, int nFld);
	void setRecField(DWORD nRec, int nFld, BYTE* data);
    
	int getRecCount();
	BOOL isOpen();
	dbFile();
	dbFile(char *fileName);
	void Open(char *FileName);
	void Close();
	int nFields;
	icStringCollection fields = icStringCollection();
	icStringCollection types = icStringCollection();
	int *LField;				// размеры полей
	int *Decimal;              // кол. разрядов после запятой
	virtual ~dbFile();
    
private:
	DWORD firstRec, lastRec;
	BYTE *records;
	BOOL opened;
	FILE *file;
	DB_HEADER fHdr;
	DB_FIELD *fld;
	char *strSQLType(DB_FIELD *fld);
};

#endif /* __DBFILE_H */
