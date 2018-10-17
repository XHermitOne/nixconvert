// Содержит необходимые структуры и функции для разбора текстового 
// файла настроек	
//////////////////////////////////////////////////////////////////////

#if !defined( __TEXTPARSER_H )
#define __TEXTPARSER_H

class icStructCFG;

BOOL SetWidthCol(char *strLine, int **parColl, int &numColl);
BOOL DivideFixPattern(char *str, icStringCollection *m_buff, int *par, int numPar);
BOOL DivideFixTagPattern(char *str, icStringCollection *m_buff, int *par, int numPar, icStructCFG *pcfg);

BOOL CreateArchive(char *pFileName);

BOOL DividePattern(char *str, icStringCollection *m_buff, char *sdiv);
BOOL DividePatternAlgbr(char *str, icStringCollection *m_buff, icStringCollection *mDiv);

//	File & Dir
BOOL LoadTextFile(char *Name, char *ret);
BOOL isFile(char *fileName);
BOOL isDir(char *dirName);
char *getFilePath(char *fileName);
char *getFileName(char *fileName);
char *MakeFileName(char *fileName, char *imageFileName);
char *replacePath(char *Path, icStringCollection *strWord, icStringCollection *strReplace, int iCount);

#endif /*__TEXTPARSER_H*/
