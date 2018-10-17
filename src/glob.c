#include "econvert.h"

int outputMode = 0;

/**
* Считывает из потока символы до встречи определенного и формирует строку
*/
char *readBefSymb(char *arch, BYTE smb)
{
  BYTE *buffer = new BYTE[255];
  int len = 0;
  BYTE cur;
  
  do
  {
    *(arch) >> cur;
    buffer[len++] = cur;
  }
  while (cur!=smb);
  buffer[len-1] = 0;
  return (char *) buffer;
}

/**
* Считывает из потока символы до встречи '\0' и формирует строку
*/
char *readFromArch(char *arch)
{
  return readBefSymb(arch, 0);
}

/**
* Считывает из потока данные до встречи определенного символа
*/
BYTE *readBefByte(char *arch, BYTE smb)
{
  BYTE *buffer = new BYTE[255];
  int len = 0;
  BYTE cur;
  do
  {
    *(arch) >> cur;
    buffer[len++] = cur;
  }
  while (cur!=smb);
  BYTE* ret = new BYTE[len+1];
  for (int i=0; i<len; i++)
  {
    ret[i]=buffer[i];
  }
  ret[len] = 0;
  delete buffer;
  return ret;
}

/**
* Выбирает из строки значение параметра. Например -xSOMEFILE.TXT
*/
char *readKey(char* buffer, BYTE key)
{
  char *begStr = NULL;
  char endChar;
  
  char *ret = "none";
  int len = strlen(buffer)+1;
  
  for (int i=0; i<len; i++)
  {
    if (buffer[i]=='-' || buffer[i]=='/')
    {
      if (buffer[i+1]==key || buffer[i+1]==key-32)
      {
        if (buffer[i+2] == '\"')
        {
          endChar = '\"';
          begStr = &buffer[i+3];
          i+=2;
        }
        else
        {
          endChar = ' ';
          begStr = &buffer[i+2];
        }
      }
    }
    else
    if (buffer[i]==endChar || buffer[i]=='\0')
    {
      if (begStr != NULL)
      {
        char tmpB = buffer[i];
        buffer[i] = 0;
        ret = begStr;
        buffer[i] = tmpB;
      }
      return ret;
    }
  }
  return ret;
}


/**
*   Функция выводит сообщения в файл лога 
*/
void WritetoLog(char *mess)
{
    logAddLine(mess);
}

/**
*   Функция выводит сообщения в файл лога либо на экран в зависимости от режима 
*   работы
*/
void icMessageBox(char *mess)
{
	
	if (GetOutPutMode() != ONLY_WINDOW_OUTPUT)
	{
		WritetoLog(mess);
	}
}

/**
*   Функция возвращает идентификатор режима работы на вывод сообщений.
*/
int GetOutPutMode()
{
	return outputMode;
}

/**
*   Функция устанавливает идентификатор режима работы на вывод сообщений.
*/
void SetOutPutMode(int mode)
{

	outputMode = mode;
}