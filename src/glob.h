#if !defined( __GLOB_H )
#define __GLOB_H

#include "ictypes.h"

char *str(int i);
char *readFromArch(char *arch);
char *readBefSymb(char *arch, BYTE smb);
BYTE* readBefByte(char *arch, BYTE smb);
char *readKey(char* buffer, BYTE key);

void icMessageBox(char *mess);
void WritetoLog(char *mess);

//	Определяет режим вывода 
#define LOGFILE_OUTPUT  0
#define LOGFILE_WINDOW_OUTPUT 1
#define ONLY_WINDOW_OUTPUT 2

extern int outputMode;
int GetOutPutMode();
void SetOutPutMode(int mode);

#endif /* __GLOB_H */
