/**
* Модуль функций записи в лог
* @file
*/


#ifdef __cplusplus
#include <cstdio>			// FILE *
#else
#include <stdio.h>			// FILE *
#endif
//#include <stdio.h>
#include <time.h>
#include <stdarg.h>
#include <stdlib.h>
#include <string.h>

#include "tools.h"
#include "log.h"
#include "strfunc.h"

/**
* Текущее время-дата
*/
char TimeDate[128];
char *getTimeDate()
{
    TimeDate[0]=0;
    time_t loc_time;
    time(&loc_time);

    tm *today = localtime(&loc_time);
    strftime(TimeDate, 128, "%d/%m/%y (%H:%M:%S)", today);

    return TimeDate;
}

/**
* Класс менеджера лога
*/
class LogInit
{
    public:
        FILE *out;
        bool isNew;

        LogInit( char *LogName )
        {
            out = NULL;
            isNew = false;
            char *cfg_path = getCfgPath();
    
            char *full_log_filename = (char*) calloc(strlen(cfg_path) + strlen(LogName)+1, sizeof(char));
            strcpy(full_log_filename, cfg_path);
            strcat(full_log_filename, LogName);
        
            out = fopen(full_log_filename, "a");
            fprintf(out, "[START LOG] %s - - - - - - - - - - - - - - - - - - - - -\n", getTimeDate());

            full_log_filename = strfree(full_log_filename);
            //cfg_path = strfree(cfg_path);
        }

        ~LogInit()
        {
            if (out)
            {
                fprintf(out, "[STOP LOG] %s - - - - - - - - - - - - - - - - - - - - -\n", getTimeDate());
                fclose(out);
                out = NULL;
            }
        }
};

/**
* Лог файл
*/
class LogInit LogFile("/nixconvert.log");

/**
* Добавить строчку в лог
*/
void logAddLine(char *S, ...)
{
    char buffer[MAX_LOG_MSG];
    va_list ap;

    va_start(ap, S);
    vsprintf(buffer, S, ap);
	
    log_color_line(IC_CYAN_COLOR_TEXT, buffer);
}


/**
* Сообщение об ошибке
*/
void logErr(char *S, ...)
{
    char buffer[MAX_LOG_MSG];
    va_list ap;

    va_start(ap, S);
    vsprintf(buffer, S, ap);
	
    log_color_line(IC_RED_COLOR_TEXT, buffer);
}


/**
* Предупреждение
*/
void logWarning(char *S, ...)
{
    char buffer[MAX_LOG_MSG];
    va_list ap;

    va_start(ap, S);
    vsprintf(buffer, S, ap);
	
    log_color_line(IC_YELLOW_COLOR_TEXT, buffer);
}


void log_color_line(unsigned int iColor, char *S, ...)
{
    if (LogFile.out)
    {
        va_list argptr;
        va_start(argptr, S);

        char msg[MAX_LOG_MSG];
        vsnprintf(msg, sizeof(msg), S, argptr);
        va_end(argptr);

        // Сигнатуру сообщения определяем по цвету
        char signature[10];
        if (iColor == IC_CYAN_COLOR_TEXT)
            strcpy(signature, "DEBUG:");
        else if (iColor == IC_RED_COLOR_TEXT)
            strcpy(signature, "ERROR:");
        else if (iColor == IC_RED_COLOR_TEXT)
            strcpy(signature, "WARNING:");
        else
            strcpy(signature, "");
            
        fprintf(LogFile.out, "    %s %s %s\n", getTimeDate(), signature, msg);
        print_color_txt(iColor, "%s %s %s\n", getTimeDate(), signature, msg);
        fflush(LogFile.out);
    }
}
