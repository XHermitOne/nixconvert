/**
* Модуль функции MAIN
* @file
*/
#include "econvert.h"

/**
* Режим отладки
*/
BOOL DBG_MODE = FALSE;

/**
* Основная функция MAIN. Анализ параметров коммандной строки и 
* установка этих параметров. 
*
* Параметры коммандной строки:
* @param \--help|-h|-? Напечатать строки помощи
* @param \--version|-v  Напечатать версию программы
* @param \--debug|-D Включить режим отладки
* @param \--log|-L Включить режим журналирования
*/

int main (int argc, char **argv)
{
    //if (argc==1 || checkParm("--help",argc,argv) || checkParm("-h",argc,argv) || checkParm("-?",argc,argv))
    if (checkParm("--help",argc,argv) || checkParm("-h",argc,argv) || checkParm("-?",argc,argv))
    {
        printHelp();
        return 0;
    }

    if (checkParm("--version",argc,argv) || checkParm("-v",argc,argv))
    {
        printVersion ();
        return 0;
    }

    if ((checkParm("--debug",argc,argv) || checkParm("-D",argc,argv)) || (checkParm("--log",argc,argv) || checkParm("-L",argc,argv)))
        DBG_MODE = TRUE;
    else
        DBG_MODE = FALSE;

    run(argc, argv);
    return 1;
}

