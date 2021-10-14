// StructCFG.h: interface for the icStructCFG class.
//
//////////////////////////////////////////////////////////////////////

#if !defined( __STRUCTCFG_H )
#define __STRUCTCFG_H

#define MAX_LEN_TAG   256

class icTagTextFormat   
{
public:
    //  Имя тэга
    char *Name;

    //  Открывающий тэг
    char *sBeg;

    //  Закрывающий тэг
    char *sEnd;

    //  Цвет текста
    COLORREF clrText;

    //  Цвет фона
    COLORREF clrBgr;

    /////////////////////
    //  Название шрифта
    char *nameFont;

    //  размер шрифта
    int sizeFont;

    /////////////////////
    //  Флаги стилей

    //  Италик
    BOOL bItalic;

    //  Жирный текст
    BOOL bBold;

    //  Подчеркнутый
    BOOL bUnderline;

    BOOL Init(char *tagname, COLORREF clr1, COLORREF clr2);
    
    icTagTextFormat(); 
    ~icTagTextFormat();
};


class icStructCFG    
{
public:

    //  Флаг, указывающий на необходимость центрирования 
    //  текста перед шапкой
    BOOL bCentredTextBeforeHead;

    //  Флаг, указывающий на наличие шапки у отчета
    BOOL bHead;

    //  Флаг полной отчистки
    BOOL bClear;

    //  Тип кодировки
    char *strEncoding;

    //  Количество символов в TAB
    int iTab;

    //  Название тэга, скрывающего строку
    char *strTagHideRow;

    //  Название тэга, скрывающего колонку
    char *strTagHideCol;
    
    //  Название тэга, открывающего колонку
    char *strTagShowCol;

    //  Полный путь до файла конфигурации
    char *strCfgName;

    //  Название файла отчета или буффера       
    char *strSource;

    //  Имя ехсеl файла, куда конвертируется отчет
    char *strExcelName;

    //  DEPRICATED: Имя листа файла, куда конвертируется отчет
    char *strListName;

    //  Список листов, куда копируется отчет (получается несколько
    //  копий отчета)
    icStringCollection m_ListName = icStringCollection(); 

    //  Список листов, куда раскидывается отчет, определяется по специальному атрибуту
    //  PAGENAME:
    icStringCollection m_PageName = icStringCollection(); 

    //  Имя файла шаблона (excel), по которому формируется шапка    
    char *strTemplate;

    //  Имя листа шаблона (excel), по которому формируется шапка    
    char *strTemplateList;

    //  Массив расширений dbf
    icStringCollection m_DBF = icStringCollection();
    
    //  Строчка разделителей
    icStringCollection m_Div = icStringCollection();

    //  Слова по которым определяются строчки шапки, которые надо обновлять
    //  при работе с шаблоном
    icStringCollection m_sPrizn = icStringCollection();   

    //  Строчки замен
    icStringCollection strWord = icStringCollection();   
    icStringCollection strReplace = icStringCollection();   

    //  Строчки замен путей
    icStringCollection strPathWord = icStringCollection();   
    icStringCollection strPathReplace = icStringCollection();   

    //  Тэги, задающие формат текста
    icArray mTagFormat = icArray();

    //  Тэг <POST_SCRIPT>, задающий названия файла динамических параметров - используется для 
    //  задания команд скрипта, который будет отрабатывать после конвертации
    char *strPostScript;

    //  Список листов, в которые выгружались данные
    icStringCollection m_ChangedListName = icStringCollection();   

    //  Флаг удаления, не измененных листов. Флаг определяется по атрибуту <DEL_NOT_CHANGED_LIST>
    BOOL bDelNotChangedList;

    //  Флаг проверки чисел на ИНН
    BOOL bCheckINN;

    //  Функция разбирает файл конфигураций и на основе его заполняет конфигурационную 
    //  структуру
    BOOL AnalalizeCfgFile(char *CfgName);
    icTagTextFormat *ParseStyle(char *s1);    
    
    void PrintCFG(void);
    
    icStructCFG();
    ~icStructCFG();

private: 
    icTagTextFormat *_newTagTextFormat(char *sTagName, COLORREF Color1, COLORREF Color2, icStringCollection *ParseString);
    
};

#endif /* STRUCTCFG */