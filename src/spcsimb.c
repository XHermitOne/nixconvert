#include "spcsimb.h"

/**
*	Разделители страниц
*/
TagCode DivPage[] =
{
	{"\x0C", ""},    // 12
	{NULL, NULL}
};

/**
*	Заменяемые последовательности
*/
TagCode ChangeTag[] =
{
	{"=", "-"},
	{"@ ", "@-"},
	{"^", " "},
	{"+--", "---"},
	{"--!", "---"},
	{"-{-", "---"},
	{"--{", "---"},
	{"-}-", "---"},
	{"Лист #", ""},
	{"Лист  #", ""},
	// {"\x09","        "},             // 8-й элемент TAB заменяем на 8 пробелов
	//  последовательности шрифтов
	{"\x1B\x53\x30\x0F", ""},
	{"\x1B\x50\x12", ""},                       
	{"\x12\x1B\x54", ""}, 
	{"\x1B\x4D\x0F", ""},
	{"\x1B\x53\x30", ""},
	{"\x1B\x41\x08", ""},
	{"\x1B\x50", ""},
	{"\x1B\x32", ""},
	{"\x1B\x54", ""},
	{"\x1B\x43", ""},
	{"\x1B\x4D", ""},
	{"\x0F", ""},         
	// {"\xA4", ""},   это 'д' 164
	{"\x0E", ""},            // 14           
	{"\x12", ""},             // 18 
	{"\x14", ""},			//20
	{NULL, NULL}
};

/**
*	Удаляемые последовательности
*/
TagCode EmptyLP [] = 
{
	{"\x1B\x50\x12", "\x1B\x4D"},               // пика
	{"\x1B\x4D", "\x1B\x50\x12"},               // элит   
	{"\x0F", "\x12"},                          // сжатый 
	{"\x0E", "\x14"},                          // широкий
	{"\x1B\x53\x30\x0F", "\x12\x1B\x54"},      // сжатый подстрочный 
	{"\x1B\x4D\x0F", "\x1B\x50\x12"},          // сжатый элит
	{NULL, NULL}
};

/**
*	Италик
*/
TagCode ItalicLP[] = 
{
	{"\x1B\x34", "\x1B\x35"},
	{NULL, NULL}
};

/**
*	Подчеркнутый
*/
TagCode UnderlineLP[] =
{
	{"\x1B\x2D\x31", "\x1B\x2D\x30"},
	{NULL, NULL}
};

/**
*	Жирный 
*/
TagCode BoldLP[] =
{
	{"\x1B\x45", "\x1B\x46"},  
	{"\x1B\x47", "\x1B\x48"},     // двойной удар
	{NULL, NULL}
};

/**
*	Все тэги форматирования
*/
TagCode AllTags[] = 
{
	{"\x1B\x34", "\x1B\x35"},
	{"\x1B\x2D\x31", "\x1B\x2D\x30"}, //Вкл./выкл режим подчеркивания
	{"\x1B\x45", "\x1B\x46"},  
	{"\x1B\x47", "\x1B\x48"},     // двойной удар
	{NULL, NULL}
};

/**
*	Определяем специальные сиволы
*/

/**
*	Указывает на строчку форматирования таблицы
*/
char simbFrmStr[] = "@-";

/**
*	Указывает на конец таблицы, если в строчке больше нет других символов
*/
char simbEndTable[] = "}";

/**
*	Указывает на область ячеек. Пример:  A1:B5
*/
char simbRange[] = ":";				

/**
*	Указывает на ассоциативную адрессацию. Пример:  col1::row2
*/
char simbAssocCrd[] = "::";

/**
*	Символ указывает на присутствие в примечании (поле Prim) специальой 
*	команды. Пример: "A1 :rowhide"
*/
char simbCmd[] = ":";

/**
*		Команды
*/

/**
*	Скрыть ряд
*/
char cmdRowHide[]  = "rowhide";

/**
*	Показать ряд
*/
char cmdRowShow[] = "rowshow";

/**
*	Скрыть колонку
*/
char cmdColHide[] = "colhide";

/**
*	Показать колонку
*/
char cmdColShow[] = "colshow";

/**
*	Установить стиль
*   1             | имя стиля          | None
*   2             | цвет текста        | "AABBFF" - (0xAA) -синий, (0xBB) - зеленый,
*                 |                    |            (0xFF) - красный
*   3             | цвет фона          | "FF00FF"
*   4             | флаг Italic        | "t" - установить данный стиль
*   5             | флаг жирного шрифта| "f" - оставляет все как есть (не отменяет символы форматирования LPrint)
*   6             | флаг подчеркивания | "t"
*   7             | название шрифта    | "Arial Cyr"
*   8             | размер шрифта      | "12"
*  :textformat: "S1", "000000","ffffff", "f","t", "f", "Arial Cyr", "10"
*/
char cmdCellStyle[] = "textformat";
