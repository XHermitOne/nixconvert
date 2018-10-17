#if !defined( __SPCSIMB_H )
#define __SPCSIMB_H

#include <stdlib.h>

// #define SPCSIMB

//
struct TagCode
{
	char *Seq;		// начальная последовательность последовательность
	char *SeqEnd;   // конечная последовательность   
};

/**
*	Разделители страниц 
*/
extern TagCode DivPage[];

/**
*	Заменяемые последовательности
*/
extern TagCode ChangeTag[];

/**
*	Удаляемые последовательности
*/
extern TagCode EmptyLP [];

/**
*	Италик
*/
extern TagCode ItalicLP[];

/**
*	Подчеркнутый
*/
extern TagCode UnderlineLP[];

/**
*	Жирный 
*/
extern TagCode BoldLP[];

/**
*	Все тэги форматирования
*/
extern TagCode AllTags[];

/**
*	Определяем специальные сиволы
*/

/**
*	Указывает на строчку форматирования таблицы
*/
extern char simbFrmStr[];

/**
*	Указывает на конец таблицы, если в строчке больше нет других символов
*/
extern char simbEndTable[];

/**
*	Указывает на область ячеек. Пример:  A1:B5
*/
extern char simbRange[];

/**
*	Указывает на ассоциативную адрессацию. Пример:  col1::row2
*/
extern char simbAssocCrd[];

/**
*	Символ указывает на присутствие в примечании (поле Prim) специальой 
*	команды. Пример: "A1 :rowhide"
*/
extern char simbCmd[];

/**
*		Команды
*/

/**
*	Скрыть ряд
*/
extern char cmdRowHide[];

/**
*	Показать ряд
*/
extern char cmdRowShow[];

/**
*	Скрыть колонку
*/
extern char cmdColHide[];

/**
*	Показать колонку
*/
extern char cmdColShow[];

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
extern char cmdCellStyle[];

#endif /*__SPCSIMB_H*/
