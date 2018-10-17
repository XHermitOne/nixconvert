/**
* Модуль функций работы с памятью
* @file
*/

#include "econvert.h"


/**
* Прверка корректности выделения памяти
* Не рабочая функция
*/
void memcheck(void)
{
    if (mcheck(NULL) != 0) 
        if (DBG_MODE) logWarning("Memory check error"); 
    else
        if (DBG_MODE) logAddLine("Memory check ok"); 
}


/**
* Корректное освобождение памяти
*/
void *memfree(void *ptr)
{        
    if (ptr != NULL)
        {            
            enum mcheck_status m_status;
            m_status = mprobe(ptr);
            
            if (m_status == MCHECK_FREE)
                if (DBG_MODE) logErr("[%x] released", ptr);
            else if (m_status == MCHECK_DISABLED)
                if (DBG_MODE) logErr("[%x] disabled", ptr);
            else if (m_status == MCHECK_HEAD)
                if (DBG_MODE) logErr("[%x] head", ptr);
            else if (m_status == MCHECK_TAIL)
                if (DBG_MODE) logErr("[%x] tail", ptr);
            else
            {            
                // if (DBG_MODE) logWarning("[%x] Memory free", ptr);
                free(ptr);
                ptr = NULL;
            }
        }
    return ptr;
}
