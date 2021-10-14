/**
* Модуль функций работы со строками
* @file
*/

#include "econvert.h"

/**
* Слияние строк
*/
char *concatenate(char *str1, char *str2)
{
    int i,len,len2;

    len = strlen(str1);
    len2 = strlen(str2);
    for(i=0; i < len2; i++)
    {
        str1[len+i]=str2[i];
    }
    str1[len+i]='\x00';
    return (str1);
}

/**
* Усечение строки
*/
char *trim_space(char *str)
{
    int len;
    char *ret;

    len=strlen(str);
    ret= (char*) calloc(len+1, sizeof(char));
    
    ret[0]='\x00';
    
    concatenate(ret, str);
    
    str = strfree(str);
    return ret;
}

/**
*   Удаление начальных и завершающих пробелов из строки
*/
char *strtrim(char *str, BOOL bFree)
{
    char *ret = NULL;
    ret = strtrim_left(str, FALSE);
    ret = strtrim_right(ret, TRUE);
    
    if (bFree)
        str = strfree(str);
    return ret;
}


char *strtrim_left(char *str, BOOL bFree) 
{
    char *ret = NULL;
    char *pointer = str;
    
    if (str)
        ret = strcopy(str);
    else
    {
        ret = strgen_empty();
        return ret;
    }

    while (isspace(*pointer)) 
        ++pointer;

    memmove(ret, pointer, strlen(pointer) + 1);

    if (bFree)
        str = strfree(str);
    return ret;
}


char *strtrim_right(char *str, BOOL bFree)
{
    char *ret = (char *) calloc(strlen(str) + 1, sizeof(char));
    char *end = str + strlen(str);

    while ((end != str) && isspace(*(end-1))) 
        --end;

    memmove(ret, str, end - str);
    ret[end - str] = '\0';

    if (bFree)
        str = strfree(str);
        
    return ret;
}

static int to_utf8(char *from, char *to, const char *codepage)
{
    size_t Lfrom, Lto;
    int ret;
    iconv_t d = iconv_open("UTF-8", codepage);
    
    Lfrom = strlen(from);
    Lto = 2 * Lfrom;
    ret = iconv(d, &from, &Lfrom, &to, &Lto);
    iconv_close(d);
    return ret;
}

char *cp1251_to_utf8(char *from, BOOL bFree)
{
    char *result = NULL;

    if (strempty(from))
        result = strgen_empty();
    else
    {
        result=(char *)calloc(strlen(from) * 2 + 1, sizeof(char));
        to_utf8(from, result, "CP1251");
    }
    
    if (bFree)
        from = strfree(from);
    return result;
}

char *cp866_to_utf8(char *from, BOOL bFree)
{
    char *result = NULL;
    
    if (strempty(from))
        result = strgen_empty();
    else
    {    
        result = (char *) calloc(strlen(from) * 2 + 1, sizeof(char));
        to_utf8(from, result, "CP866");
    }
    
    if (bFree)
        from = strfree(from);
    return result;
}

char *strreplace_old(char *src, char *from, char *to, BOOL bFree)
{
    size_t size    = strlen(src) + 1;
    size_t fromlen = strlen(from);
    size_t tolen   = strlen(to);
    char *value = (char *) malloc(size);
    char *dst = value;
    char *psrc = src;
    
    if (value != NULL)
    {
        for ( ;; )
        {
            const char *match = strstr(psrc, from);
            if (match != NULL)
            {
                size_t count = match-src;
                char *temp;
                size += tolen-fromlen;
                temp = (char *)realloc(value, size);
                if (temp == NULL)
                {
                    value = strfree(value);
                    return NULL;
                }
                dst = temp+(dst-value);
                value = temp;
                memmove(dst, psrc, count);
                psrc += count;
                dst += count;
                memmove(dst, to, tolen);
                psrc += fromlen;
                dst += tolen;
            }
            else /* No match found. */
            {
                strcpy(dst, psrc);
                break;
            }
        }
    }
   
    if (bFree)
        src = strfree(src);
    return value;
}

char *strreplace(char *str, const char *from, const char *to, BOOL bFree)
{
    char *ret = NULL;
    char *r = NULL;
    const char *p;
    const char *q;
    size_t oldlen = strlen(from);
    size_t count, retlen, newlen = strlen(to);

    if (oldlen != newlen) 
    {
        for (count = 0, p = str; (q = strstr(p, from)) != NULL; p = q + oldlen)
            count++;
        /* this is undefined if p - str > PTRDIFF_MAX */
        retlen = p - str + strlen(p) + count * (newlen - oldlen);
    } 
    else
        retlen = strlen(str);

    if ((ret = (char *) malloc(retlen + 1)) == NULL)
    {
        printf("Memory allocation error\n");
        return NULL;
    }

    for (r = ret, p = str; (q = strstr(p, from)) != NULL; p = q + oldlen) 
    {
        /* this is undefined if q - p > PTRDIFF_MAX */
        ptrdiff_t l = q - p;
        memcpy(r, p, l);
        r += l;
        memcpy(r, to, newlen);
        r += newlen;
    }
    strcpy(r, p);

    if (bFree)
    {
        //if (DBG_MODE) logAddLine("Str free [%x]", str);
        str = strfree(str);
    }
        
    return ret;
}


char *strreplace_all(char *src, sr *r)
{
    char *ret = src;
    int i;

    for(i=0; r[i].search; i++)
        ret = strreplace(ret, r[i].search, r[i].replace);

    //ret = strreplace(ret, "\\", "/");
    return ret;
}


static  char    *buff;      /**< buffer for strings */
static  char    **ptrs;     /**< buffer for pointers to strings */

static  char    *bufp;      /**< current buffer pointer */
static  char    **ptrp;     /**< next word address */

static  int words;          /**< number of words in string */
static  int space;          /**< number of characters to store */


const char *skip(const char *str, const char *white)
{
    while (*str && strchr(white,*str) != CNULL)
        str++;
    return str;
}


void copy(char ch)
{
    if (buff != CNULL)
        *bufp ++= ch;
    else
        space++;
}


void newword(char *cp)
{
    if (buff != CNULL)
        *ptrp++ = cp;
    else
        words++;
}


void subsplit(const char    *str,const char *delim,const char *quotes)
{
    int sloshed;
    char quotec;
    
    words = 0;
    space = 0;
    str = skip(str, delim);
    if (*str)
        newword(bufp);

    for (quotec='\0', sloshed=FALSE; *str; str++)
    { 
        if (quotec != '\0')
            /* in quotes */
            if (sloshed)
            /* in quotes after a slosh */
            { 
                if (*str != quotec && *str != SLOSH)
                    /* not something that's escaped */
                    copy(SLOSH);

                copy(*str);

                /* forget slosh */
                sloshed = FALSE;
            }
            else
                /* in quotes not after a slosh */
                if (*str == quotec)
                    /* leave quotes */
                    quotec='\0';
                else
                    if ((sloshed=(*str == SLOSH)))
                        /* notice slosh */
                        ;
                    else
                        copy(*str);
        else
            /* not in quotes */
            if (sloshed)
            {
                /* not in quotes after a slosh */
                if (*str != SLOSH
                    && strchr(quotes,*str) == CNULL
                    && strchr(delim,*str) == CNULL)
                    copy(SLOSH);

                copy(*str);

                /* forget slosh */
                sloshed=FALSE;
            }
            else
            {
                /* not in quotes not after a slosh */
                
                if (strchr(quotes,*str) != CNULL)
                    /* enter quotes */
                    quotec = *str;
                else
                    if (strchr(delim,*str) != CNULL)
                    {
                        /* find next word */
                        str=skip(str,delim);
                        if (*str)
                        { 
                            copy('\0');
                            newword(bufp);
                        }

                        str--;
                    }
                    else
                        if ((sloshed=(*str == SLOSH)))
                            ;
                        else
                            copy(*str);
            }
        }

    /* catch trailing sloshes */
    if (sloshed)
        copy(SLOSH);

    copy('\0');
    newword(CNULL);
}


char **qstrsplit(const char *str,const char *delim,const char *quotes)
{
    /* default delimiters */
    if (delim == CNULL)
        delim=" \t\n";

    /* mark pass one */
    buff=CNULL;

    /* count words and characters */
    subsplit(str,delim,quotes);

    /* allocate room for characters */
    if ((buff=vnew(char, space)) == CNULL)
        return CPNULL;

    /* allocate room for words */
    if ((ptrs=vnew(char *,words+1)) == CPNULL)
    { 
        buff = strfree(buff);
        return CPNULL;
    }

    /* initialise pointers */
    bufp=buff;
    ptrp=ptrs;

    /* point to allocated space */
    *ptrp++ = buff;

    /* copy words into buffer */
    subsplit(str,delim,quotes);

    return &ptrs[1];    /* return pointer to words */
}


char **strsplit(char *str, char *delim)
{
    return qstrsplit(str, delim, DEFQUOTES);
}


/**
*   Привести строку к нижнему регистру
*/
char *strlwr_lat(char *pstr)
{
    char *p = pstr;
    while (*p)
    {
        if(*p >= 'A' && *p <= 'Z')
            *p = tolower(*p);
        p++;
    }
    return pstr;
}


/**
*   Привести строку к верхнему регистру
*/
char *strupr_lat(char *pstr)
{
    char *p = pstr;
    while (*p)
    {
        if(*p >= 'A' && *p <= 'Z')
            *p = toupper(*p);
        p++;
    }
    return pstr;
}


char *strnset(char *str, int ch, size_t n)
{  
    for(int i=0; i < (int) n; i++)
    {
        if(str[i] == STR_NULL ) 
            return (str); // return when find null
        str[i] = ch;
    }
    return  (str);
}


/**
* Проверка на пустую строку
*/
BOOL strempty(char *str)
{
    int len = 0;
    
    if (str == NULL)
        return TRUE;
        
    len = strlen(str);
    if (len > 0)
        return FALSE;
    else
        return TRUE;
}


/**
*   Проверка на строку с значащами символами. Есть ли в строке содержание?
*   Если строка пустая или в строке одни пробелы, то фунция возращает FALSE
*   иначе TRUE
*/
BOOL strcontent(char *str)
{
    BOOL is_empty = strempty(str);
    
    if (is_empty)
        return FALSE;

    
    char *p = str;
    while (*p)
    {
        if ((*p != ' ') && (*p != '\t') && (*p != '\n') && (*p != '\r'))
            return TRUE;
        p++;
    }
    
    return FALSE;
}


/**
* Найти слово в строке
*/
BOOL find_word(char *source, char *search)
{
    int i = 0;                //  Position inside source
    int search_length;        //  Length of string search
    int source_length;        //  Length of string source
    BOOL found = FALSE;       //  Flag 1-> found the value

    search_length = strlen(search);
    source_length = strlen(source);

    /*  While we haven't found it and we haven't readched the      */
    /*  point where our current position plus the length of the    */
    /*  string we are looking for is longer than the string itself,*/
    /*  check the next search_length characters for a match        */
    while (!found && ((i+search_length) <= (source_length)))
    {
        found = (strncasecmp(source+i, search, search_length) == 0);
        i++;
    }
 
    return (found);
}

/**
*   C substring function: It returns a pointer to the substring 
*/ 
char *substr(char *str, unsigned int position, unsigned int length) 
{
    char *ret = NULL;
    unsigned int c;
 
    if ((position + length) > strlen(str))
        length = strlen(str) - position;
        
    ret = (char *) calloc(length + 1, sizeof(char));
 
    if (ret == NULL)
    {
        printf("FUNCTION [substr]. Unable to allocate memory.\n");
        return NULL;
    }
 
    for (c = 0 ; c < length ; c++)
        *(ret + c) = (char) str[position + c];
 
    *(ret + c) = '\0';
 
    return ret;
}

char *strright_pos(char *str, int position)
{
    return substr(str, position, strlen(str)-position);
}

/**
*   Взять length символов с начала строки.
*/
char *strleft(char *str, int length, BOOL bFree)
{
    char *ret = substr(str, 0, length);
    if (bFree)
        str = strfree(str);
    return ret;
}

/**
*   Взять length символов с конца строки.
*/
char *strright(char *str, int length, BOOL bFree)
{
    int len = strlen(str);
    char *ret = substr(str, len-length, len);
    if (bFree)
        str = strfree(str);
    return ret;
}

/** 
*   detecting whether base is starts with str
*/
BOOL startswith(char* base, char* str) 
{
    return (strstr(base, str) - base) == 0;
}


/** 
*   detecting whether base is ends with str
*/
BOOL endswith(char* base, char* str) 
{
    int blen = strlen(base);
    int slen = strlen(str);
    return (blen >= slen) && (0 == strcmp(base + blen - slen, str));
}


/** 
*   getting the first index of str in base
*/
int strfind(char* base, char* str) 
{
    return strfind_offset(base, str, 0);
}

int strfind_offset(char* base, char* str, int startIndex) 
{
    int result;
    int baselen = strlen(base);
    
    // str should not longer than base
    if (strlen(str) > baselen || startIndex > baselen) 
    {
        result = -1;
    } 
    else 
    {
        if (startIndex < 0 ) 
        {
            startIndex = 0;
        }
        char *pos = strstr(base + startIndex, str);
        if (pos == NULL) 
        {
            result = -1;
        } 
        else 
        {
            result = pos - base;
        }
    }
    return result;
}

/**
*   Найти символ в строке начиная с startIndex
*/
int strfind_char(char *str, char c, int startIndex) 
{
    int len = strlen(str);
    int ret = -1;
    
    for (unsigned int i=startIndex; i < len; i++)
        if (str[i] == c)
        {
            ret = i;
            break;
        }
         
    return ret;
}

/**
*   Количество подстрок
*/
unsigned int strfind_count(char *base, char *str) 
{
    return strfind_count_offset(base, str, 0);
}

unsigned int strfind_count_offset(char *base, char *str, int startIndex) 
{
    unsigned int result = 0;
    unsigned int baselen = strlen(base);
    
    // str should not longer than base
    if (strlen(str) > baselen || startIndex >= baselen) 
        result = 0;
    else 
    {
        if (startIndex < 0 ) 
            startIndex = 0;
        char *pos = strstr(base + startIndex, str);
        if (pos == NULL) 
            result = 0;
        else 
            result = 1 + strfind_count_offset(base, str, startIndex + 1);
    }
    return result;
}

/** 
*   use two index to search in two part to prevent the worst case
 *  (assume search 'aaa' in 'aaaaaaaa', you cannot skip three char each time)
 */
int strfind_last(char *base, char *str) 
{
    int result;
    
    // str should not longer than base
    if (strlen(str) > strlen(base)) 
    {
        result = -1;
    } 
    else 
    {
        int start = 0;
        int endinit = strlen(base) - strlen(str);
        int end = endinit;
        int endtmp = endinit;
        while(start != end) 
        {
            start = strfind_offset(base, str, start);
            end = strfind_offset(base, str, end);

            // not found from start
            if (start == -1) 
            {
                end = -1; // then break;
            } 
            else 
                if (end == -1) 
                {
                    // found from start
                    // but not found from end
                    // move end to middle
                    if (endtmp == (start+1)) 
                    {
                        end = start; // then break;
                    } 
                    else 
                    {
                        end = endtmp - (endtmp - start) / 2;
                        if (end <= start) 
                        {
                            end = start+1;
                        }
                        endtmp = end;
                    }
                } 
                else 
                {
                    // found from both start and end
                    // move start to end and
                    // move end to base - strlen(str)
                    start = end;
                    end = endinit;
                }
        }
        result = start;
    }
    return result;
}


/**
*   Проверка на равество двух строк
*/
BOOL strequal(char *str1, char *str2)
{
    if ((str1 == NULL) && str2 )
        return FALSE;
    if ((str2 == NULL) && str1 )
        return FALSE;
    return strcmp(str1, str2) == 0;
}

/**
*   Строка наоборот
*/
char *strreverse(char *str) 
{
   int length, c;
   char *begin, *end, temp;
 
   length = strlen(str);
   begin  = str;
   end    = str;
 
   for (c = 0; c < length - 1; c++)
      end++;
 
   for (c = 0; c < length/2; c++)
   {        
      temp   = *end;
      *end   = *begin;
      *begin = temp;
 
      begin++;
      end--;
   }
   return str;
}


/**
*   Установить все символы строки в c
*/
char *strset(char *str, char c, int length)
{
    unsigned int len = 0;
    
    if (length < 0)
        len = strlen(str);
    else
        len = length;
    
    for (unsigned int i=0; i < len; i++)
        str[i] = c;
    return str;
};


/**
*   Создать строку и становить все символы строки в c
*/
char *strgen(char c, unsigned int len)
{
    char *s = (char *) calloc(len + 1, sizeof(char));
    char *ret = strset(s, c, len);
    ret[len] = '\0';
    return ret;
}

char *strgen_empty(void)
{
    char *s = (char *) calloc(1, sizeof(char));
    s[0] = '\0';
    return s;
}

/**
*   Функция "съедает" n символов в начале строки
*/
char *str_eat_chars(char *str, size_t n)
{
    char *result;
    char *emptystr = "";
    unsigned int len = strlen(str);
    
    n = min(n, len);
    len -= n;

    if (len)
    {
        result = (char *) calloc(len+1, sizeof(char));
        strcpy(result, str + n);
    }
    else
        return emptystr;
    
    return result;
}
          
          
char *strfree(char *str)
{      
    // if (DBG_MODE) logWarning("[%x] free", str);
    return (char *) memfree((void *) str);
}


/**
*   Форматированный вывод в строку с выделением памяти.
*/
char *strprintf(char *str, char *fmt,...)
{
    char buffer[MAX_STR];
    va_list ap;

    va_start(ap, fmt);
    vsprintf(buffer, fmt, ap);

    if (str == NULL)
        // str = (char *) calloc(strlen(buffer)+1, sizeof(char));
        str = strdup(buffer);
    else
        strcpy(str, buffer);
    
    return str;
}


/**
*   Форматированный вывод с дублированием строки.
*/
char *dupprintf(char *fmt,...)
{
    char buffer[MAX_STR];
    va_list ap;

    va_start(ap, fmt);
    vsprintf(buffer, fmt, ap);

    return strdup(buffer);
}


/**
*   Форматированный вывод с заменой строки.
*/
char *replaceprintf(char *str, char *fmt,...)
{
    char buffer[MAX_STR];
    va_list ap;

    va_start(ap, fmt);
    vsprintf(buffer, fmt, ap);

    strfree(str);
    return strdup(buffer);
}


char *strconcatenate(char *str1, char *str2, BOOL bFree)
{
    char *ret = (char *) calloc((strlen(str1)+strlen(str2)+1), sizeof(char));

    ret = strcpy(ret, str1);    
    if (bFree)
        str1 = strfree(str1);
        
    return concatenate(ret, str2);
}


/**
*   Заменить символ с номером char_index на new_char в строке
*/
char *strreplacechar(char *str, unsigned int char_index, char new_char)
{
    unsigned int len = strlen(str);
    
    if (char_index<0 || char_index>=len)
        return str;
    
    for (unsigned int i=0; i < len; i++)
    {
        if (i == char_index)
        {
            *(str+i) = new_char;
            break;
        }
    }
    return str;
}

/**
*   Заменить символ с номером char_index на строку new_str в строке
*/
char *strreplace_pos(char *str, unsigned int pos, char *new_str, BOOL bFree)
{
    unsigned int len = strlen(str);
    
    if ((pos < 0) || (pos >= len))
        return str;
        
    char *result = (char *) calloc(len + strlen(new_str), sizeof(char));
    
    for (unsigned int i=0; i < len; i++)
    {
        if (i == pos)
        {
            char *v = strleft(str, i);
            strcpy(result, v);
            strcat(result, new_str);
            strcat(result, str + i + 1);
            v = strfree(v);
            break;
        }
    }
    
    if (bFree)
        str = strfree(str);
    return result;
}

/**
*   Взять подстроку слева до указанного символа
*/
char *strleft_to(char *str, char symb, BOOL bFree)
{
    char *p = str;
    char *result = NULL;
    
    while (*p)
    {
        if (*p == symb)
        {
            result = strleft(str, p-str);
            if (bFree)
                str = strfree(str);
            return result;
        }
        p++;
    }
    
    result = strcopy(str);
    if (bFree)
        str = strfree(str);
    return result;
}

/**
*   Взять подстрку справа до указанного символа
*/
char *strright_to(char *str, char symb, BOOL bFree)
{
    char *start = str+strlen(str)-1;
    char *p = start; //Перейти на последний элемент
    char *result = NULL;
    
    while (*p)
    {
        if (*p == symb)
        {
            result = strright(str, start - p);
            if (bFree)
                str = strfree(str);
            return result;
        }
        p--;
    }
    
    result = strcopy(str);
    if (bFree)
        str = strfree(str);
    return result;
}

/**
*   Функция определяет есть ли в строке ведущий 0
*/
BOOL is_leading0(const char *str)
{
    if ((strlen(str) > 2) && (str[0] == '0') && (str[1] == '.'))
        // Это число с плавающей запятой
        return FALSE;
    if ((strlen(str) > 1) && (str[0] == '0')) 
        // Это ведущий 0
        return TRUE;
    // Во всех остальных случаях это не ведущий ноль
    return FALSE;
}

/**
*   Функция определяет является ли строка числом
*/
BOOL isnumeric(const char *str, BOOL bCheckINN)
{
    if (str == NULL || *str == '\0' || isspace(*str))
    {
        // if (DBG_MODE) logAddLine("WARNING. Empty string <%s>", str);
        return FALSE;
    }
    
    char *p;
    strtod(str, &p);
    BOOL result = (*p == '\0');
    // if (DBG_MODE) logAddLine("Is NUM <%s> [%s] [%d]", str, p, result);
    
    // Впереди явным образом задается 0 значит это ведущий 0 и это не число
    if (is_leading0(str))
    {
        return FALSE;
    }
    
    // Число может быть ИНН 
    // Обработка такого случая
    if (result && bCheckINN)
    {
        // Если все символы строки цифровые и длина строки 10 или 12,
        // то считается что это ИНН
        BOOL is_digit = TRUE;
        int i = 0;
        while (i < strlen(str) && is_digit == TRUE)
        {
            if ((str[i] == '.') || (str[i] == ','))
            {
                // Нашли <.> или <,> Значит это просто большое число
                i = -1; 
                break;
            }
            is_digit = isdigit(str[i]);  
            i++;
        }
        result = (!((i == 10) || (i == 12)));
        if (DBG_MODE) logAddLine("INN <%s> [%d] [%d]", str, i, result);
    }
    
    return result;
}

int decimal_point(const char *number)
{
    if (number == NULL || *number == '\0' || isspace(*number))
        return -1;
    
    // Указатель на конец строки
    int i = 0;
    int last = strlen(number);
    int count = 0;  // Количество цифр после десятичной точки
    
    for (i = last; i--; i >= 0)
    {
        // if (DBG_MODE) logAddLine("Find decimal point <%c>", *(number+i));
        if (*(number+i) == '.')
            break;
        else
            count++;
    }
    if (i < 0)
        // Десятичной точки вообще нет, значит нет и цифр после точки
        count = 0;
    return count;    
}

char *strinit(char *str, char *init)
{
    // if (DBG_MODE) logAddLine("strinit START");
    
    str = strfree(str);

    // if (DBG_MODE) logAddLine("strinit STOP");
    return init;
}

/**
*   Аналог strdup
*/
char *strcopy(const char *str)
{
    size_t len;
    char *copy = NULL;

    if (!str)
        return NULL;

    len = strlen(str) + 1;
    copy = (char *) calloc(len, sizeof(char));

    if (copy)
        // memcpy(copy, str, len);
        strcpy(copy, str);
    else
        if (DBG_MODE) logErr("Memory allocation error");

    return copy;
}
