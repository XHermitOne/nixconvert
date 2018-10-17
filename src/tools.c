/**
* Модуль сервисных функций
* @file
*/

#include "econvert.h"
#include "tools.h"

/**
* checkParm
*
* Checks for the given parameter in the program's command line arguments
* @return the argument number (1 to argc-1) or 0 if not present
*/

int checkParm(const char *check, int argc, char **argv)
{
    int i;

    for (i = 1; i < argc; i++)
    {
        if (!strcasecmp(check, argv[i]))
            return i;
    }

    return 0;
}

/**
* Вывод ошибок в виде окна сообщений
* Я использую в основном для отладки
*/
void errBox(char *Fmt,...)
{
    char buffer[128];
    va_list ap;

    va_start(ap, Fmt);
    vsprintf(buffer, Fmt, ap);
	
    //log_color_line(IC_RED_COLOR_TEXT, "Error: ");
    log_color_line(IC_RED_COLOR_TEXT, buffer);
}



char sayError(char *msg) 
{ 
    errBox(msg);
    return 0;
}



char sayMessage(char *Msg, int aOptions)
{
    return sayError(Msg);
}



char *getHomePath(void)
{
    return getenv("HOME");
}


/**
* Проверить существует ли директория
*/
BOOL dir_exists(char *path)
{
    BOOL result = FALSE;
    struct stat st;
    
    if (!strempty(path))
        result = (BOOL) (stat(path, &st) == 0);
    return result;
}


static int do_mkdir(const char *path, mode_t mode)
{
    struct stat st;
    int status = 0;

    if (stat(path, &st) != 0)
    {
        /* Directory does not exist. EEXIST for race condition */
        if (mkdir(path, mode) != 0) // && errno != EEXIST)
            status = -1;
    }
    else if (!S_ISDIR(st.st_mode))
    {
        if (DBG_MODE) logAddLine("ERROR: %s is not directory",path);
        status = -1;
    }

    return (status);
}

/**
* mkpath - ensure all directories in path exist
* Algorithm takes the pessimistic view and works top-down to ensure
* each directory in path exists, rather than optimistically creating
* the last element and working backwards.
*/
int mkpath(const char *path, mode_t mode)
{
    char *pp;
    char *sp;
    int status;
    char *copypath = strcopy(path);

    status = 0;
    pp = copypath;
    while (status == 0 && (sp = strchr(pp, '/')) != 0)
    {
        if (sp != pp)
        {
            /* Neither root nor double slash in path */
            *sp = '\0';
            status = do_mkdir(copypath, mode);
            *sp = '/';
        }
        pp = sp + 1;
    }
    if (status == 0)
        status = do_mkdir(path, mode);
    copypath = strfree(copypath);
    return (status);
}


char CfgPath[] = ".nixconvert";              /**< Папка конфигурациронных файлов*/
char CfgFileName[] = "nixconvert.cfg";       /**< Имя файла конфигурации основных параметров*/

static char FullCfgFileName[MAX_PATH];

static BOOL createCfgPath(char *Path);

static char _sCfgPath[MAX_PATH] = "";

char *getCfgPath(void)
{
    if (!strempty(_sCfgPath))
    {
        if(DBG_MODE) logAddLine("CfgPath: %s", _sCfgPath);
        return _sCfgPath;
    }
        
    char *home_path = getHomePath();
    
    //_sCfgPath = (char*) calloc(strlen(home_path) + 1 + strlen(CfgPath) + 1, sizeof(char));
    strcpy(_sCfgPath, home_path);
    strcat(_sCfgPath, "/");
    strcat(_sCfgPath, CfgPath);

    if (!dir_exists(_sCfgPath)) 
        createCfgPath(_sCfgPath);

    //if(DBG_MODE) logAddLine("Get CfgPath: %s", _sCfgPath);
    
    return _sCfgPath;
}


BOOL createCfgPath(char *Path=NULL)
{
    int do_free=0;
    char *conf_path=Path;
    if (conf_path==NULL)
    {
        conf_path=getCfgPath();
        do_free=1;
    }
        
        
    BOOL result=(BOOL) mkpath(conf_path, S_IRWXU | S_IRGRP | S_IXGRP | S_IROTH | S_IXOTH);
    if(DBG_MODE) logAddLine("Create path: %s\n", conf_path);

    if (do_free)
        conf_path = strfree(conf_path);    
        
    return result;
}

/**
* Cуществует ли файл
* @return TRUE  - exist / FALSE - don't exist
*/
BOOL file_exists(char *FileName)
{
    BOOL bExist = FALSE;
    if (!strempty(FileName))
    {
		FILE *in = fopen(FileName, "rb");
		if (in != NULL)
        {
            bExist = TRUE;
			fclose( in );
        }
    }
    return bExist;
}

/**
* Копирование файла
* @return: 
*/
int copy_file(char *to, char *from)
{
    char newname[MAXPATHLEN+1];
    char answer[20];
    struct stat stf, stt;
    int fdin, fdout;
    int n, code = 0;
    char iobuf[64 * 1024];
    char *dirname = NULL, *s;
    
    if((fdin = open(from, O_RDONLY)) < 0)
    {
        if (DBG_MODE) logAddLine("Cannot read %s", from);
        return (-1);
    }
    fstat(fdin, &stf);
    if ((stf.st_mode & S_IFMT) == S_IFDIR)
    {
        close(fdin);
        if (DBG_MODE) logAddLine("%s is a directory", from);
        return (-2);
    }

    if(stat(to, &stt) >= 0)
    {
        /* Файл уже существует */
        if ((stt.st_mode & S_IFMT) == S_IFDIR)
        {
            /* И это каталог */

            /* Выделить последнюю компоненту пути from */
            if ((s = strrchr(from, '/')) && s[1])
                s++;
            else    
                s = from;

            dirname = to;

            /* Целевой файл - файл в этом каталоге */
            sprintf(newname, "%s/%s", to, s);
            to = newname;

            if (stat(to, &stt) < 0)
                goto not_exist;
        }

        if (stt.st_dev == stf.st_dev && stt.st_ino == stf.st_ino)
        {
            if (DBG_MODE) logAddLine("%s: cannot copy file to itself", from);
            return (-3);
        }
        switch (stt.st_mode & S_IFMT)
        {
            case S_IFBLK:
            case S_IFCHR:
            case S_IFIFO:
                break;

            default:
                if (DBG_MODE) logAddLine("%s already exists, overwrite ? ", to);
                fflush(stdout);

                *answer = '\0';
                gets(answer);

                if (*answer != 'y')
                {     /* NO */
                    close(fdin);
                    return (-4);
                }
                break;
        }
    }

not_exist:

    if (DBG_MODE) logAddLine("COPY %s TO %s\n", from, to);

    if ((stf.st_mode & S_IFMT) == S_IFREG)
    {
        /* Проверка наличия свободного места в каталоге dirname */
        struct statvfs fs;
        char tmpbuf[MAXPATHLEN+1];

        if(dirname == NULL)
        {
            /* То 'to' - это имя файла, а не каталога */
            strcpy(tmpbuf, to);
            if (s = strrchr(tmpbuf, '/'))
            {
                if (*tmpbuf != '/' || s != tmpbuf)
                {
                    /* Имена "../xxx"
                    * и второй случай:
                    * абсолютные имена не в корне,
                    * то есть не "/" и не "/xxx"
                    */
                    *s = '\0';
                }
                else
                {
                    /* "/" или "/xxx" */
                    if(s[1]) s[1] = '\0';
                }
                dirname = tmpbuf;
            }
            else
                dirname = ".";
        }

        if (statvfs(dirname, &fs) >= 0)
        {
            size_t size = (geteuid() == 0 ) ?
                /* Доступно суперпользователю: байт */
                fs.f_frsize * fs.f_bfree :
                /* Доступно обычному пользователю: байт */
                fs.f_frsize * fs.f_bavail;

            if (size < stf.st_size)
            {
                if (DBG_MODE) logAddLine("Not enough free space on %s: have %lu, need %lu", dirname, size, stf.st_size);
                close(fdin);
                return (-5);
            }
        }
    }

    if ((fdout = creat(to, stf.st_mode)) < 0)
    {
        if (DBG_MODE) logAddLine("Can't create %s", to);
        close(fdin);
        return (-6);
    }
    else
    {
        fchmod(fdout, stf.st_mode);
        fchown(fdout, stf.st_uid, stf.st_gid);
    }

    while (n = read (fdin, iobuf, sizeof iobuf))
    {
        if (n < 0)
        {
            if (DBG_MODE) logAddLine("read error");
            code = (-7);
            goto done;
        }
        if (write(fdout, iobuf, n) != n)
        {
            if (DBG_MODE) logAddLine("write error");
            code = (-8);
            goto done;
        }
    }

done:
    close (fdin);
    close (fdout);

    /* Проверить: соответствует ли результат ожиданиям */
    if (stat(to, &stt) >= 0 && (stt.st_mode & S_IFMT) == S_IFREG)
    {
        if (stf.st_size < stt.st_size)
        {
            if (DBG_MODE) logAddLine("File has grown at the time of copying");
        }
        else
            if (stf.st_size > stt.st_size)
            {
                if (DBG_MODE) logAddLine("File too short, target %s removed", to);
                unlink(to);
                code = (-9);
            }
    }
    return code;
}

/**
*   Удалить файл
*/
BOOL del_file(char *FileName)
{
    if (unlink(FileName) >= 0 )
    {
        if (DBG_MODE) print_color_txt(IC_YELLOW_COLOR_TEXT, "Delete file <%s>\n", FileName);
        return TRUE;
    }
    else
    {
        if (DBG_MODE) print_color_txt(IC_RED_COLOR_TEXT, "ERROR. Delete file <%s>\n", FileName);
        return FALSE;
    }
}

/**
*   Прочитать файл и вернуть массив символов
*/
unsigned int load_txt_file_size(char *FileName, char *result)
{
	unsigned int size = 0;
	FILE *f = fopen(FileName, "rb");
	if (f == NULL) 
	{ 
		result = NULL;
        if (DBG_MODE) logAddLine("ERROR. load_txt_file_size. Open file <%s>", FileName);
		return -1; // -1 means file opening fail 
	} 
	fseek(f, 0, SEEK_END);
	size = ftell(f);
	fseek(f, 0, SEEK_SET);
	result = (char *) malloc(size+1);
	if (size != fread(result, sizeof(char), size, f)) 
	{ 
		result = strfree(result);
        if (DBG_MODE) logAddLine("ERROR. load_txt_file_size. Read file <%s>", FileName);
		return -2; // -2 means file reading fail 
	} 
	fclose(f);
	result[size] = 0;
    if (DBG_MODE) logAddLine("INFO. load_txt_file_size. File <%s>\tSize <%d>", FileName, size);
	return size;
}

char *load_txt_file(char *FileName)
{
    char *result;
	unsigned int size = 0;
	FILE *f = fopen(FileName, "rb");
    
	if (f == NULL) 
	{ 
        if (DBG_MODE) print_color_txt(IC_RED_COLOR_TEXT, "ERROR. load_txt_file. Open file <%s>\n", FileName);
		return NULL;
	} 
    
	fseek(f, 0, SEEK_END);
	size = ftell(f);
	fseek(f, 0, SEEK_SET);
	result = (char *) calloc(size + 1, sizeof(char));
    
	if (size != fread(result, sizeof(char), size, f)) 
	{ 
		result= strfree(result);
        if (DBG_MODE) print_color_txt(IC_RED_COLOR_TEXT, "ERROR. load_txt_file. Read file <%s>\n", FileName);
		return NULL;
	} 
	fclose(f);
    
    //Удалить все 0 и заменить их на пробелы
    for (unsigned int i=0; i < size; i++)
        if (result[i] == '\0')
            result[i] = ' ';
	result[size] = '\0';
    
    if (DBG_MODE) print_color_txt(IC_WHITE_COLOR_TEXT, "INFO. load_txt_file. File <%s>\tSize <%d>\n", FileName, size);
	return result;
}

double max(double a, double b)
{
    return (((a) > (b)) ? (a) : (b));
}

double min(double a, double b)
{
    return (((a) < (b)) ? (a) : (b));
}


/**
* C++ version char* style "itoa".  It would appear that
* itoa() isn't ANSI C standard and doesn't work with GCC on Linux
* @param nValue value to convert to ascii
* @param szResult buffer for the result
* @param nBase base for conversion
* @return value converted to ascii
*/
char *itoa(int value, int base)
{
    char *result = (char *) calloc(MAX_NUM_LEN + 1, sizeof(char)); 
    
	// check that the base if valid
	if (base < 2 || base > 16) 
		 return NULL;

	char *out = result;
	int quotient = value;

	do 
    {
		*out = "0123456789abcdef"[ abs( quotient % base ) ];
		++out;
		quotient /= base;
	} while (quotient);

	// Only apply negative sign for base 10
	if (value < 0 && base == 10) 
        *out++ = '-';

	*out = 0;

	return result;
}


/**
 * Double to ASCII
 */
static double PRECISION = 0.00000000000001;
static int MAX_NUMBER_STRING_SIZE = 32;
char * dtoa(double n) 
{
    //char s[80];
    char *s = (char *) calloc(80, sizeof(char));
    
    if (isnan(n)) 
        strcpy(s, "nan");
    else if (isinf(n))
        strcpy(s, "inf");
    else if (n == 0.0)
        strcpy(s, "0");
    else 
    {
        int digit, m, m1;
        char *c = s;
        int neg = (n < 0);
        if (neg)
            n = -n;
            
        m = log10(n);
        int useExp = (m >= 14 || (neg && m >= 9) || m <= -9);
        if (neg)
            *(c++) = '-';
            
        if (useExp) 
        {
            if (m < 0)
                m -= 1.0;
            n = n / pow(10.0, m);
            m1 = m;
            m = 0;
        }
        if (m < 1.0) 
            m = 0;
            
        while (n > PRECISION || m >= 0) 
        {
            double weight = pow(10.0, m);
            if (weight > 0 && !isinf(weight)) 
            {
                digit = floor(n / weight);
                n -= (digit * weight);
                *(c++) = '0' + digit;
            }
            if (m == 0 && n > 0)
                *(c++) = '.';
            m--;
        }
        if (useExp) 
        {
            int i, j;
            *(c++) = 'e';
            if (m1 > 0) 
                *(c++) = '+';
            else 
            {
                *(c++) = '-';
                m1 = -m1;
            }
            m = 0;
            while (m1 > 0) 
            {
                *(c++) = '0' + m1 % 10;
                m1 /= 10;
                m++;
            }
            c -= m;
            for (i = 0, j = m-1; i<j; i++, j--) 
            {
                c[i] ^= c[j];
                c[j] ^= c[i];
                c[i] ^= c[j];
            }
            c += m;
        }
        *(c) = '\0';
    }
    return s;
}

int bytes_to_int(BYTE *bytes)
{
    if (bytes == NULL)
        return 0;
        
    int result;
    result = (bytes[1] << 8)  | 
             bytes[0];
    return result;
}

double bytes_to_double(BYTE *bytes)
{
    if (bytes == NULL)
        return 0.0;

    double result;
    memcpy(&result, bytes, sizeof(double));   
    return result;
}

long time_to_long(char *buffer)
{
    long Time;
    int i, j, k;
    for(Time = i = j = 0; i < 3; i++)
    {
        sscanf(&buffer[j], "%d", &k);
        Time = Time * 60 + k;
        for(; buffer[j] != ':' && buffer[j]; j++);
        if (!buffer[j]) 
            break; 
        else 
            j++;
    }
    return Time;
}


char *long_to_time(long ltime)
{
    char *buffer = (char *) calloc(9, sizeof(char));
    if (ltime >= 0 && ltime != 1000000000) 
        sprintf(buffer, "%02ld:%02ld:%02ld", (ltime/3600)%24, (ltime/60)%60, ltime%60);
    else 
        strcpy(buffer,"##:##:##");
    if (ltime != 1000000000 && ltime >= 86400) 
        sprintf(&buffer[9], "%06ld", ltime/86400);
    return buffer;
}

long get_now_time(void)
{
    time_t sys_time;
    return (long) time(&sys_time);
}

/**
*   Нормализация пути
*/
char *norm_path(char *src, size_t src_len, BOOL bFree) 
{
    char * res;
    size_t res_len;

    if (src_len < 0)
        src_len = strlen(src)-1;
        
    const char *ptr = src;
    const char *end = &src[src_len];
    const char *next;

    if (src_len == 0 || src[0] != '/') 
    {
        // relative path

        char pwd[PATH_MAX];
        size_t pwd_len;

        if (getcwd(pwd, sizeof(pwd)) == NULL) 
        {
            return NULL;
        }

        pwd_len = strlen(pwd);
        res = (char *) malloc(pwd_len + 1 + src_len + 1);
        memcpy(res, pwd, pwd_len);
        res_len = pwd_len;
    } 
    else 
    {
        res = (char *) malloc((src_len > 0 ? src_len : 1) + 1);
        res_len = 0;
    }

    for (ptr = src; ptr < end; ptr=next+1) 
    {
        size_t len;
        next = (char *) memchr(ptr, '/', end-ptr);
        if (next == NULL) 
        {
            next = end;
        }
        len = next-ptr;
        switch(len) 
        {
            case 2:
                if (ptr[0] == '.' && ptr[1] == '.') 
                {
                    const char * slash = (char *) memrchr(res, '/', res_len);
                    if (slash != NULL) 
                    {
                        res_len = slash - res;
                    }
                    continue;
                }
                break;
            case 1:
                if (ptr[0] == '.') 
                {
                    continue;

                }
                break;
            case 0:
                continue;
        }
        res[res_len++] = '/';
        memcpy(&res[res_len], ptr, len);
        res_len += len;
    }

    if (res_len == 0) 
    {
        res[res_len++] = '/';
    }
    res[res_len] = '\0';
    
    if (bFree)
        src = strfree(src);
        
    return res;
}

/**
*   Конвертация представления пути из dos(C:\\path\\) в unix(C:/path/)
*/
char *dos_to_unix_path(char *src, BOOL bFree)
{
    //Подразумевается что в unix все папки в нижнем регистре
    return strlwr_lat(strreplace(src, "\\", "/", bFree));
}


/**
*   Поменять расширение в имени файла
*/
char *change_filename_ext(char *filename, const char *new_ext, BOOL bFree)
{
    char *ext_filename = strrchr(filename, '.');
    char *result_filename = NULL;
    
    if (ext_filename != NULL)
    {
        char *base = strleft(filename, ext_filename - filename);
        result_filename = dupprintf( "%s%s", base, new_ext);
        base = strfree(base);
    }
    else
        result_filename = dupprintf( "%s%s", filename, new_ext);
    
    if (bFree)
        filename = strfree(filename);
    return result_filename;
}

/**
*   Проверка на то же самый файл
*/
BOOL is_samefile(const char *filename1, const char *filename2)
{
    struct stat st1, st2;

    if (stat(filename1, &st1) != 0 || stat(filename2, &st2) != 0)
        return FALSE;

    if (st1.st_dev != st2.st_dev || st1.st_ino != st2.st_ino)
        return FALSE;
        
    return TRUE;
}

/**
*   Печать цветного текста
*/
void print_color_txt(unsigned int iColor, char *str, ...)
{
    va_list argptr;
    va_start(argptr, str);

    char msg[4096];
    char color[10];
    vsnprintf(msg, sizeof(msg), str, argptr);
    va_end(argptr);

    if (iColor ==  IC_RED_COLOR_TEXT)
        strcpy(color, "\x1b[31;1m");
    else if (iColor ==  IC_GREEN_COLOR_TEXT)
        strcpy(color, "\x1b[32m");
    else if (iColor ==  IC_YELLOW_COLOR_TEXT)
        strcpy(color, "\x1b[33m");
    else if (iColor ==  IC_BLUE_COLOR_TEXT)
        strcpy(color, "\x1b[34m");
    else if (iColor ==  IC_PURPLE_COLOR_TEXT)
        strcpy(color, "\x1b[35m");
    else if (iColor ==  IC_CYAN_COLOR_TEXT)
        strcpy(color, "\x1b[36m");
    else if (iColor ==  IC_WHITE_COLOR_TEXT)
        strcpy(color, "\x1b[37m");
    else if (iColor ==  IC_NORMAL_COLOR_TEXT)
        strcpy(color, "\x1b[0m");
    
    printf("%s %s\x1b[0m", color, msg);
}


unsigned int digit_count(int n)
{
    int i = 0;

    if (n < 0) 
    {
        i = 1;
        n = -n;
    }
    while (n > 0) 
    {
        n = n / 10;
        i += 1;
    }
    return i;    
}