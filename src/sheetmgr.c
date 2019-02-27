/**
*   Модуль функций менеджера листа ODS файла
*   @file
*/

#include "econvert.h"


/**
*   Преобразование ширины из ширины относительно символа,
*   используемого в формате XLS в пользовательскую ширину
*   Разъяснение по ширине:
*   https://support.microsoft.com/en-us/help/214123/description-of-how-column-widths-are-determined-in-excel
*   Функция писалась методом апроксимации к желаемому результату
*/
static double _convert_width_cm(double width)
{
    double result_width = 0.0;
    
    if (width != 0.0)
    {
        double cm = width * XLS_WIDTH_CORRECT_COEF;
        // Округление до 2-x десятичных знаков
        cm = round(cm * 100.0) / 100.0; 
        
        if (cm > 1.0)
        {
            // Конвертация больших ширин
            //double tg = 3.52 / 18.0;
            //double b = 3.86 - tg * 19.0;
            //result_width = (cm - b) / tg;
            double coeff = 2.26 / 8.43;
            result_width = cm / coeff * 1.4;
            //result_width = (cm - 1.08) / 4.08 * 21.0 + 5.0;
        }
        else if (cm >= 0.5)
        {
            // Конвертация маленьких ширин
            result_width = cm * 4.2;
        }
        else
        {
            // Конвертация сверх-маленьких ширин
            result_width = cm * 3.2; //2.95;
        }
    
        if (DBG_MODE) logAddLine("Convert Column WIDTH <%f> in CM <%f> result <%f>", width, cm, result_width);
    }
    return result_width;
}

/**
*   Определить стиль левой границы
*/
static unsigned int _xls_read_left_border_style(unsigned long line_style)
{
     return (line_style & 0x0000000F);
}

static unsigned int _xls_read_right_border_style(unsigned long line_style)
{
     return ((line_style & 0x000000F0) >> 4);
}

static unsigned int _xls_read_top_border_style(unsigned long line_style)
{
     return ((line_style & 0x00000F00) >> 8);
}

static unsigned int _xls_read_bottom_border_style(unsigned long line_style)
{
     return ((line_style & 0x0000F000) >> 12);
}


static unsigned int _xls_read_horizontal_align(BYTE align)
{
    return (align & 0x07);
}

static unsigned int _xls_read_vertical_align(BYTE align)
{
    return ((align & 0x70) >> 4);
}

static BOOL _xls_read_font_bold(unsigned int weight)
{
    return (weight == 0x2BC) ? TRUE : FALSE;
}

static BOOL _xls_read_font_italic(unsigned int flags)
{
    return ((flags & 0x0002) >> 1);
}

static BOOL _xls_read_font_strikeout(unsigned int flags)
{
    return ((flags & 0x0008) >> 3);
}

static lxw_format *_xlsx_clone_format(lxw_workbook *oBook, lxw_format *format)
{
    if (format != NULL)
    {
        lxw_format *new_format = workbook_add_format(oBook);
        // ВНИМАНИЕ! Нельзя переноcить данные обычным копированием памяти:
        // Например: memcpy(new_format, format, sizeof(lxw_format));
        // Иначе возникает ошибка 
        // *** Error in `./nixconvert': double free or corruption (!prev): 0x000000000481f5a0 ***
        // при закрытии workbook'а
        // Поэтому переносим все данные штатными функциями
        format_set_font_name(new_format, format->font_name);
        format_set_font_size(new_format, format->font_size);
        format_set_font_color(new_format, format->font_color);
        if (format->bold) format_set_bold(new_format);
        if (format->italic) format_set_italic(new_format);
        format_set_underline(new_format, format->underline);
        format_set_num_format(new_format, format->num_format);
        if (format->hidden) format_set_hidden(new_format);
        format_set_align(new_format, format->text_h_align);
        if (format->text_wrap) format_set_text_wrap(new_format);
        format_set_bg_color(new_format, format->bg_color);
        format_set_fg_color(new_format, format->fg_color);

        format_set_bottom(new_format, format->bottom);
        format_set_left(new_format, format->left);
        format_set_right(new_format, format->right);
        format_set_top(new_format, format->top);        
        
        return new_format;
    }
    return NULL;
}

static BOOL _xlsx_set_pagebreaks(lxw_worksheet *oSheet, xls::xlsWorkSheet *oXLSSheet)
{
    xls::BOF tmp;
    BYTE *buff;

    xls::ole2_seek(oXLSSheet->workbook->olestr, oXLSSheet->filepos);
    do
    {
		size_t read;
        read = xls::ole2_read(&tmp, 1, 4, oXLSSheet->workbook->olestr);
		if (read != 4)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        buff = (BYTE *) malloc(tmp.size);
        read = xls::ole2_read(buff, 1, tmp.size, oXLSSheet->workbook->olestr);
		if (read != tmp.size)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        switch (tmp.id)
        {
            case 0x1B:     //HORIZONTALPAGEBREAKS
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                unsigned int count = (buff[1] << 8) | buff[0];
                lxw_row_t *breaks = (lxw_row_t *) calloc(count + 1, sizeof(lxw_row_t)); 
                unsigned int i = 0;
    
                for (i=0; i < count; i++)
                {
                    breaks[i] = (buff[3+i*6] << 8) | buff[2+i*6];
                    if (DBG_MODE) logAddLine("Page break <%d> [%d]", i, breaks[i]);
                }
                breaks[i] = 0;
                
                if (count)
                    worksheet_set_h_pagebreaks(oSheet, breaks);
                break;
        }
        //if (DBG_MODE) logAddLine("Free [%x]", buff);
        memfree(buff);
    }
    while ((!oXLSSheet->workbook->olestr->eof) && (tmp.id != 0x0A));
    return TRUE;
}

static BOOL _xlsx_set_pagesetup(lxw_worksheet *oSheet, xls::xlsWorkSheet *oXLSSheet)
{
    xls::BOF tmp;
    BYTE *buff;

    unsigned int paper_type; 
    unsigned int scale; 
    unsigned int fit_width; 
    unsigned int fit_height; 
    unsigned int flags; 
    BOOL fit_flag=FALSE;
    BOOL orientation;

    xls::ole2_seek(oXLSSheet->workbook->olestr, oXLSSheet->filepos);
    do
    {
		size_t read;
        read = xls::ole2_read(&tmp, 1, 4, oXLSSheet->workbook->olestr);
		if (read != 4)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        buff = (BYTE *) malloc(tmp.size);
        read = xls::ole2_read(buff, 1, tmp.size, oXLSSheet->workbook->olestr);
		if (read != tmp.size)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        switch (tmp.id)
        {
                
            case 0xA1:     //PAGESETUP
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                paper_type = (buff[1] << 8) | buff[0]; 
                scale = (buff[3] << 8) | buff[2]; 
                fit_width = (buff[7] << 8) | buff[6]; 
                fit_height = (buff[9] << 8) | buff[8]; 
                flags = (buff[11] << 8) | buff[10]; 
                orientation = (flags & 0x0002) >> 1;
                
                if (DBG_MODE) logAddLine("Scale <%d> \tFit <%d> [%d : %d] ", scale, fit_flag, fit_width, fit_height);
                

                if (fit_flag)
                    worksheet_fit_to_pages(oSheet, fit_width, fit_height);
                else
                    worksheet_set_print_scale(oSheet, scale);
                    
                worksheet_set_paper(oSheet, paper_type); 
                if (!orientation)
                    worksheet_set_landscape(oSheet);
                    
                break;
                
            case 0x81:     //SHEETPR
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                fit_flag = (((buff[1] << 8) | buff[0]) & 0x0100) >> 8;
                if (DBG_MODE) logAddLine("Fit flag <%d>", fit_flag);
                break;                
        }
        //if (DBG_MODE) logAddLine("Free [%x]", buff);
        memfree(buff);
    }
    while ((!oXLSSheet->workbook->olestr->eof) && (tmp.id != 0x0A));
    return TRUE;
}

static BOOL _xlsx_set_margins(lxw_worksheet *oSheet, xls::xlsWorkSheet *oXLSSheet)
{
    xls::BOF tmp;
    BYTE *buff;

    double left_margin = -1; 
    double right_margin = -1; 
    double top_margin = -1; 
    double bottom_margin = -1; 

    xls::ole2_seek(oXLSSheet->workbook->olestr, oXLSSheet->filepos);
    do
    {
		size_t read;
        read = xls::ole2_read(&tmp, 1, 4, oXLSSheet->workbook->olestr);
		if (read != 4)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        buff = (BYTE *) malloc(tmp.size);
        read = xls::ole2_read(buff, 1, tmp.size, oXLSSheet->workbook->olestr);
		if (read != tmp.size)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        switch (tmp.id)
        {
            case 0x26:     //LEFTMARGIN
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                left_margin = bytes_to_double(buff);
                if (DBG_MODE) logAddLine("Left margin <%f>", left_margin);
                break;                
            case 0x27:     //RIGHTMARGIN
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                right_margin = bytes_to_double(buff);
                if (DBG_MODE) logAddLine("Right margin <%f>", right_margin);
                break;                
            case 0x28:     //TOPMARGIN
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                top_margin = bytes_to_double(buff);
                if (DBG_MODE) logAddLine("Top margin <%f>", top_margin);
                break;                
            case 0x29:     //BOTTOMMARGIN
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                bottom_margin = bytes_to_double(buff);
                if (DBG_MODE) logAddLine("Bottom margin <%f>", bottom_margin);
                break;                
        }
        //if (DBG_MODE) logAddLine("Free [%x]", buff);
        memfree(buff);
        //if (DBG_MODE) logAddLine("OK");
    }
    while ((!oXLSSheet->workbook->olestr->eof) && (tmp.id != 0x0A));
    
    worksheet_set_margins(oSheet, left_margin, right_margin, top_margin, bottom_margin);
    return TRUE;
}

static double _xls_get_standardwidth(xls::xlsWorkSheet *oXLSSheet)
{
    xls::BOF tmp;
    BYTE *buff;
    unsigned int standardwidth = 0;

    xls::ole2_seek(oXLSSheet->workbook->olestr, oXLSSheet->filepos);
    do
    {
		size_t read;
        read = xls::ole2_read(&tmp, 1, 4, oXLSSheet->workbook->olestr);
		if (read != 4)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        buff = (BYTE *) malloc(tmp.size);
        read = xls::ole2_read(buff, 1, tmp.size, oXLSSheet->workbook->olestr);
		if (read != tmp.size)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        switch (tmp.id)
        {
            case 0x99:     //STANDARDWIDTH
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                standardwidth = (buff[1] << 8) | buff[0];
                if (DBG_MODE) logAddLine("Standard Width <%d>", standardwidth);
                break;                
        }
        //if (DBG_MODE) logAddLine("Free [%x]", buff);
        memfree(buff);
    }
    while ((!oXLSSheet->workbook->olestr->eof) && (tmp.id != 0x0A));
    return ((double)standardwidth);
}

static double _xls_get_defcolwidth(xls::xlsWorkSheet *oXLSSheet)
{
    xls::BOF tmp;
    BYTE *buff;
    double width = 0.0;

    xls::ole2_seek(oXLSSheet->workbook->olestr, oXLSSheet->filepos);
    do
    {
		size_t read;
        read = xls::ole2_read(&tmp, 1, 4, oXLSSheet->workbook->olestr);
		if (read != 4)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        buff = (BYTE *) malloc(tmp.size);
        read = xls::ole2_read(buff, 1, tmp.size, oXLSSheet->workbook->olestr);
		if (read != tmp.size)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        switch (tmp.id)
        {
            case 0x55:     //DEFCOLWIDTH
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                
                unsigned int defcolwidth = (buff[1] << 8) | buff[0];
                if ((defcolwidth != 8) && (defcolwidth != 0))
                    width = defcolwidth * XLS_ARIAL_SIZE_COEF;
                if (DBG_MODE) logAddLine("Def col [%d] Width <%f>", defcolwidth, width);
                break;                
        }
        //if (DBG_MODE) logAddLine("Free [%x]", buff);
        memfree(buff);
    }
    while ((!oXLSSheet->workbook->olestr->eof) && (tmp.id != 0x0A));
    return width;
}

static double _xls_get_default_col_width(xls::xlsWorkSheet *oXLSSheet)
{
    double standardwidth = _xls_get_standardwidth(oXLSSheet);
    return standardwidth ? standardwidth : _xls_get_defcolwidth(oXLSSheet);
}

static double _xls_show_default_style(xls::xlsWorkSheet *oXLSSheet)
{
    xls::BOF tmp;
    BYTE *buff;

    xls::ole2_seek(oXLSSheet->workbook->olestr, oXLSSheet->filepos);
    do
    {
		size_t read;
        read = xls::ole2_read(&tmp, 1, 4, oXLSSheet->workbook->olestr);
		if (read != 4)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        buff = (BYTE *) malloc(tmp.size);
        read = xls::ole2_read(buff, 1, tmp.size, oXLSSheet->workbook->olestr);
		if (read != tmp.size)
        {
            if (DBG_MODE) errBox("Error in <xls::ole2_read> function");
            return FALSE;
        }
        
        switch (tmp.id)
        {
            case 0x293:     //STYLE
                if (DBG_MODE) xls::xls_showBOF(&tmp);
                if (DBG_MODE) errBox("Default style %x %x", buff[1], buff[0]);
                break;                
        }
        //if (DBG_MODE) logAddLine("Free [%x]", buff);
        memfree(buff);
    }
    while ((!oXLSSheet->workbook->olestr->eof) && (tmp.id != 0x0A));
}

/**
*   Установка выравнивания в формат ячейки
*/
static lxw_format *_set_format_align(lxw_format *format, unsigned int align)
{
    //if (DBG_MODE) logAddLine("Align <%X>", align);
    if ((align & 0x0007) == 0x01)
    {
        //if (DBG_MODE) logAddLine("Align HORIZONTAL LEFT");
        format_set_align(format, LXW_ALIGN_LEFT);
    }
    if ((align & 0x0007) == 0x02)
    {
        //if (DBG_MODE) logAddLine("Align HORIZONTAL CENTER");
        format_set_align(format, LXW_ALIGN_CENTER);
    }
    if((align & 0x0007) == 0x03)
    {
        //if (DBG_MODE) logAddLine("Align HORIZONTAL RIGHT");
        format_set_align(format, LXW_ALIGN_RIGHT);
    }
    if((align & 0x0007) == 0x04)
    {
        //if (DBG_MODE) logAddLine("Align HORIZONTAL FILL");
        format_set_align(format, LXW_ALIGN_FILL);
    }
    if((align & 0x0007) == 0x05)
    {
        //if (DBG_MODE) logAddLine("Align HORIZONTAL JUSTIFY");
        format_set_align(format, LXW_ALIGN_JUSTIFY);
    }
    if((align & 0x0070) == 0x00)
    {
        //if (DBG_MODE) logAddLine("Align VERTICAL TOP");
        format_set_align(format, LXW_ALIGN_VERTICAL_TOP);
    }
    if((align & 0x0070) == 0x10)
    {
        //if (DBG_MODE) logAddLine("Align VERTICAL CENTER");
        format_set_align(format, LXW_ALIGN_VERTICAL_CENTER);
    }
    if((align & 0x0070) == 0x20)
    {
        //if (DBG_MODE) logAddLine("Align VERTICAL BOTTOM");
        format_set_align(format, LXW_ALIGN_VERTICAL_BOTTOM);
    }
    if((align & 0x0070) == 0x40)
    {
        //if (DBG_MODE) logAddLine("Align VERTICAL JUSTIFY");
        format_set_align(format, LXW_ALIGN_VERTICAL_JUSTIFY);
    }
    if(align & 0x0008)
    {
        //if (DBG_MODE) logAddLine("Align TEXT WRAP");
        format_set_text_wrap(format);
    }
    return format;
}

static lxw_format *_set_format_borders(lxw_format *format, unsigned int linestyle)
{
    unsigned int style = _xls_read_left_border_style(linestyle);
    if (style)
        format_set_left(format, style);
    style = _xls_read_right_border_style(linestyle);
    if (style)
        format_set_right(format, style);
    style = _xls_read_top_border_style(linestyle);
    if (style)
        format_set_top(format, style);
    style = _xls_read_bottom_border_style(linestyle);
    if (style)
        format_set_bottom(format, style);
    return format;    
}

static lxw_format *_set_format_default_font(lxw_format *format)
{
    //Выставить шрифт по умолчанию
    //if (DBG_MODE) logAddLine("Set default font <Arial>");
    format_set_font_name(format, "Arial");
    format_set_font_size(format, 10);
}

static lxw_format *_set_format_font(lxw_format *format, unsigned int font_index, xls::xlsWorkBook *oBook)
{
    unsigned int i_font = font_index - 1;
    if (i_font < oBook->fonts.count)
    {
        char *font_name = (char *) oBook->fonts.font[i_font].name;
        format_set_font_name(format, font_name);
        double font_size = (double) oBook->fonts.font[i_font].height / XLS_FONT_SIZE_COEF;
        format_set_font_size(format, font_size);
        BOOL font_bold = _xls_read_font_bold(oBook->fonts.font[i_font].bold);
        if (font_bold)
            format_set_bold(format);
        unsigned int font_italic = _xls_read_font_italic(oBook->fonts.font[i_font].flag);
        if (font_italic)
            format_set_italic(format);
        format_set_underline(format, oBook->fonts.font[i_font].underline);
        unsigned int font_strikeout = _xls_read_font_strikeout(oBook->fonts.font[i_font].flag);
        if (font_strikeout)
            format_set_font_strikeout(format);
        //if (DBG_MODE) logAddLine("Font <%s> size: <%f> flags: <%X> bold: <%d> italic: <%d> underline: <%X>", font_name, font_size, oBook->fonts.font[i_font].flag, font_bold, font_italic, oBook->fonts.font[i_font].underline);
    }
    else
        _set_format_default_font(format);
    return format;    
}

/**
*   Установить формат чисел
*   nformat_index - это индекс формата 
*       (0-163) - системный зарезервированный индекс
*       (164-..) - пользвательский индекс
*/
static lxw_format *_set_format_num_format(lxw_format *format, unsigned int nformat_index, xls::xlsWorkBook *oBook)
{
    if (nformat_index < 163)
    {
        // if (DBG_MODE) logAddLine("Num format: <%d>", nformat_index);
        format_set_num_format_index(format, nformat_index);        
    }
    else
    {
        char *format_value = (char *)oBook->formats.format[nformat_index-164].value;
        // if (DBG_MODE) logAddLine("Num format: <%s>", format_value);
        format_set_num_format(format, format_value);
    }
    return format;
}

static BOOL _xls_is_row_hidden(xls::xlsRow *row)
{
    unsigned long flags = row->flags;
    // unsigned int height = (unsigned int)row->height;
    // return ((flags & 0x00000020) >> 1) || (height == 0);
    return ((flags & 0x00000020) >> 1);
}

static BOOL _xls_is_col_hidden(xls::st_colinfo::st_colinfo_data *colinfo_data)
{
    unsigned long flags = colinfo_data->flags;
    unsigned int width = (unsigned int)colinfo_data->width;
    return (flags == 0x0001) || (width == 0);
}

static unsigned int _xls_read_background_color(unsigned int color_idx)
{
    //if (DBG_MODE) logAddLine("Color: <%x>", xls::xls_getColor((color_idx & 0x007F), 1));
    return xls::xls_getColor((color_idx & 0x007F), 1);
}

/**
*   Проверка является ли режим выравнивания - выравниванием по умолчанию
*/
static BOOL _xls_is_default_align(xls::xlsWorkBook *oXLSBook, xls::xlsCell *oXLSCell)
{
    return (oXLSBook->xfs.xf[oXLSCell->xf].align == 0x20);
}

static BOOL _xls_is_default_font(xls::xlsWorkBook *oXLSBook, xls::xlsCell *oXLSCell)
{
    return FALSE;
}

static BOOL _xls_is_default_borders(xls::xlsWorkBook *oXLSBook, xls::xlsCell *oXLSCell)
{
    // if (DBG_MODE) logAddLine("Borders: <%x>", oXLSBook->xfs.xf[oXLSCell->xf].linestyle);
    return ((oXLSBook->xfs.xf[oXLSCell->xf].linestyle & 0xFFFF) == 0);
}

static BOOL _xls_is_default_numformat(xls::xlsWorkBook *oXLSBook, xls::xlsCell *oXLSCell)
{
    // Ето индекс формата 
    // (0-163) - системный зарезервированный индекс
    // (164-..) - пользвательский индекс
    WORD num_format = oXLSBook->xfs.xf[oXLSCell->xf].format;
    BYTE *format_value = NULL;
    if (num_format > 163)
        format_value = oXLSBook->formats.format[num_format-164].value;
        // if (DBG_MODE) logAddLine("Is DEFAULT XLS NumFormat: <%d> <%s>", num_format, format_value);
        return (num_format == 164);
    return (num_format == 0);
}

static BOOL _xls_is_default_rotation(xls::xlsWorkBook *oXLSBook, xls::xlsCell *oXLSCell)
{
    return (oXLSBook->xfs.xf[oXLSCell->xf].rotation == 0);
}

static BOOL _xls_is_default_bgcolor(xls::xlsWorkBook *oXLSBook, xls::xlsCell *oXLSCell)
{
    return (_xls_read_background_color(oXLSBook->xfs.xf[oXLSCell->xf].groundcolor) == 0xFFFFFF);
}

/**
*   Открыть xls шаблон и подготовить его для заполнения
*/
lxw_workbook *open_workbook_xls(char *xls_filename)
{
    char *xlsx_filename = change_filename_ext(xls_filename, ".xlsx");
    lxw_workbook *xlsx_book = NULL;
    
    if (DBG_MODE) logAddLine("lib XLS: <%s>", xls::xls_getVersion());
    xls::xlsWorkBook *xls_book = xls::xls_open(xls_filename, "UTF-8");    
	if (!xls_book) 
    {
		if (DBG_MODE) logAddLine("File <%s> not found!", xls_filename);
        xlsx_filename = strfree(xlsx_filename);
		return NULL;
	}
    xls_parseWorkBook(xls_book);
    
    xlsx_book = new_workbook(xlsx_filename);
    // Перебор по листам
    for (unsigned int i_sheet=0; i_sheet < xls_book->sheets.count; i_sheet++)
    {
        if (DBG_MODE) logAddLine("Sheet: <%s>", xls_book->sheets.sheet[i_sheet].name);
        lxw_worksheet *xlsx_sheet = workbook_add_worksheet(xlsx_book, (char *) xls_book->sheets.sheet[i_sheet].name);
        
        xls::xlsWorkSheet *xls_sheet = xls::xls_getWorkSheet(xls_book, i_sheet);
		xls::xls_parseWorkSheet(xls_sheet);
        
        unsigned int i_max_col = 0;        
        double default_col_width = _convert_width_cm(_xls_get_default_col_width(xls_sheet));
        
        // Установить колонки
        for (unsigned int i_colinfo=0; i_colinfo < xls_sheet->colinfo.count; i_colinfo++)
        {
            //Выставить колонки
            double col_width = _convert_width_cm((double)xls_sheet->colinfo.col[i_colinfo].width);
            unsigned int first_col = xls_sheet->colinfo.col[i_colinfo].first;
            unsigned int last_col = xls_sheet->colinfo.col[i_colinfo].last;
            if (DBG_MODE) logAddLine("Column [%d] width: <%f> first <%d> last <%d>", i_colinfo, col_width, first_col, last_col);
            if (_xls_is_col_hidden(&xls_sheet->colinfo.col[i_colinfo]))
            {
                if (DBG_MODE) logAddLine("Col [%d] is hidden Flags <%x>", i_colinfo, xls_sheet->colinfo.col[i_colinfo].flags);
                lxw_row_col_options col_options = {.hidden = 1, .level = 0, .collapsed = 0};
                worksheet_set_column_opt(xlsx_sheet, first_col, last_col, col_width, NULL, &col_options);
            }
            else
                worksheet_set_column(xlsx_sheet, first_col, last_col, col_width, NULL);
            i_max_col = last_col;
        }
            
        // Перебор по строкам
        for (unsigned int i_row=0; i_row <= xls_sheet->rows.lastrow; i_row++)
        {
			xls::xlsRow *xls_row = xls::xls_row(xls_sheet, i_row);
            if (_xls_is_row_hidden(xls_row))
            {
                if (DBG_MODE) logAddLine("Row [%d] is hidden", i_row);
                lxw_row_col_options row_options = {.hidden = 1, .level = 0, .collapsed = 0};
                worksheet_set_row_opt(xlsx_sheet, i_row, (double)xls_row->height / XLS_HEIGHT_COEF, NULL, &row_options);
            }
            else
            {
                double row_height = ((double)xls_row->height) / XLS_HEIGHT_COEF;
                //if (DBG_MODE) logAddLine("Row [%d] height <%f>", i_row, row_height);
                if (!row_height)
                    row_height = LXW_DEF_ROW_HEIGHT;
                worksheet_set_row(xlsx_sheet, i_row, row_height, NULL);
            }
            
            
            for (unsigned int i_col=0; i_col <= xls_sheet->rows.lastcol; i_col++)
            {
				xls::xlsCell *xls_cell = xls::xls_cell(xls_sheet, i_row, i_col);

				if (!xls_cell)
                { 
					continue;
                }
                
                if (i_col > i_max_col)
                {
                    if (DBG_MODE) logAddLine("Column [%d] default width: <%f> Max col: %d", i_col, default_col_width, i_max_col);
                    worksheet_set_column(xlsx_sheet, i_col, i_col, default_col_width, NULL);
                    i_max_col = i_col;
                }
                
                // if (DBG_MODE) logAddLine("Cell [%d : %d]", i_row, i_col);
                // Установка формата ячейки
                lxw_format *format = workbook_add_format(xlsx_book);
                if (!_xls_is_default_align(xls_book, xls_cell))
                    _set_format_align(format, xls_book->xfs.xf[xls_cell->xf].align);
                if (xls_book->xfs.xf[xls_cell->xf].type != 0xFFF5)
                    _set_format_borders(format, xls_book->xfs.xf[xls_cell->xf].linestyle);                
                if (!_xls_is_default_font(xls_book, xls_cell))
                    _set_format_font(format, xls_book->xfs.xf[xls_cell->xf].font, xls_book);
                if (!_xls_is_default_numformat(xls_book, xls_cell))
                    _set_format_num_format(format, xls_book->xfs.xf[xls_cell->xf].format, xls_book);
                if (!_xls_is_default_rotation(xls_book, xls_cell))
                    format_set_rotation(format, xls_book->xfs.xf[xls_cell->xf].rotation);
                if (!_xls_is_default_bgcolor(xls_book, xls_cell))
                    format_set_bg_color(format, _xls_read_background_color(xls_book->xfs.xf[xls_cell->xf].groundcolor));

                // ЗАполнение ячейки данными                
                if (xls_cell->colspan || xls_cell->rowspan)
                {
                    worksheet_merge_range(xlsx_sheet, i_row, i_col, 
                                          i_row + xls_cell->rowspan - 1,
                                          i_col + xls_cell->colspan - 1,
                                          (char *) xls_cell->str, format);
                    if (isnumeric((char *) xls_cell->str))
                        worksheet_write_number(xlsx_sheet, i_row, i_col, atof((char *) xls_cell->str), format);
                    
                }
                else if (isnumeric((char *) xls_cell->str))
                    worksheet_write_number(xlsx_sheet, i_row, i_col, xls_cell->d, format);
                else if (!strempty((char *) xls_cell->str))
                    worksheet_write_string(xlsx_sheet, i_row, i_col, (char *) xls_cell->str, format);
                else
                    worksheet_write_string(xlsx_sheet, i_row, i_col, NULL, format);
            }
        }
        // Установить разрывы страниц
        _xlsx_set_pagebreaks(xlsx_sheet, xls_sheet);
        // Установить параметры страницы
        _xlsx_set_pagesetup(xlsx_sheet, xls_sheet);
        // Установить поля страницы
        _xlsx_set_margins(xlsx_sheet, xls_sheet);
        if (DBG_MODE) logAddLine("Xls close worksheet <%s>", xls_book->sheets.sheet[i_sheet].name);
        xls::xls_close_WS(xls_sheet); 
        if (DBG_MODE) logAddLine("...ok");        
    }

    //if (DBG_MODE) logAddLine("Xls close workbook");
    xls::xls_close(xls_book);
    if (DBG_MODE) logAddLine("Open workbook XLS <%s> is OK. Workbook <%s> is active", xls_filename, xlsx_filename);
    
    return xlsx_book;
}

/**
*   Закрыть xlsx файл.
* Ошибка. При закрытии XLSX книги выдается:
*------------------------------------------------------------------------------
*** Error in `./nixconvert': double free or corruption (!prev): 0x0991fc70 ***
*------------------------------------------------------------------------------
* Возможно проблема в шаблоне.
* Необходимо открыть шаблон в LibreOffice и сохранить перезаписав.
* Удалить форматирование всех ячеек.
*/
BOOL close_workbook_xlsx(lxw_workbook *oBook)
{
    if (DBG_MODE) logAddLine("lib XLSX Writer: <%s>", LXW_VERSION);
    
    if (oBook != NULL)
    {
        // Закрыть книгу
        char *book_filename = (char *) oBook->filename;
        if (DBG_MODE) logAddLine("Close workbook <%s> ...", book_filename);
        workbook_close(oBook);
        
        if (DBG_MODE) logAddLine("...OK");
        return TRUE;
    }
    return FALSE;
}


static lxw_cell *_get_cell(lxw_worksheet *oSheet, unsigned int row, unsigned int col)
{  
    lxw_cell *current_cell = NULL;
    
    lxw_row *current_row = lxw_worksheet_find_row(oSheet, (lxw_row_t) row);
    current_cell = lxw_worksheet_find_cell(current_row, (lxw_col_t) col);
    
    return current_cell;
}


static char *_get_sst_string(lxw_worksheet *oSheet, int iStringID)
{
    sst_element *first_sst = STAILQ_FIRST(oSheet->sst->order_list);
    sst_element *current = NULL;

    //Пусто
    if (STAILQ_EMPTY(oSheet->sst->order_list))
        return NULL;
        
    if (first_sst && (first_sst->index == iStringID)) 
        current = first_sst;
    if (!current)
        STAILQ_FOREACH(current, oSheet->sst->order_list, sst_order_pointers)
            if (current->index == iStringID) 
                break;
    if (current)
        return strcopy(current->string);

    return NULL;
}


static char *_get_cell_value(lxw_worksheet *oSheet, unsigned int row, unsigned int col)
{
    lxw_cell *cell = _get_cell(oSheet, row, col);
    
    if (cell)
        if (cell->type == NUMBER_CELL)
        {
            return dupprintf("%f", cell->u.number);
        }
        else if (cell->type == STRING_CELL)
            return _get_sst_string(oSheet, cell->u.string_id);
        else if (cell->type == INLINE_STRING_CELL)
            return strcopy(cell->u.string);
        else if (cell->type == FORMULA_CELL)
            return strcopy(cell->u.string);
    return NULL;
}


static lxw_format *_get_cell_format(lxw_worksheet *oSheet, unsigned int i_row, unsigned int i_col)
{
    lxw_cell *cell = _get_cell(oSheet, i_row, i_col);
    lxw_format *format = (cell) ? cell->format : NULL;
    //if (DBG_MODE) logAddLine("Get cell [%d : %d] format <%s>", i_row, i_col, format);
    return format;
}


lxw_format *get_cell_format(lxw_worksheet *oSheet, unsigned int i_row, unsigned int i_col)
{
    return _get_cell_format(oSheet, i_row, i_col);
}


/**
*   Клонировать формат ячейки
*/
lxw_format *clone_cell_format(lxw_workbook *oBook, lxw_worksheet *oSheet, unsigned int i_row, unsigned int i_col)
{
    lxw_format *format = _get_cell_format(oSheet, i_row, i_col);
    lxw_format *new_format = NULL;
    
    if (format != NULL)
        new_format = _xlsx_clone_format(oBook, format);
    else
        if (DBG_MODE) logWarning("Null cloned cell format");
    return new_format;
}

/**
*   Клонировать формат ячейки без обрамления и интерьера
*   Т.е. только формат данных
*/
lxw_format *clone_cell_data_format(lxw_workbook *oBook, lxw_worksheet *oSheet, unsigned int i_row, unsigned int i_col)
{
    lxw_format *format = _get_cell_format(oSheet, i_row, i_col);
    lxw_format *new_format = NULL;
    if (format != NULL)
    {
        new_format = workbook_add_format(oBook);
        format_set_num_format(new_format, format->num_format);
    }
    return new_format;
}


BOOL set_cell_value(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    unsigned int row, unsigned int col, char *value,
                    lxw_format *format)
{
    // ВНИМАНИЕ! Корректность заполнения формата ячейки
    if (format == NULL)
        format = _get_cell_format(oSheet, row, col);
 
    BOOL is_num = isnumeric(value);
    //if (DBG_MODE) logAddLine("Set cell [%d : %d] value <%s> [%d]", row, col, value, is_num);
    
    if (is_num)
    {
        // Обработка числовых значений
        double f_value = atof(value);
        if (DBG_MODE) logAddLine("Set cell [%d : %d] numeric value <%f> string value (%s)", row, col, f_value, value);

        if (row >= MAX_ROW_COUNT) 
        {
            if (DBG_MODE) logWarning("Maximum row count [%d]", row);
            return FALSE;
        }
        if (col >= MAX_COLUMN_COUNT) 
        {
            if (DBG_MODE) logWarning("Maximum column count [%d]", col);
            return FALSE;
        }
        
        if (format == NULL)
            // Формат у числа обязан быть!
            format = workbook_add_format(oBook);
        if (strempty(format->num_format))
        {
            // Формат устанавливать только если он не определен
            // т.к. он может задаваться для ячейки принудительно
            int decimal = decimal_point(value);
            if (decimal > 0)
            {    
                // Если число с десятичной точкой, то установить формат "0.00"
                char *str_decimal = strgen('0', decimal);
                char *str_format = strconcatenate("0.", str_decimal);
                //if (DBG_MODE) logAddLine("Value <%s>\tnum: <%f>\tFormat: [%s]\tCell num format: [%s]", value, f_value, str_format, format->num_format);
                format_set_num_format(format, str_format);
                str_decimal = strfree(str_decimal);        
                //str_format = strfree(str_format);
            }
            else
                // Если это целое число, то формат целого числа
                format_set_num_format(format, "0");
        }
        worksheet_write_number(oSheet, row, col, f_value, format);
    }
    else
        worksheet_write_string(oSheet, row, col, value, format);    
    return TRUE;
}

/**
*	Функция по имени файла вычисляет название книги.
*	c:\prt\example.xls -> example.xls
*/
char *getWorkBookName(char *xlsName)
{
	// вычленяем название книги
	char *s1 = strcopy(xlsName);
	int size = strlen(xlsName);
    s1 = strreverse(s1);
	int num = strfind(s1, "/");
    s1 = strfree(s1);

	if (num >= 0)			
		s1 = strright_pos(xlsName, size - num);
	else
		s1 = strcopy(xlsName);

	return s1;
}

/**
*   Очистить лист от всех значений
*/
static void _clean_range(lxw_worksheet *oSheet, unsigned int row1, unsigned int col1, unsigned int row2, unsigned col2)
{
    lxw_format *format = NULL;

    for (unsigned int i_row=row1; i_row < row2; i_row++)
        for (unsigned int i_col=col1; i_col < col2; i_col++)
        {
            worksheet_write_string(oSheet, i_row, i_col, "", NULL);
        }    
}

void cleanSheet(lxw_worksheet *oSheet)
{
    return _clean_range(oSheet, oSheet->dim_rowmin, oSheet->dim_colmin,
                        oSheet->dim_rowmax, oSheet->dim_colmax);
}

/**
*   Очистить область ячеек от всех значений
*/
void cleanRange(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    unsigned int row1 = lxw_name_to_row(sAddressBegin);
    unsigned int col1 = lxw_name_to_col(sAddressBegin);
    unsigned int row2 = lxw_name_to_row(sAddressEnd);
    unsigned int col2 = lxw_name_to_col(sAddressEnd);
    return _clean_range(oSheet, row1, col1, row2, col2);
}

/**
*   Заполнить область ячеек массивом формул
*/
void setRangeValueArray(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                        char *sAddressBegin, char *sAddressEnd, icStringCollection *lValues)
{
    if (strempty(sAddressBegin))
    {
        if (DBG_MODE) errBox("Not define address range [%s : %s]", sAddressBegin, sAddressEnd);
    }
    else if (strempty(sAddressEnd) || strequal(sAddressBegin, sAddressEnd))
    {
        //Диапазон - это 1 ячейка
        char *value = lValues->GetValue();
        setCellValue(oBook, oSheet, sAddressBegin, value);
        value = strfree(value);
    }
    else
    {
        char *value = lValues->GetValue();
        
        unsigned int row1 = lxw_name_to_row(sAddressBegin);
        unsigned int col1 = lxw_name_to_col(sAddressBegin);
        unsigned int row2 = lxw_name_to_row(sAddressEnd);
        unsigned int col2 = lxw_name_to_col(sAddressEnd);

        //if (DBG_MODE) logAddLine("Set range values [%d, %d] : [%d, %d]", row1, col1, row2, col2);
        
        for (unsigned int i_row=row1; i_row <= row2; i_row++)
            for (unsigned int i_col=col1; i_col <= col2; i_col++)
            {
                if (value)            
                {
                    BOOL result = set_cell_value(oBook, oSheet, i_row, i_col, value);
                    if (!result)
                        continue;
                }
                if (lValues->GetNext())
                {
                    lValues = lValues->GetNext();
                    value = strinit(value, lValues->GetValue());
                }
                else
                {
                    value = strfree(value);
                    return;
                }
            }
        value = strfree(value);
    }
}

icStringCollection *getRangeValueArray(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    icStringCollection *collection = NULL;
    char *value = NULL;
    
    if (strempty(sAddressBegin))
        if (DBG_MODE) errBox("Not define address range [% : %s]", sAddressBegin, sAddressEnd);
    else if (strempty(sAddressEnd) || strequal(sAddressBegin, sAddressEnd))
    {
        //Диапазон - это 1 ячейка
        collection = new icStringCollection();
        value = getCellValue(oSheet, sAddressBegin);
        if (value)
            collection->AppendValue(value);
        else
            collection->AppendValue("");        
    }
    else
    {
        collection = new icStringCollection();
        unsigned int row1 = lxw_name_to_row(sAddressBegin);
        unsigned int col1 = lxw_name_to_col(sAddressBegin);
        unsigned int row2 = lxw_name_to_row(sAddressEnd);
        unsigned int col2 = lxw_name_to_col(sAddressEnd);
    
        if (DBG_MODE) logAddLine("Get range [%s : %s] values [%d, %d] : [%d, %d]", sAddressBegin, sAddressEnd, row1, col1, row2, col2);

        for (unsigned int i_row=row1; i_row <= row2; i_row++)
            for (unsigned int i_col=col1; i_col < col2; i_col++)
            {
                lxw_cell *cell = _get_cell(oSheet, i_row, i_col); 
                value = strcopy(cell->u.string);
                collection->AppendValue(value);
            }
    }
    return collection;
}

void setRangeColumnWidth(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, double iWidth)
{
    unsigned int col1 = lxw_name_to_col(sAddressBegin);
    unsigned int col2 = lxw_name_to_col(sAddressEnd);
    
    //if (DBG_MODE) logAddLine("Set range columns [%d : %d] width: <%d>", col1, col2, iWidth);

    worksheet_set_column(oSheet, col1, col2, iWidth, NULL);
}

/**
*	Открывает нужную книгу и лист
*	bAddList: Признак, создания нового листа, если в книге такого листа нет
*/
BOOL openBookAndList(lxw_workbook *oBook, char *FNameValue, char *ListValue,
                     BOOL bAddList)
{
    BOOL covTrue = TRUE;
    BOOL covFalse = FALSE;
    
    lxw_worksheet *oSheet = NULL;
    
	//	Если Excel файл существует, открываем его
	if (file_exists(FNameValue))
	{
		BOOL bFindBook = FALSE;
		
		// Ищем нужную книгу
		long n=1L;
        long bound = 1; 
		char *strN = NULL;
        char *s1 = NULL;
        char *s2 = NULL;

		for(n; n<=bound; n++)
		{
			// вычленяем название книги
			s1 = strcopy(FNameValue);
			s1 = strreverse(s1);
			int num = strfind(s1, "/");
			int size = strlen(FNameValue);
            s1 = strfree(s1);

			if (num>=0)			
				s1 = substr(FNameValue, size - num, strlen(FNameValue) - (size - num));
			else
				s1 = strcopy(FNameValue);

			s1 = strlwr_lat(s1);
			strN = strlwr_lat(strN);

			if (strequal(strN, s1))
			{
				bFindBook = TRUE;
				break;
			}

		}

		// Если книга не открыта, то откроем ее
        oSheet = get_worksheet_by_name(oBook, ListValue);
        
        if (!oSheet)
        {
			//	добавляем новый с таким именем если указан соответствующий признак
			//	bAddList
			if (!strequal(ListValue, "") && bAddList)
                oSheet = workbook_add_worksheet(oBook, ListValue);
        }
	}	
	else 
	{
		// Пытаемся создать файл с указаным именем
		char *ExcelDir;
		FNameValue = strreverse(FNameValue);
		int index = strfind(FNameValue, "/");
		FNameValue = strreverse(FNameValue);
		ExcelDir = strleft(FNameValue, strlen(FNameValue) - index-1);

		//	Проверяем существует ли дирректория
		if (dir_exists(ExcelDir))
		{
			// проверяем существует ли такой лист
            oSheet = get_worksheet_by_name(oBook, ListValue);
            
            if (!oSheet)
            {
                //	добавляем новый с таким именем
                oSheet = workbook_add_worksheet(oBook, ListValue);
            }
		}
		else
		{
			errBox("ERROR: Неверно указана директория : <%s>", FNameValue);
			return FALSE;
		}
	}

	return TRUE;
}

/**
*	Принудительн закрывает книгу
*/
BOOL ForceCloseWorkbook(lxw_workbook *oBook, char *xlsName)
{
	if (oBook)
	{
        close_workbook_xlsx(oBook);
        oBook = NULL;
		return TRUE;
	}

	return FALSE;
}


/**
*	Сохраняет текущую книгу
*/
void SaveActiveBook(lxw_workbook *oBook)
{
    ForceCloseWorkbook(oBook, NULL);
}

/**
*	Сохраняет текущую книгу и по необходимости закрывает приложение 
*/
void SaveAndQuit(lxw_workbook *oBook)
{
    SaveActiveBook(oBook);
}

/**
*	Удаляет неизмененные листы
*/
BOOL DelUnchangedList(char *bookName, icStringCollection *m_ChangedListName,
 					  lxw_workbook *oBook)
{
	icStringCollection selSheets = icStringCollection();
    icStringCollection delSheets = icStringCollection();
	
	selSheets.RemoveAll();
	delSheets.RemoveAll();

	//	Открываем книгу
	char *ListName = "";

	if (!file_exists(bookName))
		return FALSE;

	//	Получаем колекцию листов

	//	Удаляем не выделенные листы

	return TRUE;
}

/**
*   Удалить листы из книги
*/
BOOL deleteSheets(lxw_workbook *oBook, icStringCollection *SheetNames)
{
    return TRUE;
}

/**
*   Удалить колонки заданной области ячеек
*/
BOOL deleteRangeColumns(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    return TRUE;
}

/**
*   Переместить листы в укзанную книгу на указанный лист
*/
BOOL moveSheet(icStringCollection *SheetNames, char *xlsName, char *sSheetName)
{
    return TRUE;
}

/**
*   Копировать листы в укзанную книгу на указанный лист
*/
BOOL copySheet(icStringCollection *SheetNames, char *xlsName, char *sSheetName)
{
    return TRUE;
}

/**
*   Установка цвета текста и цвета фона для области ячеек
*/
BOOL setRangeInterior(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd,
                      COLORREF BackgroundColor)
{
    return TRUE;
}

icInterior *getRangeInterior(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    return new icInterior();
}


/**
*   Установить отображение/скрытие колонок
*/
BOOL setRangeColumnHidden(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, BOOL bEnable)
{
    unsigned int col1 = lxw_name_to_col(sAddressBegin);
    unsigned int col2 = lxw_name_to_col(sAddressEnd);
    lxw_row_col_options options = {.hidden = (bEnable)?1:0, .level = 0, .collapsed = 0};
    
    if (DBG_MODE) logAddLine("Set column hidden <%d> range [%d : %d]", bEnable, col1, col2);
    
    worksheet_set_column_opt(oSheet, col1, col2, LXW_DEF_COL_WIDTH, NULL, &options);
    
    return TRUE;
}

/**
*   Установить отображение/скрытие строк
*/
BOOL setRangeRowHidden(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, BOOL bEnable)
{
    unsigned int row1 = lxw_name_to_row(sAddressBegin);
    unsigned int row2 = lxw_name_to_row(sAddressEnd);
    lxw_row_col_options options = {.hidden = (bEnable)?1:0, .level = 0, .collapsed = 0};
    
    if (DBG_MODE) logAddLine("Set row hidden <%d> range [%d : %d]", bEnable, row1, row2);
    
    for (unsigned int i_row=row1; i_row <= row2; i_row++)
    {
        worksheet_set_row_opt(oSheet, i_row, LXW_DEF_ROW_HEIGHT, NULL, &options);
    }
    return TRUE;
}

/**
*   Установить шрифт области ячеек
*/
static BOOL _set_cell_font(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                           unsigned int row, unsigned int col,
                           char *sFontName, unsigned int iFontSize, COLORREF TextColor,
                           BOOL bBold, BOOL bItalic, BOOL bUnderline)
{
    char *value = _get_cell_value(oSheet, row, col);
        
    lxw_format *format = _get_cell_format(oSheet, row, col);
    
    if (bBold || bItalic || bUnderline || !strempty(sFontName) || (iFontSize > 0) || TextColor)
    {
        if (format == NULL)
            format = workbook_add_format(oBook);
        else
            format = _xlsx_clone_format(oBook, format);
        
        if (bBold)
            format_set_bold(format);
        if (bItalic)
            format_set_italic(format);
        if (bUnderline)
            format_set_underline(format, LXW_UNDERLINE_SINGLE);
        if (!strempty(sFontName))
            format_set_font_name(format, sFontName);
        if (iFontSize > 0)
            format_set_font_size(format, iFontSize);
        if (TextColor)
            format_set_font_color(format, TextColor);
    }    
        
    set_cell_value(oBook, oSheet, row, col, value, format);
    //Если это число, породилась промежуточная строка
    //и необходимо очистить память после ее использования
    if (isnumeric(value))
        value = strfree(value);    
}

BOOL setRangeFont(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                  char *sAddressBegin, char *sAddressEnd,
                  char *sFontName, unsigned int iFontSize, COLORREF TextColor,
                  BOOL bBold, BOOL bItalic, BOOL bUnderline)
{
    if (strempty(sAddressBegin))
    {
        if (DBG_MODE) errBox("Not define address range [%s : %s]", sAddressBegin, sAddressEnd);
    }
    else if (strempty(sAddressEnd) || strequal(sAddressBegin, sAddressEnd))
    {
        //Диапазон - это 1 ячейка
        unsigned int i_row = lxw_name_to_row(sAddressBegin);
        unsigned int i_col = lxw_name_to_col(sAddressBegin);
        _set_cell_font(oBook, oSheet, i_row, i_col, 
                       sFontName, iFontSize, TextColor,
                       bBold, bItalic, bUnderline);
    }
    else
    {
        unsigned int row1 = lxw_name_to_row(sAddressBegin);
        unsigned int col1 = lxw_name_to_col(sAddressBegin);
        unsigned int row2 = lxw_name_to_row(sAddressEnd);
        unsigned int col2 = lxw_name_to_col(sAddressEnd);
        

        for (unsigned int i_row=row1; i_row <= row2; i_row++)
            for (unsigned int i_col=col1; i_col <= col2; i_col++)
                _set_cell_font(oBook, oSheet, i_row, i_col, 
                               sFontName, iFontSize, TextColor,
                               bBold, bItalic, bUnderline);
    }
    
    return TRUE;
}

icFont *getRangeFont(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    return NULL;
}

static BOOL _set_cell_border(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                            unsigned int row, unsigned int col,
                            DWORD topStyle, DWORD leftStyle, DWORD bottomStyle, DWORD rightStyle)
{
    char *value = _get_cell_value(oSheet, row, col);
      
    lxw_format *format = _get_cell_format(oSheet, row, col);
    
    if (format == NULL)
        format = workbook_add_format(oBook);
    else if (topStyle || leftStyle || bottomStyle || rightStyle)
        format = _xlsx_clone_format(oBook, format);

    if (topStyle)
        format_set_top(format, topStyle);
    if (leftStyle)
        format_set_left(format, leftStyle);
    if (bottomStyle)
        format_set_bottom(format, bottomStyle);
    if (rightStyle)
        format_set_right(format, rightStyle);
        
    if (DBG_MODE) logAddLine("Set cell [%d : %d] border top: %d left: %d bottom: %d right: %d", row, col, topStyle, leftStyle, bottomStyle, rightStyle);
    set_cell_value(oBook, oSheet, row, col, value, format);

    //Если это число, породилась промежуточная строка
    //и необходимо очистить память после ее использования
    if (isnumeric(value))
        value = strfree(value);    
}

/**
*   Установить обрамление области ячеек
*/
BOOL setRangeBorder(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    char *sAddressBegin, char *sAddressEnd,
                    DWORD topStyle, DWORD leftStyle, DWORD bottomStyle, DWORD rightStyle)
{
    //if (DBG_MODE) logAddLine("Set border [%s : %s]", sAddressBegin, sAddressEnd);
    if (strempty(sAddressBegin))
    {
        if (DBG_MODE) errBox("Not define address range [%s : %s]", sAddressBegin, sAddressEnd);
    }
    else if (strempty(sAddressEnd) || strequal(sAddressBegin, sAddressEnd))
    {
        //Диапазон - это 1 ячейка
        unsigned int i_row = lxw_name_to_row(sAddressBegin);
        unsigned int i_col = lxw_name_to_col(sAddressBegin);
        _set_cell_border(oBook, oSheet, i_row, i_col, topStyle, leftStyle, bottomStyle, rightStyle);
    }
    else
    {
        unsigned int row1 = lxw_name_to_row(sAddressBegin);
        unsigned int col1 = lxw_name_to_col(sAddressBegin);
        unsigned int row2 = lxw_name_to_row(sAddressEnd);
        unsigned int col2 = lxw_name_to_col(sAddressEnd);

        for (unsigned int i_row=row1; i_row <= row2; i_row++)
            for (unsigned int i_col=col1; i_col <= col2; i_col++)
                _set_cell_border(oBook, oSheet, i_row, i_col, topStyle, leftStyle, bottomStyle, rightStyle);
        
    }
    
    return TRUE;
}

BOOL setRangeBorderAround(lxw_workbook *oBook, lxw_worksheet *oSheet, 
                    char *sAddressBegin, char *sAddressEnd,
                    DWORD topStyle, DWORD leftStyle, DWORD bottomStyle, DWORD rightStyle)
{
    if (strempty(sAddressBegin))
    {
        if (DBG_MODE) errBox("Not define address range [%s : %s]", sAddressBegin, sAddressEnd);
    }
    else if (strempty(sAddressEnd) || strequal(sAddressBegin, sAddressEnd))
    {
        //Диапазон - это 1 ячейка
        unsigned int i_row = lxw_name_to_row(sAddressBegin);
        unsigned int i_col = lxw_name_to_col(sAddressBegin);
        _set_cell_border(oBook, oSheet, i_row, i_col, topStyle, leftStyle, bottomStyle, rightStyle);
    }
    else
    {
        unsigned int row1 = lxw_name_to_row(sAddressBegin);
        unsigned int col1 = lxw_name_to_col(sAddressBegin);
        unsigned int row2 = lxw_name_to_row(sAddressEnd);
        unsigned int col2 = lxw_name_to_col(sAddressEnd);

        //Вершняя + нижняя строка
        for (unsigned int i_col=col1; i_col <= col2; i_col++)
        {
            _set_cell_border(oBook, oSheet, row1, i_col, topStyle, 0, 0, 0);
            _set_cell_border(oBook, oSheet, row2, i_col, 0, 0, bottomStyle, 0);
        }
            
        //Права + левая колонка
        for (unsigned int i_row=row1; i_row <= row2; i_row++)
        {
            _set_cell_border(oBook, oSheet, i_row, col1, 0, leftStyle, 0, 0);
            _set_cell_border(oBook, oSheet, i_row, col2, 0, 0, 0, rightStyle);
        }        
    }
}

/**
*   Установить выравнивание области ячеек
*/
BOOL setRangeAlignment(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd,
                       BOOL bVerticalAlignment, BOOL bHorizontalAlignment)
{
    return TRUE;
}

icAlignment *getRangeAlignment(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    return new icAlignment();
}

/**
*   Установить числовой формат области ячеек
*/
BOOL setRangeNumberFormat(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    return TRUE;
}

icNumberFormat *getRangeNumberFormat(lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd)
{
    return new icNumberFormat();
}

BOOL setCellValue(lxw_workbook *oBook, lxw_worksheet *oSheet, char *sCellAddress, char *sValue)
{
    unsigned int i_row = lxw_name_to_row(sCellAddress);
    unsigned int i_col = lxw_name_to_col(sCellAddress);
    
    if (sValue)
    {
        //if (DBG_MODE) logAddLine("setCellValue [%d, %d] : <%s>", i_row, i_col, sValue);
        return set_cell_value(oBook, oSheet, i_row, i_col, sValue);
    }
    return FALSE;
}

char *getCellValue(lxw_worksheet *oSheet, char *sCellAddress)
{
    unsigned int i_row = lxw_name_to_row(sCellAddress);
    unsigned int i_col = lxw_name_to_col(sCellAddress);
    return _get_cell_value(oSheet, i_row, i_col);
}

char *getNextCellAddress(lxw_worksheet *oSheet, char *sCellAddress, int iColumnOffset)
{
    unsigned int i_row = lxw_name_to_row(sCellAddress);
    unsigned int i_col = lxw_name_to_col(sCellAddress) + iColumnOffset;
    return get_cell_address(i_row, i_col);
}

char *getOffsetCellAddress(lxw_worksheet *oSheet, char *sCellAddress, int iRowOffset, int iColumnOffset)
{
    unsigned int i_row = lxw_name_to_row(sCellAddress) + iRowOffset;
    unsigned int i_col = lxw_name_to_col(sCellAddress) + iColumnOffset;
    return get_cell_address(i_row, i_col);
}


char *get_cell_address(unsigned int iRow, unsigned int iColumn)
{
    if (iRow >= MAX_ROW_COUNT)
    {
        if (DBG_MODE) logWarning("Maximum row count [%d]", iRow);
        return NULL;
    }
    if (iColumn >= MAX_COLUMN_COUNT)
    {
        if (DBG_MODE) logWarning("Maximum column count [%d]", iColumn);
        return NULL;
    }
    
    // Количество букв в английском алфавите   v
    char *cell_address = strgen(' ', iColumn / 26 + 1 + digit_count(iRow));
    lxw_rowcol_to_cell(cell_address, iRow, iColumn);
    return cell_address; 
}

BOOL setRangeValue(lxw_workbook *oBook, lxw_worksheet *oSheet, char *sAddressBegin, char *sAddressEnd, char *sValue)
{
    if (strequal(sAddressBegin, sAddressEnd))
        return setCellValue(oBook, oSheet, sAddressBegin, sValue);
    else
    {
        unsigned int row1 = lxw_name_to_row(sAddressBegin);
        unsigned int col1 = lxw_name_to_col(sAddressBegin);
        unsigned int row2 = lxw_name_to_row(sAddressEnd);
        unsigned int col2 = lxw_name_to_col(sAddressEnd);
    
        if (DBG_MODE) logAddLine("Set range value [%d, %d] : [%d, %d]", row1, col1, row2, col2);

        for (unsigned int i_row=row1; i_row <= row2; i_row++)
        {
            for (unsigned int i_col=col1; i_col <= col2; i_col++)
            {
            }
        }
    }
        
    return TRUE;
}

static char *_get_worksheet_name(lxw_worksheet *oSheet)
{
    char *result = strgen_empty();
    
    if (DBG_MODE) logAddLine("Get worksheet name");
    if (oSheet)
    {
        result = strcopy(oSheet->name);
        if (DBG_MODE) logAddLine("\t<%s>", result);
    }
    else
        if (DBG_MODE) logAddLine("\tNot define worksheet");    
        
    return result;
}

/**
*   Получить лист и сделать его активным
*/
lxw_worksheet *get_worksheet_by_name(lxw_workbook *oBook, char *sSheetName)
{
    lxw_worksheet *worksheet = STAILQ_FIRST(oBook->worksheets);

    if (DBG_MODE) logAddLine("Get worksheet by name <%s>", sSheetName);    
     
    if (!strequal(_get_worksheet_name(worksheet), sSheetName)) 
        STAILQ_FOREACH(worksheet, oBook->worksheets, list_pointers)
            if (strequal(_get_worksheet_name(worksheet), sSheetName)) 
                break;

    if (worksheet)
        worksheet_select(worksheet);
    return worksheet;
}

lxw_worksheet *get_worksheet_by_idx(lxw_workbook *oBook, unsigned int iIndex)
{
    lxw_worksheet *worksheet = STAILQ_FIRST(oBook->worksheets);

    while (worksheet && iIndex > 0)
    {
        STAILQ_NEXT(worksheet, list_pointers);
        iIndex--;
    }
    
    if (worksheet)
        worksheet_select(worksheet);
    return worksheet;
}