
#include <stdio.h>
#include <xlsxwriter.h>

lxw_workbook *lxw_test(char *xlsx_filename)
{
    lxw_workbook  *workbook  = new_workbook(xlsx_filename);
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
    worksheet_write_number(worksheet, 1, 0, 123, NULL);

    return workbook;
}

void test_xlsx(char *FileName)
{
    printf("START test_xlsx.\n");
    //lxw_workbook *book = open_workbook_ods(FileName);
    lxw_workbook *book = lxw_test("/home/xhermit/dev/prj/work/nixconvert/tst/tst_xlsx.xlsx");
    printf("Open...\n");
    if (book)
        workbook_close(book);
    printf("STOP test_xlsx\n");    
}

int main (int argc, char **argv)
{

    test_xlsx("/home/xhermit/dev/prj/work/nixconvert/tst/invoice_original.ods");
    
    return 1;
}

