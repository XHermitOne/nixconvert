/**
* Модуль функции тестирования основных программных блоков.
* @file
*/
#include "econvert.h"

/**
* Режим отладки
*/
BOOL DBG_MODE = TRUE;

/*
void test_StructCFG(char *CFGFileName)
{
    printf("START test_StructCFG. CFG file <%s>\n", CFGFileName);
    icStructCFG *cfg = new icStructCFG();
    cfg->AnalalizeCfgFile(CFGFileName);
    cfg->PrintCFG();
    printf("STOP test_StructCFG\n");
}


void test_dbFile(char *FileName)
{
    printf("START test_dbFile. DBF file <%s>\n", FileName);
    dbFile *dbf = new dbFile(FileName);
    unsigned int rec_count = dbf->getRecCount();
    printf("Record count: <%d>\n", rec_count);

    for (unsigned int i=0; i<rec_count; i++)
    {
        printf("\tRecord: <%d>\n", i);
        dbf->readRecord(i);

        char *fname = cp866_to_utf8(dbf->getField("FNAME"));
        char *table = cp866_to_utf8(dbf->getField("TABLE"));
        char *cell = cp866_to_utf8(dbf->getField("CELL"));
        char *account = cp866_to_utf8(dbf->getField("ACCOUNT"));

        printf("\t\tFNAME: <%s> TABLE: <%s> CELL: <%s> ACCOUNT: <%s>\n", fname, table, cell, account);

        fname = strfree(fname);
        table = strfree(table);
        cell = strfree(cell);
        account = strfree(account);

        dbf->Next();
    }

    dbf->Close();
    printf("STOP test_dbFile\n");
}


void test_ods(char *FileName, char *NewFileName)
{
    printf("START test_ods. ODS file <%s>\n", FileName);
    QString ods_filename(FileName);
	ods::Book book(ods_filename);
    printf("Sheet count: <%d>\n", book.sheet_count());

    for (unsigned int i=0; i<book.sheet_count(); i++)
    {
        ods::Sheet *sheet = book.sheet(i);
        printf("\tSheet: <%s>\tRow count: <%d>\n",
                sheet->name().toLatin1().data(),
                sheet->rows()->count());
        for (unsigned int j=0; j<sheet->rows()->count(); j++)
        {
            ods::Row *row = sheet->rows()->value(j);
            //row->CreateCell(15)->SetValue("Blah-blah-blah");
            for (unsigned int jj=0; jj < row->cells().count(); jj++)
            {
                ods::Cell *cell = row->cells().value(jj);
                //cell->SetValue("Blah-blah-blah");
                printf("\t\tCell address: <%s> ", cell->Address().toLatin1().data());
                if (cell->HasFormula())
                    printf("formula: <%s> ", cell->formula()->formula().toLatin1().data());

                if (!cell->IsEmpty())
                {
                    if (cell->value().IsString())
                        printf("string: <%s>", cell->value().AsString()->toLatin1().data());
                    if (cell->value().IsCurrency())
                        printf("currency: <%f>", *cell->value().AsCurrency());
                    if (cell->value().IsDouble())
                        printf("double: <%f>", *cell->value().AsDouble());
                    if (cell->value().IsPercentage())
                        printf("percentage: <%f>", *cell->value().AsPercentage());
                    if (cell->value().IsBool())
                        printf("bool: <%d>", *cell->value().AsBool());
                    if (cell->value().IsDate())
                        printf("date: <%s>", cell->value().AsDate()->toString().toLatin1().data());
                    if (cell->value().IsTime())
                        printf("time: <%s>", cell->value().AsDate()->toString().toLatin1().data());

                }
                printf("\n");
            }
        }

    }
    book.Save(QString(NewFileName));

    printf("STOP test_ods\n");
}



void test_sheet_manager(char *FileName)
{
    printf("START test_ods_sheet_manager. ODS file <%s>\n", FileName);
    lxw_workbook *book = open_workbook_ods(FileName);
    printf("Sheet count: <%d>\n", book->num_sheets);

    for (unsigned int i=0; i < book->num_sheets; i++)
    {
        lxw_worksheet *sheet = get_worksheet_by_idx(book, i);
        printf("\tSheet: <%s>\n", sheet->name);

        cleanRange(sheet, "B17", "F21");
        //cleanSheet(sheet);

        icStringCollection str_coll = icStringCollection();
        str_coll.Init("B17", "F21", "B17fgh", "F2eeee1", "Cr12", (char *)NULL);
        str_coll.PrintItemValues();
        setRangeValueArray(sheet, "B17", "F21", &str_coll);

        icStringCollection *ret_coll = NULL;
        ret_coll = getRangeValueArray(sheet, "B17", "F21");
        ret_coll->PrintItemValues();

        setRangeColumnWidth(sheet, "B17", "F21", 250);

        setCellValue(sheet, "S3", "BBB");
        setCellValue(sheet, "AB3", "BBB");
        setCellValue(sheet, "AP3", "BBB");
        //setRangeValue(sheet, "A3", "A3", "BBB");
        //setRangeValue(sheet, "B3", "D3", "AAA");
    }
    close_workbook_xlsx(book);

    printf("STOP test_ods_sheet_manager\n");
}


void test_functions(void)
{
    printf("START test_functions.\n");

    char *cell1 = NULL;
    char *cell2 = NULL;

    //cell1= divideToCellBeg("AB14:EP22");
    //cell2= divideToCellEnd("AB14:EP22");
    //if (DBG_MODE) logAddLine("Function: divideToCell Cell1 <%s> Cell2 <%s>", cell1, cell2);
    //cell1 = strfree(cell1);
    //cell2 = strfree(cell2);

    //cell1= divideToCellBeg("AF12");
    //cell2= divideToCellEnd("AF12");
    //if (DBG_MODE) logAddLine("Function: divideToCell Cell1 <%s> Cell2 <%s>", cell1, cell2);
    //cell1 = strfree(cell1);
    //cell2 = strfree(cell2);

    printf("STOP test_functions\n");
}

*/
void test_xlsx(char *FileName)
{
    printf("START test_xlsx.\n");
    lxw_workbook *book = open_workbook_xls(FileName);
    //lxw_workbook *book = new_workbook(FileName);
    printf("Open...\n");
    if (book)
        close_workbook_xlsx(book);
    printf("STOP test_xlsx\n");
}

void test_string_collection(void)
{
    icStringCollection *str_collection = new icStringCollection();

    str_collection->AppendValue("AAAA");
    //str_collection->AppendValue("BBBB");
    //str_collection->AppendValue("CCCC");
    str_collection->PrintItemValues();

    str_collection->InsertValue(0, "DDD");
    //str_collection = str_collection->GetRoot();
    str_collection->PrintItemValues();

    str_collection->RemoveAll();
    str_collection->PrintItemValues();

    delete str_collection;
}

void test_array(void)
{
    icArray *array = new icArray();

    void *v1 = calloc(12, sizeof(void));
    void *v2 = calloc(12, sizeof(void));
    void *v3 = calloc(12, sizeof(void));

    array->AppendItem(v1);
    array->AppendItem(v2);
    array->AppendItem(v3);

    //array->FreeAll();

    delete array;
}

void test_convert_number(void)
{
    printf("START test_convert_number.\n");
    char *src = "  1,762,735.00";
    char *dst = strcopy(src);
    dst = convertNumber(dst);
    printf("Convert number <%s> result <%s>\n", src, dst);

    src = "33,651,773.93";
    dst = strinit(dst, strcopy(src));
    dst = convertNumber(dst);
    dst = convertDate(dst);
    printf("Convert number <%s> result <%s>\n", src, dst);

    dst = cp866_to_utf8(dst, TRUE);
    printf("<%s> to utf8 <%s>\n", src, dst);

    printf("STOP test_convert_number\n");
}

void test_convert_col_width(void)
{
    printf("START test_convert_col_width.\n");

    lxw_workbook  *workbook  = new_workbook("./debug.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    for (int i=0; i < 20; i++)
    {
        double q = worksheet_set_column(worksheet, i+2, i+2, ((double) i), NULL, NULL);
        worksheet_write_number(worksheet, i, 0, i, NULL);
    }
    workbook_close(workbook);


    printf("STOP test_convert_col_width\n");
}

int main (int argc, char **argv)
{

    //test_functions();
    //test_StructCFG("/home/xhermit/dev/prj/work/nixconvert/default.cfg");
    //test_StructCFG("/home/xhermit/dev/prj/work/nixconvert/tst/cfg/u_oper00.cfg");
    //test_dbFile("/home/xhermit/dev/prj/work/nixconvert/tst/export.dbf");
    //test_ods("/home/xhermit/dev/prj/work/nixconvert/tst/invoice_original.ods", "/home/xhermit/dev/prj/work/nixconvert/tst/invoice.ods");
    //test_sheet_manager("/home/xhermit/dev/prj/work/nixconvert/tst/invoice.ods");
    //test_xlsx("/home/xhermit/dev/prj/work/nixconvert/tst/ttn_original.xls");
    //test_xlsx("/home/xhermit/dev/prj/work/nixconvert/tst/bpo/bpo_735u.xls");
    //test_xlsx("/home/xhermit/dev/prj/work/nixconvert/tst/oc4/oc4.xls");
    //test_xlsx("/home/xhermit/dev/prj/work/nixconvert/tst/bpr735/bpr_735u.xls");
    test_xlsx("/home/xhermit/dev/prj/work/nixconvert/tst/1.xls");
    //test_string_collection();
    //test_array();
    //test_convert_number();
    //test_convert_col_width();

    return 1;
}

