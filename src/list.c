/**
*   Модуль реализации списков
*   @file
*/

#include "econvert.h"


/*******************************************************************************
*   Массив объектов
*******************************************************************************/

icArray::icArray()
{
    RemoveAll();
}


icArray::~icArray()
{
    FreeAll();
}


void *icArray::SetIndexItem(void* Item, unsigned int iIndex)
{
    //В пустые ячейки можно вставлять только Append'ом
    if (iIndex >= Count)
        return NULL;
    Items[iIndex] = Item;
    return Item;
} 


void *icArray::GetIndexItem(unsigned int iIndex)
{
    //В пустые ячейки можно вставлять только Append'ом
    if (iIndex >= Count)
        return NULL;
    
    return Items[iIndex];
}


void *icArray::RemoveIndexItem(unsigned int iIndex)
{
    void *Item = NULL;
    //В пустые ячейки можно вставлять только Append'ом
    if (iIndex >= Count)
        return NULL;
    
    Item = Items[iIndex];
    for (int i = iIndex; i < Count; i++)
        Items[i] = Items[i+1];
    Count--;    
    return Item;
}


void *icArray::AppendItem(void *Item)
{
    Items[Count] = Item;
    Count++;
}

void icArray::RemoveAll(void)
{
    for (int i=0; i<MAX_ARRAY_ELEMENT_COUNT; i++)
        Items[i] = NULL;
    Count = 0;
}

void icArray::FreeAll(void)
{
    for (int i=0; i < MAX_ARRAY_ELEMENT_COUNT; i++)
    {
        if (Items[i] != NULL)
        {
            printf("Array. Delete item <%x>\r\n", Items[i]);
            //ВНИМАНИЕ!
            //При удалении объектов класса деструктор не вызывается
            //без явного приведения к типу класса!
            delete Items[i];
        }
        Items[i] = NULL;
    }
    Count = 0;
}

/**
*   Количество элементов массива
*/
unsigned int icArray::GetCount(void)
{
    return Count;
}

void icArray::PrintItemValues(void)
{
    printf("Array <%x> (Count %d) :\r\n", this, GetCount());
    for (int i=0; i < Count; i++)
    {
        if (Items[i] != NULL)
            printf("Item: <%x>\r\n", Items[i]);
    }
    
}

/*******************************************************************************
*   Словарь объектов
*******************************************************************************/

icDictItem::icDictItem()
{
    Key = NULL;
    Value = NULL;
}


icDictItem::~icDictItem()
{
    Key = strfree(Key);
    if (Value)
    {
        delete Value;
        Value = NULL;
    }
}

icDict::icDict()
{
    RemoveAll();
}


icDict::~icDict()
{
}

/**
*   Количество элементов
*/
unsigned int icDict::GetCount(void)
{
    return Count;
}

void icDict::RemoveAll(void)
{
    for (int i=0; i<MAX_ARRAY_ELEMENT_COUNT; i++)
    {
        if (Items[i])
        {
            delete Items[i];
            Items[i] = NULL;
        }
    }
    Count = 0;
}

BOOL icDict::Set(char *Key, void *Value)
{
    icDictItem *item = NULL;
    for (unsigned int i=0; i < Count; i++)
    {
        item = Items[i];
        if (strequal(Key, item->Key))
        {
            if (item->Value)
            {
                delete item->Value;
                item->Value = NULL;
            }
            item->Value = Value;
            return TRUE;
        }
    }
    
    Items[Count] = new icDictItem();
    Items[Count]->Key = Key;
    Items[Count]->Value = Value;    
    Count++;
    return TRUE;
}

void *icDict::Get(char *Key)
{
    icDictItem *item = NULL;
    for (unsigned int i=0; i < Count; i++)
    {
        item = Items[i];
        if (strequal(Key, item->Key))
            return item->Value;
    }
}

BOOL icDict::HasKey(char *Key)
{
    icDictItem *item = NULL;
    for (unsigned int i=0; i < Count; i++)
    {
        item = Items[i];
        if (strequal(Key, item->Key))
            return TRUE;
    }
    return FALSE;
}


/*******************************************************************************
*   Список строк
*******************************************************************************/

icStringCollection::icStringCollection()
{
    Next = NULL;
    Prev = NULL;
    StrValue = NULL;
}


icStringCollection::~icStringCollection()
{
    StrValue = strfree(StrValue);
}

/**
*   Инициализация коллекции списком строк
*   ВНИМАНИЕ! Функция должна вызываться в следующем формате:
*   a.Init('str1', 'str2', .. , (char *)NULL)
*/
BOOL icStringCollection::Init(char *str, ...)
{
    va_list ap;
    char *str_next = NULL;

    va_start(ap, str);
    AppendValue(str);

    while((str_next = va_arg(ap, char *)) != NULL)
		AppendValue(str_next);

	va_end(ap);    
}

/**
*   Значение элемента списка
*/
char *icStringCollection::GetItemValue(icStringCollection *Item)
{
    if (!Item)
        Item = GetTHIS();
        
    return StrValue;
}

char *icStringCollection::GetValue(void)
{
    if (StrValue)
        return strcopy(StrValue);
    return NULL;
}

char *icStringCollection::get_value(void)
{
    return StrValue;
}

/**
*   Установть новое значение элемента списка
*   @return: Возвращает указатель на новое значение элемента.
*/
BOOL icStringCollection::SetItemValue(icStringCollection *Item, char *NewValue)
{
    if (!Item)
        Item = GetTHIS();
        
    if (NewValue)
        SetValue(NewValue);
    
    return TRUE;
}


BOOL icStringCollection::SetValue(char *NewValue)
{
    //Сначала удалить старое значение
    StrValue = strfree(StrValue);
 
    if (NewValue)
        StrValue = strcopy(NewValue);
    else
        // ВНИМАНИЕ! Здесь обязательно устанавливать NULL иначе значение
        // забивается мусором и считается что это не пустое значение
        StrValue = NULL;
        
    return TRUE;
}


BOOL icStringCollection::AppendItemValue(icStringCollection *Item, char *NewValue)
{
    icStringCollection *item = NULL;
    if (!Item)
        item = new icStringCollection();
    else
        item = Item;

    BOOL result = item->SetValue(NewValue);
    Append(NULL, item);
    
    return result;
}


BOOL icStringCollection::AppendValue(char *NewValue)
{
    BOOL not_append = (StrValue == NULL);
    
    if (not_append)
        SetValue(NewValue);
    else
        return AppendItemValue(NULL, NewValue);
    return TRUE;
}


/**
*   Значение элемента списка по индексу
*/
char *icStringCollection::GetIndexValue(unsigned int iIndex=0)
{
    icStringCollection *item = GetItem(NULL, iIndex);
    
    if (item)
        return strcopy(item->StrValue);
    return NULL;
}

/**
*   Установить значение элемента списка по индексу
*/
BOOL icStringCollection::SetIndexValue(unsigned int iIndex, char *NewValue)
{
    icStringCollection *item = GetItem(NULL, iIndex);
    
    if (item)
        return item->SetValue(NewValue);
    return FALSE;
}

/**
*   Распечатать значения элементов списка. Сделано для отладки.
*/
void icStringCollection::PrintItemValues(void)
{
    char *value;
    icStringCollection *Cur = GetRoot();

    if (DBG_MODE) logAddLine("String collection <%x> (Length %d) :", Cur, Cur->GetLength());
    else
        printf("String collection <%x> (Length %d) :\n", Cur, Cur->GetLength());
    for (; Cur; Cur=Cur->GetNext()) 
    {
        value = Cur->StrValue;
        if (value)
            if (DBG_MODE) logAddLine("\t(%x)\t%s", Cur, Cur->StrValue);
            else
                printf("\t(%x)\t%s\n", Cur, Cur->StrValue);
        else
            if (DBG_MODE) logAddLine("\t(%x)\t[NULL]", Cur);
            else
                printf("\t(%x)\t[NULL]\n", Cur);
    }
}


icStringCollection *icStringCollection::GetRoot(icStringCollection *Cur)
{
    if (!Cur)
        Cur = this;
        
    for(; Cur->Prev; Cur=Cur->Prev);
    
    return Cur;
}


icStringCollection *icStringCollection::GetLast(icStringCollection *Cur)
{
    if (!Cur)
        Cur = this;
        
    for(; Cur->Next; Cur=Cur->Next);
    
    return Cur;
}


void icStringCollection::Append(icStringCollection *Cur, icStringCollection *Item)
{
    if (Cur == NULL)
        Cur = GetLast();
        
    Item->Next = Cur->Next;
    if (Cur->Next)
        Cur->Next->Prev = Item;
    Cur->Next = Item;
    Item->Prev = Cur;
}

/**
*   Удалить элемент из списка физически
*/
void icStringCollection::Delete(icStringCollection *Item)
{
    if (!Item)
        Item = this;
    
    if (Item->Prev)
        Item->Prev->Next = Item->Next;
        
    if (Item->Next)
        Item->Next->Prev = Item->Prev;
        
    delete Item;
}


/**
*   Удалить элемент из списка
*/
void icStringCollection::Remove(icStringCollection *Item)
{
    icStringCollection *Cur;
    
    if (!Item)
        Item = GetTHIS();
    
    for (Cur=GetRoot(); Cur; Cur=Cur->Next)
        if(Cur == Item) 
        {
            Delete(Item);
            break;
        }
}

/**
*   Удалить все элементы из списка
*   @return: Функция возвращает корневой элемент списка.
*/
icStringCollection *icStringCollection::RemoveAll(void)
{
    icStringCollection *Cur = GetLast();
    icStringCollection *prev;
    
    do{
        prev = Cur->Prev;
        if (prev)
            Delete(Cur);
        Cur = prev;
    } while (prev);
       
    if (GetTHIS()->StrValue)
        GetTHIS()->SetValue(NULL);
    return GetTHIS();
}


/**
*   Количество элементов списка
*/
int icStringCollection::GetLength(icStringCollection *Root)
{
    icStringCollection *Cur;
    int l=0;
    
    if (Root == NULL)
        Root = GetRoot();

    if (Root->StrValue == NULL)
        return 0;
    else
        for (Cur=Root->Next, l=1; Cur; Cur=Cur->Next)
            l++;
    return l;
}


int icStringCollection::GetIndex(icStringCollection *Root, icStringCollection *Item)
{
    icStringCollection *Cur=NULL;
    int index;
    
    if (!Root)
        Root = GetRoot();
    if (!Item)
        Item = this;
    
    for (Cur=Root, index=0; Cur && Cur != Item; Cur=Cur->Next) 
        index ++;
            
    if (!Cur) 
        index=-1;
        
    return index;
}


icStringCollection *icStringCollection::GetItem(icStringCollection *Root, unsigned int iIndex)
{
    icStringCollection *Cur=NULL;
    unsigned int index;

    if (!Root)
        Root = GetRoot();        
        
    for (Cur=Root, index=0; Cur; Cur=Cur->Next) 
    {
        if (index == iIndex)
            break;
        index ++;
    }
            
    if (!Cur) 
        return NULL;
    return Cur;
}


icStringCollection *icStringCollection::GetPrev(void)
{
    return Prev;
}


icStringCollection *icStringCollection::GetNext(void)
{
    return Next;
}

/**
*   Вставить элемент по индексу
*/
void icStringCollection::Insert(unsigned int iIndex, icStringCollection *Item)
{
    icStringCollection *Cur = GetItem(NULL, iIndex);
        
    //Вставляем перед элементом
    Item->Prev = Cur->Prev;
    if (Cur->Prev)
        Cur->Prev->Next = Item;
    Item->Next = Cur;
    Cur->Prev = Item;
}

BOOL icStringCollection::InsertItemValue(icStringCollection *Item, unsigned int iIndex, char *NewValue)
{
    if (!Item)
        Item = new icStringCollection();        

    BOOL result = Item->SetValue(NewValue);
    Insert(iIndex, Item);
    
    return result;
}

/**
*   Вставить элемент по индексу
*/
BOOL icStringCollection::ins_value(unsigned int iIndex, char *NewValue)
{
    BOOL not_insert = (StrValue == NULL);
    
    if (not_insert)
        return SetValue(NewValue);
    else
        return InsertItemValue(NULL, iIndex, NewValue);
}


/**
*   Вставить элемент по индексу
*/
BOOL icStringCollection::InsertValue(unsigned int iIndex, char *NewValue)
{
    if (iIndex > 0)
        return ins_value(iIndex, NewValue);
    else
    {
        int len = GetLength();
        if (len > 1)
        {
            // Т.к. если вставить 0 элемент сбивается порядок списка
            // необходимо поменяться значениями
            ins_value(1, NULL);
            icStringCollection *item = GetItem(NULL, 1);
            item->StrValue = StrValue;
            StrValue = strcopy(NewValue);
        }
        else if (len == 1)
        {
            AppendValue(NULL);
            icStringCollection *item = GetItem(NULL, 1);
            item->StrValue = StrValue;
            StrValue = strcopy(NewValue);
        }
        else if (len == 0)
            return SetValue(NewValue);
    }
    return TRUE;
}

/**
*   Найти элемент по значению
*/
icStringCollection *icStringCollection::FindItem(char *sValue)
{
    icStringCollection *root = GetRoot();        
    icStringCollection *cur = NULL;                
    icStringCollection *ret = NULL;                
        
    for (cur=root; cur; cur=cur->Next) 
    {
        if (strequal(cur->get_value(), sValue))
        {
            ret = cur;
            break;
        }
    }
    return ret;
}
