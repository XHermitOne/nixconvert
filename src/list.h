
#if !defined( __LIST_H )
#define __LIST_H

#define MAX_ARRAY_ELEMENT_COUNT 1024


/**
*   Класс массива объектов.
*/
class icArray
{
    public:

        unsigned int Count; //Количество элементов массива
        
        void *Items[MAX_ARRAY_ELEMENT_COUNT];
        
        icArray();
        ~icArray();
        
        void *SetIndexItem(void* Item, unsigned int iIndex); 
        void *GetIndexItem(unsigned int iIndex); 
        void *AppendItem(void *Item); 
        void *RemoveIndexItem(unsigned int iIndex); 
        void RemoveAll(void);
        void FreeAll(void);
        unsigned int GetCount(void);
        void PrintItemValues(void);
};

/**
*   Класс словаря объектов.
*/
class icDictItem
{
    public:
        char *Key;
        void *Value;

        icDictItem();
        ~icDictItem();
};

class icDict
{
    public:

        unsigned int Count; //Количество элементов массива
        
        icDictItem *Items[MAX_ARRAY_ELEMENT_COUNT];
        
        icDict();
        ~icDict();
        
        unsigned int GetCount(void);
        void RemoveAll(void);
        
        BOOL Set(char *Key, void *Value);
        void *Get(char *Key);
        BOOL HasKey(char *Key);
        
};


/**
*   Класс списка строк.
*/
class icStringCollection
{
    public:

        char *StrValue; 

        icStringCollection *Prev;
        icStringCollection *Next;

        icStringCollection();
        ~icStringCollection();
        
        icStringCollection *GetTHIS(void) { return this; };
        
        icStringCollection *GetPrev(void);
        icStringCollection *GetNext(void);
        
        icStringCollection *GetRoot(icStringCollection *Cur=NULL);
        icStringCollection *GetLast(icStringCollection *Cur=NULL);

        BOOL Init(char *str, ...);
        void Append(icStringCollection *Cur=NULL, icStringCollection *Item=NULL);
        void Insert(unsigned int iIndex, icStringCollection *Item=NULL);
        void Delete(icStringCollection *Item=NULL);
        void Remove(icStringCollection *Item=NULL);
        icStringCollection *RemoveAll(void);
        
        int  GetLength(icStringCollection *Root=NULL);
        int  GetIndex(icStringCollection *Root=NULL, icStringCollection *Item=NULL);
        icStringCollection *GetItem(icStringCollection *Root=NULL, unsigned int iIndex=0);
        
        char *GetItemValue(icStringCollection *Item);
        char *GetValue(void);
        char *get_value(void);
        BOOL SetItemValue(icStringCollection *Item, char *NewValue);
        BOOL SetValue(char *NewValue=NULL);
        
        BOOL AppendItemValue(icStringCollection *Item, char *NewValue);
        BOOL AppendValue(char *NewValue=NULL);
        char *GetIndexValue(unsigned int iIndex);        
        BOOL SetIndexValue(unsigned int iIndex, char *NewValue);        
        BOOL InsertItemValue(icStringCollection *Item, unsigned int iIndex, char *NewValue);
        BOOL ins_value(unsigned int iIndex, char *NewValue=NULL);
        BOOL InsertValue(unsigned int iIndex, char *NewValue=NULL);

        icStringCollection *FindItem(char *sValue);
        
        void PrintItemValues(void);
};


#endif /*__LIST_H*/
