#pragma once


#define NUMROWS_CHUNK 35

class MMSQLRowSet
{
public:
	MMSQLRowSet(void);
	~MMSQLRowSet(void);
	IRowset *GetRowset(void) {return(rowset);};
	void SetRowset(IRowset *rs) {rowset=rs;};
	HRESULT Init(void);
	HRESULT GetColumnsInfo(void);
	void CreateDBBindings(void);
	void GetData(FILE *stream);
	ULONG GetNCols(void) {return nCols;};
	bool FetchRecord(void);
	int GetCol(ULONG col, TCHAR *dest, int maxlen);
	int GetColTitle(ULONG col, TCHAR *dest, int maxlen);
private:
	IRowset *rowset;
	ULONG nCols, recordsInMemory, currentRecord;
	bool Eof;
	DBCOLUMNINFO *pColumnsInfo;
	OLECHAR *pColumnStrings;
	DBBINDING *pDBBindings;
	DBBINDSTATUS *pDBBindStatus;
	unsigned char *pRowValues;
	HROW rghRows[NUMROWS_CHUNK]; // Row handles
	HROW* pRows; // Pointer to the row
	// handles
	IAccessor* pIAccessor; // Pointer to the
	// accessor
	HACCESSOR hAccessor; // Accessor handle
	void Read(void);
};
