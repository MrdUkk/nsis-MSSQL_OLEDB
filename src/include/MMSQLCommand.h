#pragma once
#include "MMSQLRowSet.h"

class MMSQLCommand
{
public:
	MMSQLCommand(void);
	MMSQLCommand(IDBInitialize *IDBInit);
	~MMSQLCommand(void);
	HRESULT Initialize(IDBInitialize *IDBInit);
	HRESULT CreateSession(void);
	HRESULT CreateCommand(void);
	HRESULT SetCommandText(TCHAR *cmd);
	HRESULT Execute(MMSQLRowSet *rs);
private:
	IDBCreateSession* pIDBCreateSession;
	IDBCreateCommand* pIDBCreateCommand;
	ICommandText* pICommandText;
};
