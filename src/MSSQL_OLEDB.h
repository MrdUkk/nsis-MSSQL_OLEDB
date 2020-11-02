#pragma once

#ifdef _UNICODE
#define FILEMODE TEXT("r,ccs=UNICODE")
#define MEOF WEOF
#else
#define FILEMODE TEXT("r")
#define MEOF EOF
#endif

HINSTANCE g_hInstance;

HWND g_hwndParent;

MMSQLOLEDB *db;
MMSQLQuery *q;
HRESULT hr;


class SQL_Script
	
{
public:
	SQL_Script(void);
	~SQL_Script(void);
	bool Init(TCHAR *scriptFile);
	HRESULT Execute(void);
private:
	TCHAR *Command;
	TCHAR *AllFileData;
	HANDLE heap;
};

