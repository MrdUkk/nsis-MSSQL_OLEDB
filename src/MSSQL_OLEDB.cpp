/* MSSQL OLEDB plugin for NSIS
 * Copyright (C) 2007 Stefano Giusto <sgiusto@mmpoint.it>
 *
 * This software is provided 'as-is', without any express or implied 
 * warranty. In no event will the authors be held liable for any damages 
 * arising from the use of this software. 
 *
 * Permission is granted to anyone to use this software for any purpose, 
 * including commercial applications, and to alter it and redistribute it 
 * freely, subject to the following restrictions:
 *
 *   1. The origin of this software must not be misrepresented; you must not 
 *      claim that you wrote the original software. If you use this software 
 *      in a product, an acknowledgment in the product documentation would be
 *      appreciated but is not required.
 *
 *   2. Altered source versions must be plainly marked as such, and must not 
 *      be misrepresented as being the original software.
 *
 *   3. This notice may not be removed or altered from any source 
 *      distribution.
 * ------------------------------------------------------------------------------
 * Version History
 * 1.0.0.0 03/12/2007 first release
 * 1.1.0.0 03/20/2007 fixed a bug in DllMain in Windows 2000
 * 1.2.0.0 03/29/2007 Added SQL_ExecuteScript function
 * 1.3.0.0 05/24/2007 Fixed a bug in Data Column Binding causing data truncation in rowsets
 * 1.4.0.0 09/11/2007 added support for Unicode SQL scripts in SQL_ExecuteScript function
 * 2.0.0.0 11/04/2009 added support for new plugin api
 *                    added unicode version
 *                    script length is limited by memory
 * 2.0.1.0 01/08/2009 Fixed an allocation bug. Global variables were not set to null after deletion in SQL_Logout function
 */

#include <windows.h>
#include <stdio.h>
#include <comutil.h>
#include <sqloledb.h>
//#include "..\ExDLL\exdll.h"
#include <tchar.h>
#include "pluginapi.h"
#include <MMSQLOLEDB.h>
#include <MMSQLQuery.h>
#include <MMSQLError.h>
#include "MSSQL_OLEDB.h"


static UINT_PTR PluginCallback(enum NSPIM msg)
 {
   return 0;
 }


BOOL WINAPI DllMain(
    HINSTANCE hinstDLL,  // handle to DLL module
    DWORD fdwReason,     // reason for calling function
    LPVOID lpReserved )  // reserved
{
    // Perform actions based on the reason for calling.
    switch( fdwReason ) 
    { 
        case DLL_PROCESS_ATTACH:
		  g_hInstance=(HINSTANCE) hinstDLL;
		  db=NULL;
		  q=NULL;
            break;

        case DLL_THREAD_ATTACH:
            break;

        case DLL_THREAD_DETACH:
            break;

        case DLL_PROCESS_DETACH:
            break;
    }
    return TRUE;  // Successful DLL_PROCESS_ATTACH.
}



void SQL_Logout(HWND hwndParent, int string_size, 
                                      TCHAR *variables, stack_t **stacktop,
                                      extra_parameters *extra)
{
  g_hwndParent=hwndParent;

  EXDLL_INIT();
  if(q)
	  {
	  delete q;
	  q=NULL;
	  }
  if(db)
	  {
	  delete db;
	  db=NULL;
	  }
}

void SQL_Logon(HWND hwndParent, int string_size, 
                                      TCHAR *variables, stack_t **stacktop,
                                      extra_parameters *extra)
{
  g_hwndParent=hwndParent;
  extra->RegisterPluginCallback(g_hInstance, PluginCallback);

  EXDLL_INIT();

  
  // note if you want parameters from the stack, pop them off in order.
  // i.e. if you are called via exdll::myFunction file.dat poop.dat
  // calling popstring() the first time would give you file.dat,
  // and the second time would give you poop.dat. 
  // you should empty the stack of your parameters, and ONLY your
  // parameters.

  {
	  VARIANT v;
	  _bstr_t bstr;
	  db=new MMSQLOLEDB();
	db->Initialize();
	if(!db->IsInitialized())
	{
		pushstring(TEXT("Error initializing OLEDB Environment (!IsInitialized)"));
		pushstring(TEXT("1"));
		return;
	}
	if(!db->IsOk())
	{
		pushstring(TEXT("Error initializing OLEDB Environment (!IsOk)"));
		pushstring(TEXT("1"));
		return;
	}
	V_VT(&v)=VT_I4;
	V_I4(&v)=10;
	hr=db->AddProperty(&v,DBPROP_INIT_TIMEOUT);
	TCHAR ServerName[64];
	if(popstring(ServerName))
	{
		pushstring(TEXT("Invalid number of arguments"));
		pushstring(TEXT("1"));
		return;
	}
	bstr=(TCHAR *)ServerName;
	V_VT(&v)=VT_BSTR;
	V_BSTR(&v)=bstr;
	hr=db->AddProperty(&v,DBPROP_INIT_DATASOURCE);
	TCHAR SQLUser[64];
	if(popstring(SQLUser))
	{
		pushstring(TEXT("Invalid number of arguments"));
		pushstring(TEXT("1"));
		return;
	}
	bstr=SQLUser;
	V_VT(&v)=VT_BSTR;
	V_BSTR(&v)=bstr;
	hr=db->AddProperty(&v,DBPROP_AUTH_USERID);
	TCHAR SQLPassword[64];
	if(popstring(SQLPassword))
	{
		pushstring(TEXT("Invalid number of arguments"));
		pushstring(TEXT("1"));
		return;
	}
	bstr=SQLPassword;
	V_VT(&v)=VT_BSTR;
	V_BSTR(&v)=bstr;
	hr=db->AddProperty(&v,DBPROP_AUTH_PASSWORD);
	V_BSTR(&v)=NULL;
	// if username is blank use integrated authentication
	if(_tcslen(SQLUser)==0)
		{
		hr=db->AddProperty(&v,DBPROP_AUTH_INTEGRATED);
		}
	hr=db->SetInitProps();
	if(FAILED(hr))
	{
		pushstring(TEXT("Error initializing OLEDB Connection (SetInitProps)"));
		pushstring(TEXT("1"));
		return;
	}
	hr=db->GetDBI()->Initialize();
	if(FAILED(hr))
	{
		pushstring(TEXT("Error initializing OLEDB Connection (Initialize)"));
		pushstring(TEXT("1"));
		return;
	}
	pushstring(TEXT("Logon successfull"));
	pushstring(TEXT("0"));
  }
}





void SQL_Execute(HWND hwndParent, int string_size, 
                                      TCHAR *variables, stack_t **stacktop,
                                      extra_parameters *extra)
{
  g_hwndParent=hwndParent;

  EXDLL_INIT();
{
	TCHAR command[1024];
	if(popstring(command))
	{
		pushstring(TEXT("Invalid number of arguments"));
		pushstring(TEXT("1"));
		return;
	}
	if(q)
		delete q;
	q= new MMSQLQuery(db);
	hr=q->Statement()->SetCommandText(command);
	hr=q->Execute();
	if(FAILED(hr))
		{
		pushstring(TEXT("Query Failed"));
		pushstring(TEXT("1"));
		}
	else
		{
		pushstring(TEXT("Query Successfull"));
		pushstring(TEXT("0"));
		}

}

}

void SQL_GetRow(HWND hwndParent, int string_size, 
                                      TCHAR *variables, stack_t **stacktop,
                                      extra_parameters *extra)
{
  g_hwndParent=hwndParent;

  EXDLL_INIT();
  if(!q)
	  {
		  pushstring(TEXT("No query"));
		  pushstring(TEXT("1"));
		  return;
	  }
  TCHAR row[1024];
  TCHAR col[256];
  ULONG ncols=0L;
  ULONG i;
  if(!q->RS()->GetRowset())
		{
		  pushstring(TEXT("Error getting rowset"));
		  pushstring(TEXT("1"));
		  return;
		}
  if(FAILED(q->RS()->Init()))
		{
		  pushstring(TEXT("Error initializing rowset"));
		  pushstring(TEXT("1"));
		  return;
		}
  ncols=q->RS()->GetNCols();
  row[0]='\0';
  if(!q->RS()->FetchRecord())
	  {
		  pushstring(TEXT("No more data"));
		  pushstring(TEXT("2"));
		  return;
	  }

  for(i=0;i<ncols;i++)
	  {
		  q->RS()->GetCol(i,col,256);
		  if(i>0)
			  _stprintf_s(row,1024,TEXT("%s|"),row);
		  _stprintf_s(row,1024,TEXT("%s%s"),row,col);
	  }
  pushstring(row);
  pushstring(TEXT("0"));
}

void SQL_GetError(HWND hwndParent, int string_size, 
                                      TCHAR *variables, stack_t **stacktop,
                                      extra_parameters *extra)
{
  g_hwndParent=hwndParent;

  EXDLL_INIT();
  TCHAR Errors[1024];
  Errors[0]='\0';
  MMSQLError err;
  err.ReportErrors(hr,Errors,1024);
  pushstring(Errors);
  pushstring(TEXT("0"));
}


void SQL_ExecuteScript(HWND hwndParent, int string_size, 
                                      TCHAR *variables, stack_t **stacktop,
                                      extra_parameters *extra)
{
  g_hwndParent=hwndParent;

  EXDLL_INIT();
  TCHAR fname[1024];
  SQL_Script s;
  if(popstring(fname))
	{
		pushstring(TEXT("Invalid number of arguments"));
		pushstring(TEXT("1"));
		return;
	}
  if(!s.Init(fname))
	{
		pushstring(TEXT("Error initializing script"));
		pushstring(TEXT("1"));
		return;
	}
  if(FAILED(s.Execute()))
	{
		pushstring(TEXT("Error executing script"));
		pushstring(TEXT("1"));
		return;
	}
	pushstring(TEXT("Script executed successfully"));
	pushstring(TEXT("0"));
	return;
} // SQL_ExecuteScript


SQL_Script::~SQL_Script()
{
if(Command!=NULL)
	HeapFree(heap,0,Command);
if (AllFileData != NULL)
	HeapFree(heap, 0, AllFileData);
} // ~SQL_Script

SQL_Script::SQL_Script() : Command(NULL), AllFileData(NULL)
{
	heap=GetProcessHeap();
} // SQL_Script

bool SQL_Script::Init(TCHAR *scriptFile)
{
	HANDLE hFile = CreateFile(scriptFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
	if (hFile == NULL)
		return false;

	DWORD ReadedBytes = 0, dwFileSize = GetFileSize(hFile, NULL);
	AllFileData = (TCHAR *)HeapAlloc(heap, 0, dwFileSize + 1);
	Command = (TCHAR *)HeapAlloc(heap, 0, dwFileSize + 1);
	if (AllFileData == NULL || Command == NULL)
	{
		CloseHandle(hFile);
		return false;
	}
	ReadFile(hFile, AllFileData, dwFileSize, &ReadedBytes, NULL);
	CloseHandle(hFile);
	if (ReadedBytes != dwFileSize)
		return false;
	return true;
} // Init

HRESULT SQL_Script::Execute(void)
{
	HRESULT hr;
	//split script to chunks (divide by "GO\r\n" statements)
	size_t delimitersz = _tcslen(_T("GO\r\n"));
	TCHAR *curstr, *allstr = AllFileData;
	while (allstr)
	{
		curstr = _tcsstr(allstr, _T("GO\r\n"));
		if (curstr == NULL) break;
		size_t a = curstr - allstr;
		_tcsncpy((TCHAR *)Command, allstr, a);
		Command[a] = _T('\0');

		if (q)
			delete q;
		q = new MMSQLQuery(db);
		hr = q->Statement()->SetCommandText(Command);
		if (FAILED(hr))
		{
			return(hr);
		}
		hr = q->Execute();
		if (FAILED(hr))
		{
			return(hr);
		}
		allstr += a + delimitersz;
	}
	return(hr);
} // Execute
