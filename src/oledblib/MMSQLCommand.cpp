/* MMSQLCommand - Class to handle MSSQL OLEDB Commands
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
 */

#include "stdafx.h"
#include <MMSQLCommand.h>

MMSQLCommand::MMSQLCommand(void) : 
pIDBCreateSession(NULL), pIDBCreateCommand(NULL), pICommandText(NULL)
{
}

MMSQLCommand::MMSQLCommand(IDBInitialize *IDBInit) :
pIDBCreateSession(NULL), pIDBCreateCommand(NULL), pICommandText(NULL)
{
	HRESULT hr;
	hr=Initialize(IDBInit);
}


MMSQLCommand::~MMSQLCommand(void)
{
	if(pICommandText)
		pICommandText->Release();
	if(pIDBCreateCommand)
		pIDBCreateCommand->Release();
	if(pIDBCreateSession)
		pIDBCreateSession->Release();
}

HRESULT MMSQLCommand::Initialize(IDBInitialize *IDBInit)
{
	HRESULT hr=E_FAIL;
	// Get the DB session object.
	if(IDBInit)
		hr=IDBInit->QueryInterface(IID_IDBCreateSession,(void**) &pIDBCreateSession);
	return hr;
} // Initialize

HRESULT MMSQLCommand::CreateSession(void)
{
	HRESULT hr=E_FAIL;
	// Create the session, getting an interface for command creation.
	if(pIDBCreateSession)
		hr = pIDBCreateSession->CreateSession(NULL, IID_IDBCreateCommand,(IUnknown**) &pIDBCreateCommand);
	return hr;
} // CreateSession

HRESULT MMSQLCommand::CreateCommand(void)
{
	HRESULT hr=E_FAIL;
	// Create the command object.
	if(!pIDBCreateCommand)
		CreateSession();
	if(pIDBCreateCommand)
		hr = pIDBCreateCommand->CreateCommand(NULL, IID_ICommandText,(IUnknown**) &pICommandText);
	return hr;
} // CreateCommand

HRESULT MMSQLCommand::SetCommandText(TCHAR *cmd)
{
	HRESULT hr=E_FAIL;
	// The command requires the actual text as well as an indicator
	// of its language and dialect.
	if(!pICommandText)
		CreateCommand();
	if(pICommandText)
		{
		_bstr_t wcmd1;
		wcmd1=cmd;
		hr=pICommandText->SetCommandText(DBGUID_DBSQL, wcmd1);
		}
	return hr;
} // SetCommandText

HRESULT MMSQLCommand::Execute(MMSQLRowSet *rs)
{
	HRESULT hr=E_FAIL;
	IRowset* pIRowset;
	LONG cRowsAffected;
	rs->SetRowset(NULL);
	// Execute the command.
	if(pICommandText)
		hr = pICommandText->Execute(NULL, IID_IRowset, NULL,&cRowsAffected, (IUnknown**) &pIRowset);
	if(SUCCEEDED(hr))
		rs->SetRowset(pIRowset);
	return hr;
} // CreateCommand

