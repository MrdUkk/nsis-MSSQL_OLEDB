/* MMSQLErrorRecord - Class to handle MSSQL OLEDB Error Records
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
#include <MMSQLErrorRecord.h>

MMSQLErrorRecord::MMSQLErrorRecord(HRESULT hrReturned, ULONG iRecord, IErrorRecords *pIErrorRecords) :
Description(""),Source(""),SQLInfo(""),lNativeError(0)
{
	LCID lcid = GetUserDefaultLCID();
	IErrorInfo *pIErrorInfo = NULL;
	HRESULT hr;
	if(pIErrorRecords)
	{
		hr = pIErrorRecords->GetErrorInfo(iRecord, lcid, &pIErrorInfo);
		if(SUCCEEDED(hr))
		{
			SetDescription(pIErrorInfo);
			SetSource(pIErrorInfo);
			hr = pIErrorRecords->GetBasicErrorInfo(iRecord, &ErrorInfo);
			if(pIErrorInfo)
				pIErrorInfo->Release();
		}
	}
}

MMSQLErrorRecord::MMSQLErrorRecord(HRESULT hrReturned, IErrorInfo *pIErrorInfo) :
Description(""),Source(""),SQLInfo(""),lNativeError(0)
{
	if(pIErrorInfo)
	{
		SetDescription(pIErrorInfo);
		SetSource(pIErrorInfo);
	}
}


MMSQLErrorRecord::~MMSQLErrorRecord(void)
{
}

HRESULT MMSQLErrorRecord::SetSource(IErrorInfo * pIErrorInfo)
{
	BSTR str;
	HRESULT hr;
	hr = pIErrorInfo->GetSource(&str);
	if(SUCCEEDED(hr))
		Source=str;
	return hr;
}

HRESULT MMSQLErrorRecord::SetDescription(IErrorInfo * pIErrorInfo)
{
	BSTR str;
	HRESULT hr;
	hr = pIErrorInfo->GetDescription(&str);
	if(SUCCEEDED(hr))
		Description=str;
	return hr;
}


HRESULT MMSQLErrorRecord::SetSQLErrorInfo(ULONG pos,IErrorRecords *pIErrorRecords)
{
	HRESULT hr;
	ISQLErrorInfo *pISQLErrorInfo = NULL;
	BSTR str;
	// Attempt to get the ISQLErrorInfo interface for this error
	// record through GetCustomErrorObject. Note that ISQLErrorInfo
	// is not mandatory, so failure is acceptable here.
	hr = pIErrorRecords->GetCustomErrorObject(
		pos,                         // iRecord
		IID_ISQLErrorInfo,               // riid
		(IUnknown**)&pISQLErrorInfo);   // ppISQLErrorInfo

	// If we obtained the ISQLErrorInfo interface, get the SQL
	// error string and native error code for this error.
	if( pISQLErrorInfo )
		{
		hr = pISQLErrorInfo->GetSQLInfo(&str, &lNativeError);
		SQLInfo=str;
		}
	return hr;
}
