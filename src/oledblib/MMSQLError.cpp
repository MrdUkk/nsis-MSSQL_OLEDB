/* MMSQLError - Class to handle MSSQL OLEDB Errors
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
#include <MMSQLError.h>
#include <MMSQLErrorRecord.h>

MMSQLError::MMSQLError(void) :
 pIErrorInfo(NULL), pIErrorRecords(NULL)
{

}

MMSQLError::~MMSQLError(void)
{
	if( pIErrorInfo )
		pIErrorInfo->Release();
	if( pIErrorRecords )
		pIErrorRecords->Release();
}

HRESULT MMSQLError::ReportErrors(HRESULT err, TCHAR *msg, FILE *stream)
{
	HRESULT hr;
	_bstr_t dummy;
	// Obtain the current Error object, if any, by using the
	// OLE Automation GetErrorInfo function, which will give
	// us back an IErrorInfo interface pointer if successful
	severity=msg[10];
	hr = GetErrorInfo(0, &pIErrorInfo);
	// We've got the IErrorInfo interface pointer on the Error object
	if( SUCCEEDED(hr) && pIErrorInfo )
	{
		// OLE DB extends the OLE Automation error model by allowing
		// Error objects to support the IErrorRecords interface; this
		// interface can expose information on multiple errors.
		hr = pIErrorInfo->QueryInterface(IID_IErrorRecords, 
			(void**)&pIErrorRecords);
		if(SUCCEEDED(hr))
		{
			// Get the count of error records from the object
			hr = pIErrorRecords->GetRecordCount(&cRecords);

			// Loop through the set of error records and
			// display the error information for each one
			for( iErr = 0; iErr < cRecords; iErr++ )
			{
				MMSQLErrorRecord theErr(err,iErr,pIErrorRecords);
				dummy=msg;
				ReportError(&theErr,(char *)dummy,stream);
//				myDisplayErrorRecord(originalError, iErr, pIErrorRecords,
//					pwszFile, ulLine);
			}
		}
		// The object didn't support IErrorRecords; display
		// the error information for this single error
		else
		{
//			myDisplayErrorInfo(hrReturned, pIErrorInfo, pwszFile, ulLine);
			MMSQLErrorRecord theErr(err,pIErrorInfo);
			dummy=msg;
			ReportError(&theErr,(char *)dummy,stream);
		}
	}
	// There was no Error object, so just display the HRESULT to the user
	else
	{
			MMSQLErrorRecord theErr(err,NULL);
			dummy=msg;
			ReportError(&theErr,(char *)dummy,stream);
	}

return hr;
} //ReportErrors

void MMSQLError::ReportError(MMSQLErrorRecord *theErr,char *msg,FILE *stream)
{
fprintf_s(stream,"%s\nSQL State: 0x%08x - Native: %ld\nMessage: %s %s\n%s\n",severity=='A'?"Abort":"Non Eseguito",
	theErr->GetHrError(),theErr->GetNativeError(),theErr->GetDescription(),theErr->GetSQLInfo(),msg);
}

HRESULT MMSQLError::ReportErrors(HRESULT err, TCHAR *errStr, size_t len)
{
	HRESULT hr;
	char *errStr_c=new char[len];
	errStr_c[0]='\0';
	// Obtain the current Error object, if any, by using the
	// OLE Automation GetErrorInfo function, which will give
	// us back an IErrorInfo interface pointer if successful
	hr = GetErrorInfo(0, &pIErrorInfo);
	// We've got the IErrorInfo interface pointer on the Error object
	if( SUCCEEDED(hr) && pIErrorInfo )
	{
		// OLE DB extends the OLE Automation error model by allowing
		// Error objects to support the IErrorRecords interface; this
		// interface can expose information on multiple errors.
		hr = pIErrorInfo->QueryInterface(IID_IErrorRecords, 
			(void**)&pIErrorRecords);
		if(SUCCEEDED(hr))
		{
			// Get the count of error records from the object
			hr = pIErrorRecords->GetRecordCount(&cRecords);

			// Loop through the set of error records and
			// display the error information for each one
			for( iErr = 0; iErr < cRecords; iErr++ )
			{
				char curErr[1024];
				MMSQLErrorRecord theErr(err,iErr,pIErrorRecords);
				ReportError(&theErr,curErr,1024);
				if(iErr>0)
					sprintf_s(errStr_c,len,"%s|",errStr_c);
				sprintf_s(errStr_c,len,"%s%s",errStr_c,curErr);
			}
		}
		// The object didn't support IErrorRecords; display
		// the error information for this single error
		else
		{
			MMSQLErrorRecord theErr(err,pIErrorInfo);
			ReportError(&theErr,errStr_c,len);
		}
	}
	// There was no Error object, so just display the HRESULT to the user
	else
	{
			MMSQLErrorRecord theErr(err,NULL);
			ReportError(&theErr,errStr_c,len);
	}

_bstr_t dummy=errStr_c;
_stprintf_s(errStr,len,TEXT("%s"),(TCHAR *)dummy);
delete errStr_c;
return hr;
} //ReportErrors

void MMSQLError::ReportError(MMSQLErrorRecord *theErr,char *errStr, size_t len)
{
	sprintf_s(errStr,len,"SQL State: 0x%08x - Native: %ld - Message: %s %s",
	theErr->GetHrError(),theErr->GetNativeError(),theErr->GetDescription(),theErr->GetSQLInfo());
}
