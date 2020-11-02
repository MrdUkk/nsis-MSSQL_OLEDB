/* MMSQLRowSet - Class to handle MSSQL OLEDB RowSets
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
#include <MMSQLRowSet.h>

// ROUNDUP on all platforms pointers must be aligned properly
#define ROUNDUP_AMOUNT            8
#define ROUNDUP_(size,amount) (((ULONG)(size)+((amount)-1))&~((amount)-1))
#define ROUNDUP(size)             ROUNDUP_(size, ROUNDUP_AMOUNT)

MMSQLRowSet::MMSQLRowSet(void) : rowset(NULL), 
pColumnsInfo(NULL), pColumnStrings(NULL), pDBBindings(NULL), pDBBindStatus(NULL), pRowValues(NULL),
pIAccessor(NULL), hAccessor(NULL), recordsInMemory(0), currentRecord(0), Eof(false)
{
}

MMSQLRowSet::~MMSQLRowSet(void)
{
	if(pIAccessor)
	{
		if(hAccessor)
			pIAccessor->ReleaseAccessor(hAccessor, NULL);
		pIAccessor->Release();
	}

	if(pRowValues)
		delete [] pRowValues;
	if(pDBBindStatus)
		delete [] pDBBindStatus;
	if(pDBBindings)
		delete [] pDBBindings;
	if(rowset)
		rowset->Release();
}

HRESULT MMSQLRowSet::Init(void)
{
	HRESULT hr;
	// Get the description of the rowset for use in binding structure
	// creation; see "Describing Query Results."
	hr=GetColumnsInfo();
	if(SUCCEEDED(hr))
	{
		pDBBindings = new DBBINDING[nCols];
	    memset(pDBBindings, 0, nCols * sizeof(DBBINDING));
		CreateDBBindings();
		pDBBindStatus = new DBBINDSTATUS[nCols];
		pRows = &rghRows[0]; // Pointer to the row
		// Create the accessor; see "Creating Accessors."
		hr=rowset->QueryInterface(IID_IAccessor, (void**) &pIAccessor);
		if(SUCCEEDED(hr))
		{
		hr=pIAccessor->CreateAccessor(
			DBACCESSOR_ROWDATA,// Accessor will be used to retrieve row
			// data
			nCols, // Number of columns being bound
			pDBBindings, // Structure containing bind info
			0, // Not used for row accessors
			&hAccessor, // Returned accessor handle
			pDBBindStatus // Information about binding validity
			);
		}
	}
	return hr;
} // Init

/********************************************************************
* Retrieve data from a rowset.
********************************************************************/
void MMSQLRowSet::GetData(FILE *stream)
{
	ULONG nCol;
	ULONG cRowsObtained; // Count of rows
	// obtained
	ULONG iRow; // Row count
	HROW rghRows[NUMROWS_CHUNK]; // Row handles
	HROW* pRows = &rghRows[0]; // Pointer to the row
	// handles
	IAccessor* pIAccessor; // Pointer to the
	// accessor
	HACCESSOR hAccessor; // Accessor handle
	// Create the accessor; see "Creating Accessors."
	rowset->QueryInterface(IID_IAccessor, (void**) &pIAccessor);
	pIAccessor->CreateAccessor(
		DBACCESSOR_ROWDATA,// Accessor will be used to retrieve row
		// data
		nCols, // Number of columns being bound
		pDBBindings, // Structure containing bind info
		0, // Not used for row accessors
		&hAccessor, // Returned accessor handle
		pDBBindStatus // Information about binding validity
		);
	// Process all the rows, NUMROWS_CHUNK rows at a time.
	for(nCol=0;nCol<nCols;nCol++)
	{
		if(nCol>0)
			_ftprintf_s(stream,TEXT(";"));
		_bstr_t name=pColumnsInfo[nCol].pwszName;
		_ftprintf_s(stream,TEXT("%s"),(TCHAR *)name);
	}
	_ftprintf_s(stream,TEXT("\n"));
	while (TRUE)
	{
		rowset->GetNextRows(
			0, // Reserved
			0, // cRowsToSkip
			NUMROWS_CHUNK, // cRowsDesired
			&cRowsObtained, // cRowsObtained
			&pRows ); // Filled in w/ row handles.
		// All done; there are no more rows left to get.
		if (cRowsObtained == 0)
			break;
		// Loop over rows obtained, getting data for each.
		for (iRow=0; iRow < cRowsObtained; iRow++)
		{
			rowset->GetData(rghRows[iRow], hAccessor, pRowValues);
			for (nCol = 0; nCol < nCols; nCol++)
			{
			if(nCol>0)
				_ftprintf_s(stream,TEXT(";"));
			_ftprintf_s(stream,TEXT("%s"),&pRowValues[pDBBindings[nCol].obValue]);
			}
			_ftprintf_s(stream,TEXT("\n"));
		}
		// Release row handles.
		rowset->ReleaseRows(cRowsObtained, rghRows, NULL, NULL,
			NULL);
	} // End while
	// Release the accessor.
	pIAccessor->ReleaseAccessor(hAccessor, NULL);
	pIAccessor->Release();
	return;
}

/********************************************************************
* Get the characteristics of the rowset (the ColumnsInfo interface).
********************************************************************/
HRESULT MMSQLRowSet::GetColumnsInfo(void)
{
	IColumnsInfo* pIColumnsInfo;
	HRESULT hr;
	hr=rowset->QueryInterface(IID_IColumnsInfo, (void**)&pIColumnsInfo);
	if(SUCCEEDED(hr))
	{
	hr = pIColumnsInfo->GetColumnInfo(&nCols, &pColumnsInfo,&pColumnStrings);
	if(FAILED(hr))
		nCols=0;
	pIColumnsInfo->Release();
	}
	return (hr);
}

/********************************************************************
* Create binding structures from column information. Binding
* structures will be used to create an accessor that allows row value
* retrieval.
********************************************************************/
void MMSQLRowSet::CreateDBBindings(void)
{
	ULONG nCol;
	ULONG cbRow = 0;
	for (nCol = 0; nCol < nCols; nCol++)
	{
		pDBBindings[nCol].iOrdinal = nCol+1;
		pDBBindings[nCol].obValue = cbRow + sizeof(DBSTATUS) + sizeof(ULONG);
		pDBBindings[nCol].obLength = cbRow + sizeof(DBSTATUS);
		pDBBindings[nCol].obStatus = cbRow;
		pDBBindings[nCol].pTypeInfo = NULL;
		pDBBindings[nCol].pObject = NULL;
		pDBBindings[nCol].pBindExt = NULL;
		pDBBindings[nCol].dwPart = DBPART_VALUE|DBPART_STATUS;
		pDBBindings[nCol].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
		pDBBindings[nCol].eParamIO = DBPARAMIO_NOTPARAM;
		pDBBindings[nCol].cbMaxLen = 1024;//pColumnsInfo[nCol].ulColumnSize;
		pDBBindings[nCol].dwFlags = 0;
		pDBBindings[nCol].wType = DBTYPE_STR; //pColumnsInfo[nCol].wType;
		pDBBindings[nCol].bPrecision = pColumnsInfo[nCol].bPrecision;
		pDBBindings[nCol].bScale = pColumnsInfo[nCol].bScale;
//		cbRow += pDBBindings[nCol].cbMaxLen+sizeof(IUnknown*);
		cbRow = pDBBindings[nCol].cbMaxLen+pDBBindings[nCol].obValue;
		ROUNDUP(cbRow);
	}
	pRowValues = new unsigned char[cbRow];
	return;
}

void MMSQLRowSet::Read(void)
{
	currentRecord=0;
	if(recordsInMemory)
		rowset->ReleaseRows(recordsInMemory, rghRows, NULL, NULL,
		NULL);
	rowset->GetNextRows(
		0, // Reserved
		0, // cRowsToSkip
		NUMROWS_CHUNK, // cRowsDesired
		&recordsInMemory, // cRowsObtained
		&pRows ); // Filled in w/ row handles.
	if (recordsInMemory == 0)
		Eof=true;
} // Read

bool MMSQLRowSet::FetchRecord(void)
{
if(!recordsInMemory && !Eof)
	Read();
if(currentRecord>=recordsInMemory)
	Read();
if(Eof)
	return false;
HRESULT hr=rowset->GetData(rghRows[currentRecord], hAccessor, pRowValues);
++currentRecord;
return true;
} // FetchRecord


int MMSQLRowSet::GetCol(ULONG col, TCHAR *dest, int maxlen)
{
	int retval=0;
	dest[0]=_T('\0');
	if(col<nCols)
	{
		/*if ((ULONG)((BYTE*)pData)[rgBinding[0].obStatus] == 
                  DBSTATUS_S_ISNULL)*/

		if((ULONG)((BYTE*)pRowValues)[pDBBindings[col].obStatus]!=DBSTATUS_S_ISNULL)
			{
				_bstr_t dummy=(char *)&pRowValues[pDBBindings[col].obValue];
			    _stprintf_s(dest,maxlen,_T("%s"),(TCHAR *)dummy);
			}
		retval=_tcslen(dest);
	}
	return retval;
} // GetCol

int MMSQLRowSet::GetColTitle(ULONG col, TCHAR *dest, int maxlen)
{
	int retval=0;
	dest[0]=_T('\0');
	if(col<nCols)
	{
		_bstr_t name=pColumnsInfo[col].pwszName;
		_stprintf_s(dest,maxlen,_T("%s"),(TCHAR *)name);
		retval=_tcslen(dest);
	}
	return retval;
}
