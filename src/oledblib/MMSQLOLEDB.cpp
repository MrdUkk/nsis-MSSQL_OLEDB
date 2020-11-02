/* MMSQLOLEDB - Class to handle MSSQL OLEDB database object
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

#define DBINITCONSTANTS // Initialize OLE constants...
#define INITGUID // ...once in each app.
#include "stdafx.h"
#include <MMSQLOLEDB.h>

MMSQLOLEDB::MMSQLOLEDB(void) :
dbInit(NULL), isInitialized(false), isOk(false)
{
	HRESULT hr;
	Initialize();
	if(isInitialized)
		hr=CoCreateInstance(CLSID_SQLOLEDB, NULL, CLSCTX_INPROC_SERVER,IID_IDBInitialize, (void**)&dbInit);
	if(SUCCEEDED(hr))
		isOk=true;
}

MMSQLOLEDB::~MMSQLOLEDB(void)
{
	if(dbInit)
		{
		dbInit->Uninitialize();
		dbInit->Release();
		}
	if(isInitialized)
	{
		OleUninitialize();
		CoUninitialize();
	}
}

void MMSQLOLEDB::Initialize(void)
{
	HRESULT hr;
	if(!isInitialized)
	{
		hr=CoInitialize(NULL);
		if(SUCCEEDED(hr))
			hr=OleInitialize(NULL);
		if(SUCCEEDED(hr))
			isInitialized=true;
	}

} // Initialize

HRESULT MMSQLOLEDB::SetInitProps(void)
{
	ULONG nProps;
	HRESULT hr;
	IDBProperties* pIDBProperties;
	DBPROPSET rgInitPropSet;
	hr=dbp.NumElements(&nProps);
	DBPROP *InitProperties=new DBPROP[nProps];
	// Initialize common property options.
	for (ULONG i = 0; i < nProps; i++ )
	{
		VariantInit(&InitProperties[i].vValue);
		InitProperties[i].dwOptions = DBPROPOPTIONS_REQUIRED;
		InitProperties[i].colid = DB_NULLID;
		hr=dbp.GetElement(i,&InitProperties[i].vValue,&InitProperties[i].dwPropertyID);
	}
	rgInitPropSet.guidPropertySet = DBPROPSET_DBINIT;
	rgInitPropSet.cProperties = nProps;
	rgInitPropSet.rgProperties = InitProperties;
	// Set initialization properties.
	hr=dbInit->QueryInterface(IID_IDBProperties, (void**)
		&pIDBProperties);
	hr = pIDBProperties->SetProperties(1, &rgInitPropSet);
	pIDBProperties->Release();
	delete InitProperties;
	return (hr);
} // SetInitProps