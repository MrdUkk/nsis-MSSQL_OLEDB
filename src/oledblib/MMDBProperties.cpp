/* MMDBProperties - Class to handle OLEDB connection properties
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
#include <MMDBProperties.h>

MMDBProperties::MMDBProperties(void) : 
sap(NULL), cnt(0), sapid(NULL)
{
sb[0].cElements=0;
sb[0].lLbound=0;
sap=SafeArrayCreate(VT_VARIANT,1,sb);
sapid=SafeArrayCreate(VT_I4,1,sb);
}

MMDBProperties::~MMDBProperties(void)
{
	if(sap)
		SafeArrayDestroy(sap);
	if(sapid)
		SafeArrayDestroy(sapid);
}

HRESULT MMDBProperties::Add(VARIANT *val,DBPROPID prop)
{
	HRESULT hr,hr1;
	LONG curpos=cnt;
	++cnt;
	sb[0].cElements=cnt;
	hr=SafeArrayRedim(sap,sb);
	hr1=SafeArrayRedim(sapid,sb);
	if(SUCCEEDED(hr) && SUCCEEDED(hr1))
		{
		hr=SafeArrayPutElement(sap,&curpos,val);
		hr1=SafeArrayPutElement(sapid,&curpos,(void *)&prop);
		}

	return hr;
}

HRESULT MMDBProperties::NumElements(ULONG *num)
{
HRESULT hr;
long n;
*num=0L;
hr=SafeArrayGetUBound(sap,1,&n);
if(SUCCEEDED(hr))
	*num=n+1;
return hr;
} // NumElements

HRESULT MMDBProperties::GetElement(long num,VARIANT *val,DBPROPID *prop)
{
HRESULT hr;
long idx[1];
idx[0]=num;
hr=SafeArrayGetElement(sap,idx,val);
hr=SafeArrayGetElement(sapid,idx,prop);
return hr;
} // GetElement
