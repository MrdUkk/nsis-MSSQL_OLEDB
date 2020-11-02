#pragma once
#include <windows.h>
#include <oleauto.h>
#include <comutil.h>

class MMDBProperties
{
public:
	MMDBProperties(void);
	~MMDBProperties(void);
	HRESULT Add(VARIANT * val,DBPROPID prop);
	HRESULT NumElements(ULONG *num);
	HRESULT GetElement(long num,VARIANT *val,DBPROPID *prop);
private:
	SAFEARRAY *sap,*sapid;
	SAFEARRAYBOUND sb[1];
	LONG cnt;
};
