#pragma once
#include "MMDBProperties.h"

class MMSQLOLEDB
	
{
public:
	MMSQLOLEDB(void);
	~MMSQLOLEDB(void);
	void Initialize(void);
	IDBInitialize *GetDBI(void) {return dbInit;};
	HRESULT AddProperty(VARIANT *val,DBPROPID prop) {return(dbp.Add(val,prop));};
	HRESULT SetInitProps(void);
	bool IsInitialized(void) {return isInitialized;};
	bool IsOk(void) {return isOk;};
private:
	IDBInitialize *dbInit;
	bool isInitialized, isOk;
	MMDBProperties dbp;
};

