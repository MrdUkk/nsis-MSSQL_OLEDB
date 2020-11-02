#pragma once
#include "MMSQLErrorRecord.h"

class MMSQLError
{
public:
	MMSQLError(void);
	~MMSQLError(void);
	HRESULT ReportErrors(HRESULT err, TCHAR *msg, FILE *stream);
	void ReportError(MMSQLErrorRecord *err,char *msg, FILE *stream);
	HRESULT ReportErrors(HRESULT err, TCHAR *errStr, size_t len);
	void ReportError(MMSQLErrorRecord *err,char *errStr, size_t len);
private:
    IErrorInfo *pIErrorInfo;
    IErrorRecords *pIErrorRecords;
    ULONG cRecords;
    ULONG iErr;
	HRESULT originalError;
	FILE *fout;
	TCHAR severity;

};
