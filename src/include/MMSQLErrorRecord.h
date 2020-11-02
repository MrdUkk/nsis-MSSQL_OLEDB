#pragma once

class MMSQLErrorRecord
{
public:
	MMSQLErrorRecord(HRESULT hrReturned, ULONG iRecord, IErrorRecords *pIErrorRecords);
	MMSQLErrorRecord(HRESULT hrReturned, IErrorInfo *pIErrorInfo);
	~MMSQLErrorRecord(void);
	HRESULT SetSource(IErrorInfo * pIErrorInfo);
	HRESULT SetDescription(IErrorInfo * pIErrorInfo);
	HRESULT MMSQLErrorRecord::SetSQLErrorInfo(ULONG pos,IErrorRecords *pIErrorRecords);
	char *GetDescription(void) {return Description;};
	char *GetSource(void) {return Source;};
	char *GetSQLInfo(void) {return SQLInfo;};
	LONG GetNativeError(void) {return lNativeError;};
	HRESULT GetHrError(void) {return ErrorInfo.hrError;};
private:
    _bstr_t Description,Source,SQLInfo;
	LONG lNativeError;
	ERRORINFO ErrorInfo;
};
