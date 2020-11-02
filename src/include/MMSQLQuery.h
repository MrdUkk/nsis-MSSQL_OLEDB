#pragma once
#include "MMSQLCommand.h"
#include "MMSQLRowSet.h"

class MMSQLQuery
{
public:
	MMSQLQuery(MMSQLOLEDB *db);
	~MMSQLQuery(void);
	MMSQLCommand *Statement(void) { return statement; };
	MMSQLRowSet *RS(void) { return rs; };
	HRESULT Execute(void) {	return(statement->Execute(rs)); };
;
private:
	MMSQLCommand *statement;
	MMSQLRowSet *rs;
};
