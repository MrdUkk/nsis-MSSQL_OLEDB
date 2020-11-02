/* MMSQLQuery - Class to handle a SQL Command and a SQL Rowset
 * without concerns about memory allocation
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
#include <MMSQLOLEDB.h>
#include <MMSQLQuery.h>

MMSQLQuery::MMSQLQuery(MMSQLOLEDB *db)
{
	statement=new MMSQLCommand(db->GetDBI());
	rs = new MMSQLRowSet;
}

MMSQLQuery::~MMSQLQuery()
{
	if(statement)
		delete statement;
	if(rs)
		delete rs;
}

