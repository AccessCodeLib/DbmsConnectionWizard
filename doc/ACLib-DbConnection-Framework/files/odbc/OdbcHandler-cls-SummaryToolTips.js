﻿NDSummary.OnToolTipsLoaded("File:odbc/OdbcHandler.cls",{45:"<div class=\"NDToolTip TClass LVisualBasic\"><div class=\"TTSummary\">ODBC data connection methods</div></div>",47:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Event for missing CurrentDb reference</div></div>",48:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Event for missing ODBC connection string</div></div>",49:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Info about the last linked element</div></div>",50:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Info about the last deleted element</div></div>",51:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Event for missing OLEDB connection string</div></div>",53:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">Active Hooks</div></div>",55:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">DAO.Database-Instanz des Frontends bzw. jener Jet-DB in der die Pass-Through-Abfragen erstellt werden sollen</div></div>",56:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">Database-Referenz zum Backend</div></div>",57:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">ODBC Connection string</div></div>",58:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype58\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function OpenRecordset(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Source&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal RecordsetType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.RecordsetTypeEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">dbOpenForwardOnly,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal RecordsetOptions&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.RecordsetOptionEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">DAO.RecordsetOptionEnum.dbSeeChanges,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal LockEdit&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.LockTypeEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">DAO.LockTypeEnum.dbOptimistic</td></tr></table></td><td class=\"PAfterParameters\">) As DAO.Recordset</td></tr></table></div></div><div class=\"TTSummary\">Open DAO.Recordset</div></div>",59:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype59\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function OpenRecordsetPT(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Source&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal RecordsetType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.RecordsetTypeEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">dbOpenForwardOnly,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal RecordsetOptions&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.RecordsetOptionEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">DAO.RecordsetOptionEnum.dbSeeChanges Or DAO.RecordsetOptionEnum.dbSQLPassThrough,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal LockEdit&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.LockTypeEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">DAO.LockTypeEnum.dbOptimistic</td></tr></table></td><td class=\"PAfterParameters\">) As DAO.Recordset</td></tr></table></div></div><div class=\"TTSummary\">Open Pass Through DAO.Recordset</div></div>",60:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype60\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Execute(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal CommandText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Options&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">DAO.RecordsetOptionEnum</td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">Execute SQL statement (CurrentDbBE.Execute)</div></div>",61:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype61\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function ExecutePT(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal CommandText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Options&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">DAO.RecordsetOptionEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">DAO.RecordsetOptionEnum.dbSQLPassThrough</td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">Execute SQL statement with Pass Trough Query</div></div>",62:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype62\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function LookupSql(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal SqlText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Index&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span>&amp;,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Lookup DAO.Recordset replacement function for DLookup (passing a SQL statement) via CurrentDbBE</div></div>",63:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype63\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function LookupSqlPT(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal SqlText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Index&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span>&amp;,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Lookup DAO.Recordset replacement function for DLookup (passing a SQL statement) via Pass Through Query</div></div>",64:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype64\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Lookup(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">DAO.Recordset replacement function for DLookup (via CurrentDbBE)</div></div>",65:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype65\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Count(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">DAO.Recordset replacement function for DCount (via CurrentDbBE)</div></div>",66:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype66\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Max(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">DAO.Recordset replacement function for DMax (via CurrentDbBE)</div></div>",67:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype67\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Min(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">DAO.Recordset replacement function for DMin (via CurrentDbBE)</div></div>",68:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype68\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Sum(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">DAO.Recordset replacement function for DSum (via CurrentDbBE)</div></div>",69:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype69\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function InsertIdentityReturn(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal InsertSql&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Execute insert SQL statement and return last identity value (via CurrentDbBE)</div></div>",71:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype71\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Sub LinkTable(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal SourceTableName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal LinkedTableName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">vbNullString,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal SavePWD&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal RemoveSchemaName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">True</td></tr></table></td><td class=\"PAfterParameters\">)</td></tr></table></div></div><div class=\"TTSummary\">Link backend table in Access frontend</div></div>",72:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype72\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function RelinkTables(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">Optional ByVal SavePWD&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal EventPeriod&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span></td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">Relink existing tables</div></div>",73:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype73\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function RelinkTable(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal LinkedTableName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal SavePWD&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False</td></tr></table></td><td class=\"PAfterParameters\">) As Boolean</td></tr></table></div></div><div class=\"TTSummary\">Relink linked table with possible change of server data</div></div>",74:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype74\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function RelinkPassThroughQueries(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">Optional ByVal SavePWD&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal EventPeriod&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span></td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">Alle Pass-Through-Abfragen neu verknüpfen</div></div>",75:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype75\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function RelinkPassThroughQuery(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal QueryName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal SavePWD&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False</td></tr></table></td><td class=\"PAfterParameters\">) As Boolean</td></tr></table></div></div><div class=\"TTSummary\">Relink pass through query</div></div>",76:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype76\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function RelinkTablesAndQueries(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">Optional ByVal SavePWD&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal EventPeriod&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span></td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">Relink all linked tables and pass-through queries</div></div>",77:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype77\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function DeleteOdbcTableDefs(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">Optional ByVal EventPeriod&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span></td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">Delete all linked table in the frontend. (Has no effect on the backend tables).</div></div>",78:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype78\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function IsLinkedTable(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal TableToCheck&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">DAO.TableDef</td></tr></table></td><td class=\"PAfterParameters\">) As Boolean</td></tr></table></div></div><div class=\"TTSummary\">Check if TableDef is a linked table</div></div>",79:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype79\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function IsPassThroughQuery(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal QueryToCheck&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">DAO.QueryDef</td></tr></table></td><td class=\"PAfterParameters\">) As Boolean</td></tr></table></div></div><div class=\"TTSummary\">Check if QueryDef is a pass through query</div></div>"});