﻿NDSummary.OnToolTipsLoaded("VisualBasicClass:data.adodb.AdodbHandler",{19:"<div class=\"NDToolTip TClass LVisualBasic\"><div class=\"TTSummary\">ADODB data connection methods</div></div>",21:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Event for missing ADODB connection reference (Passing the AdodbHandler event: data.adodb.AdodbHandler::ErrorMissingCurrentConnection)</div></div>",22:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Event for missing OLEDB connection string</div></div>",23:"<div class=\"NDToolTip TEvent LVisualBasic\"><div class=\"TTSummary\">Event for missing OLEDB connection string</div></div>",25:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">Active Hooks</div></div>",27:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">Currently set ADODB connection</div></div>",28:"<div class=\"NDToolTip TProperty LVisualBasic\"><div class=\"TTSummary\">OLEDB connection string</div></div>",30:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype30\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Execute(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal CommandText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByRef RecordsAffected&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Options&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">ADODB.ExecuteOptionEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">-1</span></td></tr></table></td><td class=\"PAfterParameters\">) As ADODB.Recordset</td></tr></table></div></div><div class=\"TTSummary\">Execute SQL statement</div></div>",31:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype31\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function ExecuteCommand(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal CmdText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">ByVal CmdType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">ADODB.CommandTypeEnum,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal CmdParamDefs&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByRef RecordsAffected&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Options&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">-1</span></td></tr></table></td><td class=\"PAfterParameters\">) As ADODB.Recordset</td></tr></table></div></div><div class=\"TTSummary\">Execute sql statement using ADODB.Command</div></div>",32:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype32\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function OpenRecordset(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Source&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal CursorType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">ADODB.CursorTypeEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">ADODB.CursorTypeEnum.adOpenForwardOnly,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal LockType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">ADODB.LockTypeEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">ADODB.LockTypeEnum.adLockReadOnly,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal CursorLocation&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">ADODB.CursorLocationEnum</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">-1</span>,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal DisconnectedRecordset&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False</td></tr></table></td><td class=\"PAfterParameters\">) As ADODB.Recordset</td></tr></table></div></div><div class=\"TTSummary\">Optimizes ADODB.Recordset.Open method</div></div>",33:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype33\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function OpenRecordsetCommandParam(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal CmdText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">ByVal CmdType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">ADODB.CommandTypeEnum,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal CmdParamDefs&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByRef RecordsAffected&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Options&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Long</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">-1</span>,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal DisconnectedRecordset&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Boolean</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">False</td></tr></table></td><td class=\"PAfterParameters\">) As ADODB.Recordset</td></tr></table></div></div><div class=\"TTSummary\">Open recordset using ADODB.Command</div></div>",34:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype34\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function LookupSql(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal SqlText&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Index&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHNumber\">0</span>&amp;,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Lookup ADODB.Recordset replacement function for DLookup (passing a SQL statement)</div></div>",35:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype35\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Lookup(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">ADODB.Recordset replacement function for DLookup</div></div>",36:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype36\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Count(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Long</td></tr></table></div></div><div class=\"TTSummary\">ADODB.Recordset replacement function for DCount</div></div>",37:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype37\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Max(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">ADODB.Recordset replacement function for DMax</div></div>",38:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype38\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Min(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">ADODB.Recordset replacement function for DMin</div></div>",39:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype39\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Sum(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Expr&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">ADODB.Recordset replacement function for DSum</div></div>",40:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype40\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Exists(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Domain&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal Criteria&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">vbNullString</td></tr></table></td><td class=\"PAfterParameters\">) As Boolean</td></tr></table></div></div><div class=\"TTSummary\">Check if record exists</div></div>",41:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype41\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function InsertIdentityReturn(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal InsertSql&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal IdentityTable&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">vbNullString</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Execute insert SQL statement and return last identity value (auto value)</div></div>",42:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype42\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function InsertValuesIdentityReturn(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal Source&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ParamArray InsertFields()&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Variant</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Create and execute insert SQL statement and return last identity value (auto value)</div></div>",43:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype43\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function ValueList(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal SqlSource&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ListConcatString&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\"><span class=\"SHString\">&quot;, &quot;</span>,</td></tr><tr><td class=\"PModifierQualifier first\">Optional ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Variant</td></tr></table></div></div><div class=\"TTSummary\">Generate concat string from recordset result</div></div>",44:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype44\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function LookupSqlValueCollection(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal SqlSource&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">String,</td><td></td><td class=\"last\"></td></tr><tr><td class=\"PModifierQualifier first\">Optional ByVal ValueIfNull&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName\">Variant</td><td class=\"PDefaultValueSeparator\">&nbsp;=&nbsp;</td><td class=\"PDefaultValue last\">Null</td></tr></table></td><td class=\"PAfterParameters\">) As Collection</td></tr></table></div></div><div class=\"TTSummary\">Generate collection from recordset result</div></div>"});