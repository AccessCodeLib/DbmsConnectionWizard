Attribute VB_Name = "_doc_dbms_wizard"
'---------------------------------------------------------------------------------------
' Section: ACLib DBMS Connection Wizard
'---------------------------------------------------------------------------------------
' Access add-in to configure DBMS connections and to import the ACLib DbConnection
' framework into a database.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Topic: Installation
'---------------------------------------------------------------------------------------
' Open DbmsConnectionWizard.accda (zip file from the releases) in Access and click
' "Install Add-in" in the install form.
' "Install Add-in as accde" installs a compiled version instead.
'
' Administrator rights are not required, because the entries are registered in the
' user profile (HKCU).
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Topic: Configure connection data
'---------------------------------------------------------------------------------------
' After the installation the entry "ACLib DBMS Connection Wizard" is available in the
' Add-ins menu. It starts the wizard for the connection configuration.
'
' If the table usys_DbmsConnection is missing in the frontend, the wizard creates it.
'
' A new connection configuration is created with the button "Neue Verbindung".
' It needs a unique name and has to be assigned to a DBMS.
' The access parameters are stored for each connection name.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Topic: Connection type
'---------------------------------------------------------------------------------------
' without DSN ("dsn less") - The connection string contains the driver or provider and needs no stored data source. Some DBMS require a DSN.
' with DSN - The connection string refers to a stored data source that contains the connection parameters.
' custom - The connection string is not built from the parameters, the string entered in the properties ODBC and OLEDB is used exactly as entered.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Topic: Database and server settings
'---------------------------------------------------------------------------------------
' Database - The database name.
' Server - The server or server instance, e.g. (local)\SQLExpress.
' Port - The port to use. Can usually stay empty if the default port is used.
' User - The user name if the server login uses user and password. If a login form is used, the frontend stores the last logged-in user in this field.
' Password - The password of the user.
' Windows authentication - Use the Windows login instead of a server login. Does not work with every DBMS.
' Use login form - Asks for the password in a login form. Recommended if Windows authentication is not used.
' DSN - The DSN name if the connection is built via DSN.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Topic: Provider and driver
'---------------------------------------------------------------------------------------
' OLEDB - The OLEDB provider. The most common provider of the DBMS is preset. It can be changed in the table usys_DbmsConnection_X.
'   If an alternative provider needs special parameters, building the connection string may have to be adjusted.
' ODBC - The ODBC driver, analogous to the OLEDB provider.
'
' Additional options are appended to the end of the connection string.
' OLEDB options - Additional OLEDB options.
' ODBC options - Additional ODBC options.
'
' Custom connection string:
' OLEDB - The OLEDB connection string of a custom configuration.
' ODBC - The ODBC connection string of a custom configuration.
'
' Modules and classes - The modules and classes needed in the frontend can be copied from the add-in.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Topic: Data access examples
'---------------------------------------------------------------------------------------
' DAO:
'
' (start code)
' DbCon.DAO.OpenRecordset("select ...")
' DbCon.DAO.Execute "delete from ..."
' OpenDaoRecordset("select ...")
' DaoExecute "delete from ..."
' (end code)
'
' ADODB:
'
' (start code)
' DbCon.ADODB.OpenRecordset("select ...")
' DbCon.ADODB.Execute "delete from ..."
' OpenAdoRecordset("select ...")
' AdoExecute "delete from ..."
' (end code)
'
' ODBC:
'
' (start code)
' DbCon.ODBC.OpenRecordset(....)
' DbCon.ODBC.Execute "delete from ..."
' OpenRecordsetDaoBE("select ...")
' OpenRecordsetPT("select ...")
' OdbcExecutePT "delete from ..."
' (end code)
'---------------------------------------------------------------------------------------
