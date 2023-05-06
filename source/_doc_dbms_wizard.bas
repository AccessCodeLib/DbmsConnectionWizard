Attribute VB_Name = "_doc_dbms_wizard"
'#######################################################################################
'
' Doxygen-Gruppen
'
'
'---------------------------------------------------------------------------------------
'/**
' @defgroup grpWizard DBMS-Connection-Wizard
' Access-Add-In: DBMS-Connection-Wizard
'**/

'/**
' \defgroup grpCodeLib Code-Bibliothek
' Klassen und Module für die Verwendung in den Access-Anwendungen
'**/


'#######################################################################################
'
' Startseite
'
'---------------------------------------------------------------------------------------
'
'/** \mainpage DBMS-Connection-Wizard
' \image html pictures/wizardform.png
' <CENTER>\ref grpWizard &nbsp;&nbsp;&nbsp; | &nbsp;&nbsp;&nbsp; \ref grpCodeLib</CENTER>
'**/
'
'
'#######################################################################################
'
' Beschreibungen
'
'---------------------------------------------------------------------------------------
'/** \addtogroup grpWizard
' @{ **/
'
'/**
' \page wizard_install Installation
'
' Das Access-Add-In wird über den Add-In-Manger durch hinzufügen der Datei <em>DbmsConnectionWizard.mda</em> installiert.
' Falls Sie Vista oder Windows 7 verwenden und die Benutzerkontensteuerung (UAC) aktiviert haben,
' starten Sie zuvor Access als Administrator, da für die Registrierung der Assistenten Admin-Rechte notwendig sind.
' Sollten es Ihnen nicht möglich sein, Admin-Rechte zu erhalten, so können Sie das Add-In trotzdem verwenden.
' Es steht Ihnen dann allerdings nur das Menü-Add-In zur Verfügung.
'
'
' \page wizard_connection Verbindungsdaten einstellen
'
' Nach der Installation des Access-Add-In steht unter <em>Add-Ins</em> der Eintrag <em>DBMS Connection Wizard</em> zur Verfügung,
' mit dem der Assistent für die Verbindungskonfiguration gestartet werden kann.
'
' \section  sec_wizard_table Tabelle erstellen
' Falls die Tabelle \a usys_DbmsConnection im Frontend fehlt, wird diese vom Assistenten erstellt.
'
' \image html pictures/TabAnlegen.png "Meldung bei fehlender Tabelle"
'
' \section sec_wizard_connection Verbindungskonfiguration erstellen
'
' Mit der Schaltfläche "Neue Verbindung" kann eine neue Verbindungskonfiguration angelegt werden.
'
' \image html pictures/Neue_Verbindung_erstellen.png "Neue Verbindungskonfiguration erstellen"
'
' Diese muss einen eindeutigen Namen erhalten und einem DBMS zugeordnet werden.
'
' \image html pictures/Neue_Verbindung_erstellen_Eingabe.png "Eingeben der Verbindungskennung"
'
' \section  sec_wizard_parameter Zugriffsparameter einstellen
' Die Zugriffsparameter werden für jede Verbindungskennung gespeichert.
'
' \image html pictures/Verbindungsdaten_einstellen.png "Verbindungsdaten einstellen"
'
' \page wizard_parameter Beschreibung der Parameter
' \section sec_wizard_parameter_1 Aufbereitung der Zugriffszeichenfolge
' \par Verbindungsart
'     - ohne DSN ("dsn less")
'     Mit dieser Verbindungsart wird eine Zugriffszeichenfolge inkl. Treiber bzw. Provider verwende, die ohne gespeicherte Datenquellen auskommt.
'     Es gibt allerdings einige DBMS, bei denen eine DSN erforderlich ist.
'     - mit DSN
'     Die Zugriffszeichenfolge verweist auf eine gespeicherte Datenquelle, in der die Verbindungssparameter enthalten sind.
'     - benutzerdefiniert
'     Bei dieser Einstellung wird die Zugriffszeichenfolge nicht aus den Verbindugnsparametern zusammengesetzt sondern exakt jener Zeichenfolge verwendet,
'     die in den Eigenschaften ODBC und OLEDB unter benutzerdefinierter Connectionstring eingetragen werden.
' \section sec_wizard_parameter_2 Datenbank- u. Servereinstellungen
'
' \par Database
'  Der Datenbankname
'
' \par Server
'  Der Server bzw. die Server Instanz. z. B. (local)\SQLExpress
'
' \par Port
' Der zu verwendene Port. (Kann meist leer bleiben, wenn der Standardport verwendet wird.
'
' \par User
'  Die Benutzerkennung falls eine Serveranmeldung mit Benutzer und Passwort verwendet wird.
'  Bei Verwendung eines Login-Formulars wird im Frontend in diesem Feld der zuletzt angemeldete Benutzer gespeichert.
'
' \par Password
'  Das Passwort zur Benutzerkennung
'
' \par Windows-Autentifizierung
'  Falls statt einer Serverindentifizierung die Kennung von Windows verwendet werden soll. (Funktioniert nicht mit allen DBMS.)
'
' \par Login-Formular verwenden
'  Bei Benutzeranmeldung kann ein Login-formular verwendet werden, um das Passwort abzufragen. (Empfohlene Einstellung, wenn die Windows-Autentifizierung nicht verwendet wird.)
'
' \par DSN
'  Die DSN-Kennung, falls eine Verbindung über DSN aufgebaut werden soll
'
' \section sec_wizard_parameter_3 Provider / Driver
'
' \par OLEDB
' Der OLEDB-Provider. Als Standard wird der gängigste Provider zum jeweiligen DBMS voreingestellt. Diese Einstellung kann in der Tabelle usys_DbmsConnection_X geändert werden.
' Falls ein alternativer Provider verwendet wird, muss möglicherweise die Erstellung der Verbindugnszeichenfolge optimiert werden, falls dieser Provider spezielle Parameter benötigt.
'
' \par ODBC &nbsp;
' Der ODBC-Treiber für den DOBC-Zugriff analog zum OLDEDB-Provider.
'
' \subsection sec_wizard_parameter_4 Weitere Optionen
'
' \par OLEDB
' Zusätzliche OLEDB-Optionen. Dieser werden am Ende der OLEDB-Verbindungszeichenfolge angehängt.
'
' \par ODBC &nbsp;
' Zusätzliche ODBC-Optionen. Dieser werden am Ende der OLEDB-Verbindungszeichenfolge angehängt.
'
' \section sec_wizard_parameter_5 Benutzerdefinierter Connectionstring
' \par OLEDB
' Die OLEDB-Verbindungszeichenfolge bei benutzerdefinierter Konfiguration
'
' \par ODBC &nbsp;
' Die ODBC-Verbindungszeichenfolge bei benutzerdefinierter Konfiguration
'
' \section sec_wizard_parameter_6 Module und Klassen
' In diesem Abschnitt könnten die im Frontend benötigten Module aus dem Add-In kopiert werden.
'
'**/
'
' /** @} **/


'#######################################################################################
'
' Code-Lib
'
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Code-Lib: Datenzugriff
'---------------------------------------------------------------------------------------
'/** \addtogroup grpCodeLib
' @{ **/
'
' /**
' \page codelib_samples Beispiele für den Datenzugriff
'
' \section sec_codelib_samples Beispiele für den Datenzugriff
'
' \par DAO
'           \code DbCon.DAO.OpenRecordset("select ...") \endcode
'           \code DbCon.DAO.Execute "delete from ..." \endcode
'           \code OpenDaoRecordset("select ...") \endcode
'           \code DaoExecute "delete from ..." \endcode
'
' \par ADODB
'           \code DbCon.ADODB.OpenRecordset("select ...") \endcode
'           \code DbCon.ADODB.Execute "delete from ..." \endcode
'           \code OpenAdoRecordset("select ...") \endcode
'           \code AdoExecute "delete from ..." \endcode
'
' \par ODBC
'           \code DbCon.ODBC.OpenRecordset(....) \endcode
'           \code DbCon.ODBC.Execute "delete from ..." \endcode
'           \code OpenRecordsetDaoBE("select ...") \endcode
'           \code OpenRecordsetPT("select ...") \endcode
'           \code OdbcExecutePT "delete from ..." \endcode
'
' **/
'
'/** @} **/
