Attribute VB_Name = "Module1"
Public Function ConnectStringAc() As String
Dim strServerName As String
Dim strDatabaseName As String
Dim strPort As String
Dim strUserName As String
Dim strPassword As String

    'Change to IP Address if not on local machine
    'Make sure that you give permission to log into the
    'server from this address
    'See Adding New User Accounts to MySQL
    'Make sure that you d/l and install the MySQL Connector/ODBC 3.51 Driver

strServerName = "46.16.188.16"
strPort = "3306"
strDatabaseName = "mylighth_lightmypimsdb"
strUserName = "mylighth_limpms1"
strPassword = "Ra=CtmQ2@5&I"

ConnectStringAc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=app_info.mdb;Persist Security Info=False"

End Function

