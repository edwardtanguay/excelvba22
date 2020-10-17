'TODO: get this to work
Public Sub qsql_GetData()
    Dim conn As Object, rst As Object
    Dim strSql As String

    Set conn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    'conn.Open "DRIVER=SQLite3 ODBC Driver;Database=C:\edward\nwo\excelvba22_files\data.sqlite"
    conn.Open "DRIVER=SQLite3 ODBC Driver;Database=C:\edward\filesForWeeklyBackup\LEARN2020\excelvba22_files\maindata.sqlite"
    
    strSql = "SELECT * FROM infos"
    
    rst.Open strSql, conn, 1, 1
    
    Worksheets("Main").Range("C16").CopyFromRecordset rst
    rst.Close
    
    Set rst = Nothing: Set conn = Nothing
End Sub
