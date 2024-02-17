'Provider=SQLOLEDB
'Dim conexion, cmd, resultadoCreate, resultadoSelect

' Crear objeto de conexi�n ADO
'Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexi�n
'conexion.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=AdventureWorks2022;Integrated Security=SSPI;"

' Abrir la conexi�n
'conexion.Open

' Crear objeto de comando para la creaci�n de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Crear la tabla temporal
'cmd.CommandText = "SELECT SalesOrderID, ProductID, OrderQty INTO #MiTablaTemporal FROM Sales.SalesOrderDetail;"
'Set resultadoCreate = cmd.Execute

' Cerrar el Recordset de creaci�n
'resultadoCreate.Close

' Crear un nuevo objeto de comando para la selecci�n de datos de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Seleccionar datos de la tabla temporal
'cmd.CommandText = "SELECT * FROM #MiTablaTemporal;"
'Set resultadoSelect = cmd.Execute

' Imprimir los resultados en la consola
'Do Until resultadoSelect.EOF
'    WScript.Echo "Col1: " & resultadoSelect("SalesOrderID").Value
    'WScript.Echo "Col2: " & resultadoSelect("ProductID").Value
'    WScript.Echo "---"
'    resultadoSelect.MoveNext
'Loop

' Cerrar la conexi�n
'conexion.Close

' Liberar los objetos
'Set resultadoCreate = Nothing
'Set resultadoSelect = Nothing
'Set cmd = Nothing
'Set conexion = Nothing

'---------------------------------------------
'Provider=MSOLEDBSQL
'Dim conexion, cmd, resultadoCreate, resultadoSelect

' Crear objeto de conexi�n ADO
'Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexi�n con MSOLEDBSQL
'conexion.ConnectionString = "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=AdventureWorks2022;Integrated Security=SSPI;"

' Abrir la conexi�n
'conexion.Open

' Crear objeto de comando para la creaci�n de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Crear la tabla temporal
'cmd.CommandText = "SELECT SalesOrderID, ProductID, OrderQty INTO #MiTablaTemporal FROM Sales.SalesOrderDetail;"
'Set resultadoCreate = cmd.Execute

' Cerrar el Recordset de creaci�n
'resultadoCreate.Close

' Crear un nuevo objeto de comando para la selecci�n de datos de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Seleccionar datos de la tabla temporal
'cmd.CommandText = "SELECT * FROM #MiTablaTemporal;"
'Set resultadoSelect = cmd.Execute

' Imprimir los resultados en la consola
'Do Until resultadoSelect.EOF
'    WScript.Echo "Col1: " & resultadoSelect("SalesOrderID").Value
    'WScript.Echo "Col2: " & resultadoSelect("ProductID").Value
'    WScript.Echo "---"
'    resultadoSelect.MoveNext
'Loop

' Cerrar la conexi�n
'conexion.Close

' Liberar los objetos
'Set resultadoCreate = Nothing
'Set resultadoSelect = Nothing
'Set cmd = Nothing
'Set conexion = Nothing

'----------------------------------

'Driver={SQL Server}
'Dim conexion, cmd, resultadoCreate, resultadoSelect

' Crear objeto de conexi�n ADO
'Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexi�n ODBC con {SQL Server}
'conexion.ConnectionString = "Driver={SQL Server};Server=localhost;Database=AdventureWorks2022;Trusted_Connection=Yes;"

' Abrir la conexi�n
'conexion.Open

' Crear objeto de comando para la creaci�n de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Crear la tabla temporal
'cmd.CommandText = "SELECT SalesOrderID, ProductID, OrderQty INTO ##MiTablaTemporal FROM Sales.SalesOrderDetail where SalesOrderID=43668;"
'Set resultadoCreate = cmd.Execute

' Cerrar el Recordset de creaci�n
'resultadoCreate.Close

' Crear un nuevo objeto de comando para la selecci�n de datos de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Seleccionar datos de la tabla temporal
'cmd.CommandText = "SELECT * FROM ##MiTablaTemporal;"
'Set resultadoSelect = cmd.Execute

' Imprimir los resultados en la consola
'Do Until resultadoSelect.EOF
'    WScript.Echo "Col1: " & resultadoSelect("SalesOrderID").Value
    'WScript.Echo "Col2: " & resultadoSelect("ProductID").Value
'    WScript.Echo "---"
'    resultadoSelect.MoveNext
'Loop

' Cerrar la conexi�n
'conexion.Close

' Liberar los objetos
'Set resultadoCreate = Nothing
'Set resultadoSelect = Nothing
'Set cmd = Nothing
'Set conexion = Nothing

'-------------------------------
'Driver={SQL Server}
Dim conexion, resultadoCreate, resultadoSelect

' Crear objeto de conexi�n ADO
Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexi�n ODBC con {SQL Server}
conexion.ConnectionString = "Driver={SQL Server};Server=localhost;Database=AdventureWorks2022;Trusted_Connection=Yes;"

' Abrir la conexi�n
conexion.Open

' Crear objeto de Recordset para la creaci�n de la tabla temporal
Set objRS = CreateObject("ADODB.Recordset")
objRS.ActiveConnection = conexion
objRS.CursorType = 3 ' Tipo de cursor: AdOpenStatic (conjunto de registros est�tico)
objRS.LockType = 3 ' Tipo de bloqueo: AdLockOptimistic
objRS.Open "SELECT SalesOrderID, ProductID, OrderQty INTO ##MiTablaTemporal FROM Sales.SalesOrderDetail WHERE SalesOrderID=43668;", conexion
objRS.Open "SELECT * FROM ##MiTablaTemporal;", conexion

' Imprimir los resultados en la consola
Do Until objRS.EOF
    WScript.Echo "Col1: " & objRS("SalesOrderID").Value
    ' Otras columnas...
    WScript.Echo "---"
    objRS.MoveNext
Loop

' Cerrar el Recordset
objRS.Close

' Cerrar la conexi�n
conexion.Close

' Liberar los objetos
Set resultadoCreate = Nothing
Set resultadoSelect = Nothing
Set conexion = Nothing
