'Provider=SQLOLEDB
'Dim conexion, cmd, resultadoCreate, resultadoSelect

' Crear objeto de conexión ADO
'Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexión
'conexion.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=AdventureWorks2022;Integrated Security=SSPI;"

' Abrir la conexión
'conexion.Open

' Crear objeto de comando para la creación de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Crear la tabla temporal
'cmd.CommandText = "SELECT SalesOrderID, ProductID, OrderQty INTO #MiTablaTemporal FROM Sales.SalesOrderDetail;"
'Set resultadoCreate = cmd.Execute

' Cerrar el Recordset de creación
'resultadoCreate.Close

' Crear un nuevo objeto de comando para la selección de datos de la tabla temporal
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

' Cerrar la conexión
'conexion.Close

' Liberar los objetos
'Set resultadoCreate = Nothing
'Set resultadoSelect = Nothing
'Set cmd = Nothing
'Set conexion = Nothing

'---------------------------------------------
'Provider=MSOLEDBSQL
'Dim conexion, cmd, resultadoCreate, resultadoSelect

' Crear objeto de conexión ADO
'Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexión con MSOLEDBSQL
'conexion.ConnectionString = "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=AdventureWorks2022;Integrated Security=SSPI;"

' Abrir la conexión
'conexion.Open

' Crear objeto de comando para la creación de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Crear la tabla temporal
'cmd.CommandText = "SELECT SalesOrderID, ProductID, OrderQty INTO #MiTablaTemporal FROM Sales.SalesOrderDetail;"
'Set resultadoCreate = cmd.Execute

' Cerrar el Recordset de creación
'resultadoCreate.Close

' Crear un nuevo objeto de comando para la selección de datos de la tabla temporal
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

' Cerrar la conexión
'conexion.Close

' Liberar los objetos
'Set resultadoCreate = Nothing
'Set resultadoSelect = Nothing
'Set cmd = Nothing
'Set conexion = Nothing

'----------------------------------

'Driver={SQL Server}
'Dim conexion, cmd, resultadoCreate, resultadoSelect

' Crear objeto de conexión ADO
'Set conexion = CreateObject("ADODB.Connection")

' Establecer la cadena de conexión ODBC con {SQL Server}
'conexion.ConnectionString = "Driver={SQL Server};Server=localhost;Database=AdventureWorks2022;Trusted_Connection=Yes;"

' Abrir la conexión
'conexion.Open

' Crear objeto de comando para la creación de la tabla temporal
'Set cmd = CreateObject("ADODB.Command")
'cmd.ActiveConnection = conexion
'cmd.CommandType = 1 ' Tipo de comando: Texto

' Crear la tabla temporal
'cmd.CommandText = "SELECT SalesOrderID, ProductID, OrderQty INTO ##MiTablaTemporal FROM Sales.SalesOrderDetail where SalesOrderID=43668;"
'Set resultadoCreate = cmd.Execute

' Cerrar el Recordset de creación
'resultadoCreate.Close

' Crear un nuevo objeto de comando para la selección de datos de la tabla temporal
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

' Cerrar la conexión
'conexion.Close

' Liberar los objetos
'Set resultadoCreate = Nothing
'Set resultadoSelect = Nothing
'Set cmd = Nothing
'Set conexion = Nothing

'-------------------------------
'Driver={SQL Server}
Dim conexion, objRS

 'Crear objeto de conexión ADO
Set conexion = CreateObject("ADODB.Connection")

 'Establecer la cadena de conexión ODBC con {SQL Server}
conexion.ConnectionString = "Driver={SQL Server};Server=MARCOSMG;Database=AdventureWorks2022;Trusted_Connection=Yes;"

 'Abrir la conexión
conexion.Open

 'Crear objeto de Recordset para la creación de la tabla temporal
Set objRS = CreateObject("ADODB.Recordset")
objRS.ActiveConnection = conexion
objRS.CursorType = 3 ' Tipo de cursor: AdOpenStatic (conjunto de registros estático)
objRS.LockType = 3 ' Tipo de bloqueo: AdLockOptimistic
'objRS.Open "SELECT SalesOrderID, ProductID, OrderQty INTO ##MiTablaTemporal FROM Sales.SalesOrderDetail WHERE SalesOrderID=43668;", conexion
'objRS.Open "SELECT * FROM ##MiTablaTemporal;", conexion
objRS.Open "SET NOCOUNT ON;;with cte as (" & _
		"select SalesOrderID, ProductID, OrderQty from Sales.SalesOrderDetail where SalesOrderID=43668" & _
		")" & _
		" select SalesOrderID, ProductID, OrderQty into #tabla from cte; select * from #tabla;", conexion
'objRS.Open "select * from ##tabla;", conexion
'objRS.Open "SET NOCOUNT ON; DECLARE @tabla TABLE (Col1 INT, Col2 INT, Col3 INT); " & _
'                  "WITH cte AS (" & _
'                  "    SELECT SalesOrderID, ProductID, OrderQty " & _
'                  "    FROM Sales.SalesOrderDetail " & _
'                  "    WHERE SalesOrderID = 43668" & _
'                  ") " & _
'                  "INSERT INTO @tabla (Col1, Col2, Col3) " & _
'                  "SELECT SalesOrderID, ProductID, OrderQty " & _
'                  "FROM cte; " & _
'                  "SELECT * FROM @tabla;", conexion

 'Imprimir los resultados en la consola
Do Until objRS.EOF
    WScript.Echo "Col1: " & objRS("SalesOrderID").Value
     'Otras columnas...
    WScript.Echo "---"
    objRS.MoveNext
Loop

 'Cerrar el Recordset
objRS.Close

 'Cerrar la conexión
conexion.Close

 'Liberar los objetos
Set objRS = Nothing
Set conexion = Nothing

'-----------------------------------------
