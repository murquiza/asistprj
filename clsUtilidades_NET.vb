Public Class clsUtilidades_NET

    Public Function AddPath(ByVal sFile As String, Optional ByVal vPath As Object = Nothing) As String

        On Error GoTo AddPath_Err

        ' Declaraciones
        '
        Dim sTmp As String
        Dim sPath As String

        ' Añadir el path sino lo tiene
        ' Si no se especifica el path a añadir, usar App.Path
        ' Si ya incluye el path, devolver el valor actual
        '
        If InStr(sFile, "\") Then
            AddPath = sFile
        Else
            'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
            If IsNothing(vPath) Then
                'JCLopez_i
                'sPath = VB6.GetPath
                sPath = System.AppDomain.CurrentDomain.BaseDirectory
                'JCLopez_f
            Else
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto vPath. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                sPath = Trim(CStr(vPath))
            End If
            'si se indica con cadena vacía...
            'añadir el path actual
            If Len(sPath) = 0 Then
                sPath = CurDir()
            End If
            AddPath = AddBackSlash(sPath) & sFile
        End If
        Exit Function

AddPath_Err:
        AddPath = Str(Err.Number)
    End Function

    ' Esta función quita la barra de directorio del final
    ' Devuelve una cadena con el path modificado
    '
    Public Function QuitarBackSlash(ByVal sPath As String) As String

        On Error GoTo QuitarBackSlash_Err

        If Right(sPath, 1) = "\" Then
            sPath = Left(sPath, Len(sPath) - 1)
        End If
        QuitarBackSlash = sPath
        Exit Function

QuitarBackSlash_Err:
        QuitarBackSlash = Str(Err.Number)
    End Function

    ' Devuelve el directorio origen del path especificado
    '
    Public Function GetDir(ByVal sPath As String) As String

        On Error GoTo GetDir_Err

        ' Declaraciones
        '
        Dim i As Integer

        GetDir = sPath
        For i = Len(sPath) To 1 Step -1
            If Mid(sPath, i, 1) = "\" Then
                GetDir = Left(sPath, i - 1)
                Exit For
            End If
        Next
        Exit Function

GetDir_Err:
        GetDir = Str(Err.Number)
    End Function

    ' Devuelve sólo el Nombre y extensión del fichero indicado
    ' Si se indica False en ConExt no devuelve la extensión
    '
    Public Function NameFromFileName(ByVal sFileName As String, Optional ByVal ConExt As Boolean = True) As String

        On Error GoTo NameFromFileName_Err

        ' Declaraciones
        '
        Dim sPath As String
        Dim sName As String
        Dim i As Integer

        ' Obtenemos por separado la ruta, el fichero y la extensión
        '
        sPath = " "
        sName = " "
        SplitPath(sFileName, sPath, sName)
        If ConExt = False Then
            i = InStrRev(sName, ".")
            If i Then
                sName = Left(sName, i - 1)
            End If
        End If
        NameFromFileName = sName
        Exit Function

NameFromFileName_Err:
        NameFromFileName = Str(Err.Number)
    End Function

    ' Esta función devuelve sólo el Path del fichero indicado

    Public Function PathFromFileName(ByVal sFileName As String) As String

        On Error GoTo PathFromFileName_Err

        ' Declaraciones
        '
        Dim sPath As String

        SplitPath(sFileName, sPath)
        PathFromFileName = sPath
        Exit Function

PathFromFileName_Err:
        PathFromFileName = Str(Err.Number)
    End Function

    ' Esta función devuelve sólo la extensión del fichero indicado
    '
    Public Function ExtFromFileName(ByVal sFileName As String) As String

        On Error GoTo ExtFromFileName_Err

        ' Declaraciones
        '
        Dim sPath As String
        Dim sName As String
        Dim sExt As String

        SplitPath(sFileName, sPath, sName, sExt)
        ExtFromFileName = sExt
        Exit Function

ExtFromFileName_Err:
        ExtFromFileName = Str(Err.Number)
    End Function

    ' Esta función devuelve el path de la aplicación con la barra al final
    ' o sin ella, según se especifique en el parámetro 'conBackSlash'
    '
    Public Function AppPath(Optional ByVal conBackSlash As Boolean = True) As String

        On Error GoTo AppPath_Err

        ' Declaraciones
        '
        Dim pathAplicacion As String

        'JCLopez_i
        'pathAplicacion = VB6.GetPath
        pathAplicacion = System.AppDomain.CurrentDomain.BaseDirectory
        'JCLopez_f

        If conBackSlash Then
            If Right(pathAplicacion, 1) <> "\" Then
                pathAplicacion = pathAplicacion & "\"
            End If
        Else
            If Right(pathAplicacion, 1) = "\" Then
                pathAplicacion = Left(pathAplicacion, Len(pathAplicacion) - 1)
            End If
        End If
        AppPath = pathAplicacion
        Exit Function

AppPath_Err:
        AppPath = Str(Err.Number)
    End Function

    ' Divide el nombre recibido en la ruta, nombre y extensión
    ' Esta rutina aceptará los siguientes parámetros:
    '       sTodo      Valor de entrada con la ruta completa
    '
    '   Devolverá la información en:
    '       sPath      Ruta completa, incluida la unidad
    '       vNombre    Nombre del archivo incluida la extensión
    '       vExt       Extensión del archivo
    '
    Public Sub SplitPath(ByVal sTodo As String, ByRef sPath As String, Optional ByRef vNombre As Object = Nothing, Optional ByRef vExt As Object = Nothing)

        On Error GoTo SplitPath_Err

        ' Declaraciones
        '
        Dim bNombre As Boolean
        Dim i As Short

        'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
        If Not IsNothing(vNombre) Then
            bNombre = True
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto vNombre. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            vNombre = sTodo
        End If

        'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
        If Not IsNothing(vExt) Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto vExt. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            vExt = ""
            i = InStr(sTodo, ".")
            If i Then
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto vExt. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                vExt = Mid(sTodo, i + 1)
            End If
        End If

        sPath = ""

        'Asignar el path
        For i = Len(sTodo) To 1 Step -1
            If Mid(sTodo, i, 1) = "\" Then
                sPath = Left(sTodo, i - 1)
                'Si hay que devolver el nombre
                If bNombre Then
                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto vNombre. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    vNombre = Mid(sTodo, i + 1)
                End If
                Exit For
            End If
        Next
        Exit Sub

SplitPath_Err:
        sPath = Str(Err.Number)
    End Sub

    ' Esta función añade la barra de directorio al path especificado
    ' Devuelve una cadena
    '
    Public Function AddBackSlash(ByVal sPath As String) As String

        On Error GoTo AddBackSlash_Err

        If Len(sPath) Then
            If Right(sPath, 1) <> "\" Then
                sPath = sPath & "\"
            End If
        End If
        AddBackSlash = sPath
        Exit Function

AddBackSlash_Err:
        AddBackSlash = Str(Err.Number)
    End Function

    ' Esta función convierte un número en formato texto a fomato numerico
    ' con los decimales especificados en el parámetro 'Decimales'
    ' Devuelve un variant
    '
    Public Function TextToNumeric(ByRef Texto As String, Optional ByRef Decimales As Short = 0) As Object

        On Error GoTo TextToNumeric_Err

        ' Declaraciones
        '
        Dim strEntero As String
        Dim strDecimal As String
        Dim intPosicion As Short
        Dim intContador As Short

        If Not IsNumeric(Texto) Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto TextToNumeric. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            TextToNumeric = 0
            Exit Function
        End If

        If Decimales > 0 Then
            strEntero = Texto
            strDecimal = "1"
            For intContador = 1 To Decimales
                strDecimal = strDecimal & "0"
            Next
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto TextToNumeric. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            TextToNumeric = CDbl(CInt(strEntero) / CInt(strDecimal))
        Else
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto TextToNumeric. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            TextToNumeric = CInt(Texto)
        End If
        Exit Function

TextToNumeric_Err:
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto TextToNumeric. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        TextToNumeric = Err.Number
    End Function

    ' Procedure:    AnalizarNombreFichero
    ' Objetivo:     Analiza si el nombre fichero especificado contiene
    '               caracteres especiales no válidos
    ' Parametros:   Fichero =  Nombre y ruta del fichero donde exportar, original
    ' Retorno:      Booleano
    '
    Public Function ValidaNombreFichero(ByRef Fichero As String) As Boolean

        On Error GoTo ValidaNombreFichero_Err

        If InStr(1, Fichero, "*", CompareMethod.Text) Or InStr(1, Fichero, " ", CompareMethod.Text) Then

            ValidaNombreFichero = False
            Exit Function
        End If

        ValidaNombreFichero = True
        Exit Function

ValidaNombreFichero_Err:
        ValidaNombreFichero = False
    End Function

    ' Esta función devuelve el numero de veces que un caracter se encuentra
    ' repetido en una cadena
    '
    Public Function ItemCaracter(ByRef Cadena As String, ByRef Caracter As String) As Short

        On Error GoTo ItemCaracter_Err

        ' Declaraciones
        '
        Dim i As Short
        Dim Contador As Short

        Contador = 0

        ' Inicamos bucle de buqueda
        '
        For i = 1 To Len(Cadena)
            If Mid(Cadena, i, 1) = Caracter Then
                Contador = Contador + 1
            End If
        Next i

        ItemCaracter = Contador
        Exit Function

ItemCaracter_Err:
        ItemCaracter = -1
    End Function

    ' Procedure:    ContarRegistros
    ' Objetivo:     Cuenta el número de filas de un objeto RecordSet
    ' Parametros:   Objeto RecordSet del que se han de contar los registros
    ' Retorno:      Long con el número de registros contenidos

    Public Function ContarRegistros(ByVal objRs As ADODB.Recordset) As Integer

        On Error GoTo ContarRegistro_Err

        ' Declaraciones
        '
        Dim i As Short ' Contador para bucles

        ' Inicializamos el puntero y el contador
        '
        objRs.MoveFirst()
        ContarRegistros = 0

        ' Iniciammos bucle de lectura
        '
        Do While Not objRs.EOF
            objRs.MoveNext()
            ContarRegistros = ContarRegistros + 1
        Loop

        Exit Function

ContarRegistro_Err:
        ContarRegistros = -1
    End Function

    ' Esta función devuelve el número de registros distintos existentes
    ' en un Recordset por el campo pasado como parámetro en 'Campo'
    ' Problemas con determinadas funciones del RecordSet del proveedor OLEDB
    ' obligan a ordenar previamente a la llamada a esta función el objeto recordset
    ' por el campo por el que se quiera distinguir la busqueda
    '
    Public Function ContarRegistrosDistinct(ByRef Campo As String, ByVal objRs As ADODB.Recordset, Optional ByRef Numregis As Integer = 0) As Integer

        On Error GoTo ContarRegistrosDistinct_Err

        ' Declaraciones
        '
        Dim i As Object
        Dim x As Short ' Contador para bucles
        Dim TotalRegistros As Integer ' Total de registros del RecordSet
        Dim colReg As Collection
        Dim Control As Boolean ' Indica si se ha encontrado el registro en la colección

        colReg = New Collection

        ' Entramos en un bucle en el que recorremos todo el RecordSet
        ' y en el que vamos contando todos los valores no repetidos
        '
        If Numregis = 0 Then
            TotalRegistros = ContarRegistros(objRs)
        Else
            TotalRegistros = Numregis
        End If
        objRs.MoveFirst()
        colReg.Add(objRs.Fields(Campo).Value, Str(1))
        For i = 1 To TotalRegistros
            For x = 1 To colReg.Count()
                Control = False
                If objRs.Fields(Campo).Value = colReg.Item(x) Then
                    Control = True
                    Exit For
                End If
            Next x
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto i. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            If Not Control Then colReg.Add(objRs.Fields(Campo).Value, Str(i))
            objRs.MoveNext()
        Next i
        ContarRegistrosDistinct = colReg.Count()
        Exit Function

ContarRegistrosDistinct_Err:
        ContarRegistrosDistinct = 0
    End Function
    ' Esta función devuelve el código de empleado a partir de su clave de acceso
    '
    Public Function CodUser(ByRef sClave As String) As String

        On Error GoTo CodUser_Err

        ' Declaraciones
        '
        Dim strsql As String

        ' Si el objeto RecordSet esta abierto lo cerramos
        '
        If claseBDLibrerias.BDAuxRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDLibrerias.BDAuxRecord.Close()

        strsql = "Select Num_Empl From Empleado Where Clave = '" & sClave & "'"
        claseBDLibrerias.BDAuxRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDLibrerias.BDAuxRecord.EOF Then
            Err.Raise(1)
        Else
            CodUser = claseBDLibrerias.BDAuxRecord.Fields("Num_Empl").Value
        End If
        claseBDLibrerias.BDAuxRecord.Close()
        Exit Function

CodUser_Err:
        CodUser = ""
    End Function

    ' Proposito :  Devuelve una Fecha formateada para introducir dentro de
    '              una sentencia SQL, para comparaciones.
    ' Parametros : [dtmFecha] Fecha a formatear, debe ser dato tipo fecha
    '              [blnComodin] Opcional, incluir comodines en la fecha
    '              [blnHora] Opcional, incluir la hora en la fecha
    ' Retorno :    Devuelve una cadena con la fecha formateada
    '

    Public Function FormatoFechaSQL(ByRef dtmFecha As Date, Optional ByRef blnComodin As Boolean = True, Optional ByRef blnHora As Boolean = True) As String

        ' Declaraciones
        '
        Dim strResultadoFecha As String
        Dim strFecha As String
        Dim strHora As String

        strResultadoFecha = ""
        'JCLopez_i
        'strFecha = Month(dtmFecha) & "/" & VB.Day(dtmFecha) & "/" & Year(dtmFecha)
        strFecha = dtmFecha.Month & "/" & dtmFecha.Day & "/" & dtmFecha.Year
        'strHora = Hour(dtmFecha) & ":" & Minute(dtmFecha) & ":" & Second(dtmFecha)
        strHora = dtmFecha.Hour & ":" & dtmFecha.Minute & ":" & dtmFecha.Second
        'JCLopez_f
        If blnComodin Then strResultadoFecha = "#"
        strResultadoFecha = strResultadoFecha & strFecha
        If blnHora Then strResultadoFecha = strResultadoFecha & Space(1) & strHora
        If blnComodin Then strResultadoFecha = strResultadoFecha & "#"
        FormatoFechaSQL = strResultadoFecha

    End Function

End Class
