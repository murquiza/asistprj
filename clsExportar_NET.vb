Public Class clsExportar_NET

    ' Variables locales para almacenar los valores de las propiedades
    '
    Private mvarFecEfe As Date ' Variable local para la propiedad FecEfe
    Private mvarArchivo As String ' Variable local para la propiedad Archivo
    Private mvarTipoExp As String ' Variable local para la propiedad TipoExp
    Public lngPolizas As Double ' Número de pólizas seleccionadas

    ' Asigna la parte común del nombre del archivo a grabar con los datos
    '

    ' Lee la parte cpmún del nombre de archivo especificado
    '
    Public Property Archivo() As String
        Get
            Archivo = mvarArchivo
        End Get
        Set(ByVal Value As String)
            mvarArchivo = Value
        End Set
    End Property

    ' Asigna la fecha de efecto de exportación
    '

    ' Lee la fecha de efecto de la exportación
    '
    Public Property FecEfe() As Date
        Get
            FecEfe = mvarFecEfe
        End Get
        Set(ByVal Value As Date)
            mvarFecEfe = Value
        End Set
    End Property

    ' Asigna el tipo de exportación a realizar
    '

    ' Lee el tipo de exportación a realizar
    '
    Public Property TipoExportacion() As String
        Get
            TipoExportacion = mvarTipoExp
        End Get
        Set(ByVal Value As String)
            mvarTipoExp = Value
        End Set
    End Property

    ' Procedure:  Exportacion
    ' Objetivo:   Proceso de exportacion de datos polizas a ficheros de Texto
    ' Parametros: Fecha = Fecha de ejecución proceso
    '
    Private Function GrabaFichero(ByRef VectorDatos As String, ByRef Canal As Short) As Boolean

        On Error GoTo GrabaFichero_Err

        ' Grabamos en el fichero especificado el vector de datos
        '
        PrintLine(Canal, VectorDatos)
        GrabaFichero = True

        Exit Function

GrabaFichero_Err:
        GrabaFichero = False
    End Function

    ' Esta función abre un fichero secuencial y devuelve el número de canal
    ' por el que ha sido abierto
    '
    Private Function AbreFichero(ByVal Fichero As String) As Short

        On Error GoTo Abrefichero_Err

        ' Declaraciones
        '
        Dim Canal As Short

        ' Obtenemos el numero de canal por el que abriremos el fichero
        '
        Canal = FreeFile()

        ' Apertura del fichero
        '
        FileOpen(Canal, Fichero, OpenMode.Output, OpenAccess.Write)
        AbreFichero = Canal
        Exit Function

Abrefichero_Err:
        AbreFichero = -1
    End Function

    ' Esta función cierra el fichero abierto por el canal especificado
    '
    Private Function CierraFichero(ByRef Canal As Short) As Boolean

        On Error GoTo CierraFichero_Err

        FileClose(Canal)
        CierraFichero = True

        Exit Function

CierraFichero_Err:
        CierraFichero = False
    End Function

    ' Procedure:    Exportacion_Datos_Personales
    ' Objetivo:     Exporta datos de cabecera de las polizas al fichero de texto
    '               que contendra las polizas/datos personales
    ' Parametros:   Poliza = Numero de poliza a tratar
    '               Fichero = ID Fichero de Texto de exportación
    '               Movimiento = Tipo de movimiento (A/B/M)
    '               FechaHist = Fecha de cuando se ha producido el movimiento
    '
    Public Function DatosCabecera() As Boolean

        On Error GoTo Exportacion_Datos_Err

        DatosCabecera = True


        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4016"

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        '

        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = SelectFich1
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich1. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CanalFichero = AbreFichero(Fich1)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Cabecera"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF
            For i = 0 To objSiniestros.NumeroCampos
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                CadenaTRans = Space(objSiniestros.LongitudCampos.Item(Str(i)))
                ' Opciones de alineación y formateo
                '
                Select Case objSiniestros.Alineacion.Item(Str(i))
                    Case "IZ" ' Alineación Izquierda
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.Relleno(Str$(i)). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            ''/*MUL INI */
                            'If Not IsDBNull(objSiniestros.Relleno.Item(Str(i))) And objSiniestros.Relleno.Item(Str(i)) > 0 Then
                            If Not IsDBNull(objSiniestros.Relleno.Item(Str(i))) Then
                                If objSiniestros.Relleno.Item(Str(i)) > 0 Then
                                    ''/*MUL FIN */
                                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.Relleno(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                    CadenaTRans = Replace(CadenaTRans, " ", Chr(objSiniestros.Relleno.Item(Str(i))))
                                End If
                            End If
                        End If
                    Case "DR" ' Alineación derecha
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = RSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            End If
                    Case "NU" ' Alineación campo númerico
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = Rslocal.Fields(i).Value
                                ' Comprobamos si hay decimales
                                '
                                If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                                    ' Si los hay eliminamos la coma
                                    CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                                Else
                                    ' Si no los hay añadimos dos ceros al final
                                    CadenaTRans = CadenaTRans & "00"
                                End If
                                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                ' Cambiamos lo espacios en blanco por ceros
                                CadenaTRans = Replace(CadenaTRans, " ", "0")
                            End If
                    Case "DT" ' Alineación y formato campo fecha
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "yyyyMMdd")
                            End If
                    Case "DS" ' Alineación y formato campo fecha española
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = Format(Rslocal.Fields(i).Value, "DDMMYYYY")
                            End If
                    Case "ES" ' Campo alfanumerico pero rellenando con ceros
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)))
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(Str$(i)). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                CadenaTRans = Mid(Rslocal.Fields(i).Value, 1, objSiniestros.LongitudCampos.Item(Str(i)))
                                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                CadenaTRans = Replace(CadenaTRans, " ", "0")
                            End If
                    Case Else ' Otros
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            End If
                End Select

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i
            ' Grabamos el registro de exportación ya formateado en el fichero
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Pólizas procesadas: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosCabecera = True
        Else
            Err.Raise(1)
        End If
        Exit Function

Exportacion_Datos_Err:
        FileClose(CanalFichero)
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich1. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Kill(Fich1)
        DatosCabecera = False
    End Function

    ' Procedure: Datos Garantías
    ' Objetivo:  Confecciona un vector con todos los datos de las garantías y coberturas
    '            de las polizas a exportar
    '
    Public Function DatosGarantias() As Boolean

        On Error GoTo Datosgarantias_Err

        DatosGarantias = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Double ' Número de registro del recordset leido
        Dim CdoCte As String ' Indica si el capital de la garantía es de Continente o Contenido
        Dim i As Integer ' Contador para bucles
        Dim ImporteCapital As Double ' Importe del capital de la garantia
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim PunteroCadena As Integer ' Para el tratamiento de cadenas indica la posicon dentro de la cadenaç
        ' de un determinado caracter
        Dim tmpCadena1 As String
        Dim tmpCadena2 As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4017"

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        '

        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = SelectFich2
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        ' Contamos el número de garantías para la representación gárafica d la
        ' barra de progreso y el cálculo del porcentaje realizado
        '
        lngResult = ContarRegistros(Rslocal)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich2. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CanalFichero = AbreFichero(Fich2)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngResult
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar lo
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Garantias"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF
            For i = 0 To objSiniestros.NumeroCampos
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                CadenaTRans = Space(objSiniestros.LongitudCampos.Item(Str(i)))

                ' Opciones de alineación y formateo
                '
                Select Case objSiniestros.Alineacion.Item(Str(i))
                    Case "IZ" ' Alineación Izquierda
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                        End If
                    Case "DR" ' Alineación derecha
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = RSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                        End If
                    Case "NU" ' Alineación campo númerico
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            If Codcia = "A" Then
                                ' Cuando el campo indica el importe de capital de
                                ' una garantia, el primer caracter indica a que tipo
                                ' de capital pertenece.
                                ' T = Continente
                                ' D = Contenido
                                ' O = Otros ( Capital propio )
                                '
                                If Left(Rslocal.Fields(i).Value, 1) = "T" Or Left(Rslocal.Fields(i).Value, 1) = "D" Or Left(Rslocal.Fields(i).Value, 1) = "O" Then
                                    CdoCte = Left(Rslocal.Fields(i).Value, 1)
                                    ImporteCapital = CDbl(Mid(Rslocal.Fields(i).Value, 2, Len(Rslocal.Fields(i).Value) - 1))
                                    ' Si el capital es Continente
                                    If Rslocal.Fields(i).Name = "Continente" Then
                                        If CdoCte = "T" Then
                                            CadenaTRans = CStr(ImporteCapital)
                                            ' Comprobamos si hay decimales
                                            If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                                                ' Si los hay eliminamos la coma
                                                CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                                            End If
                                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i))) & CadenaTRans
                                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                            CadenaTRans = Right(CadenaTRans, objSiniestros.LongitudCampos.Item(Str(i)))
                                        Else
                                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)))
                                        End If
                                    End If
                                    ' Si el capital es Contenido
                                    If Rslocal.Fields(i).Name = "Contenido" Then
                                        If CdoCte = "D" Then
                                            CadenaTRans = CStr(ImporteCapital)
                                            ' Comprobamos si hay decimales
                                            If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                                                ' Si los hay eliminamos la coma
                                                CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                                            End If
                                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i))) & CadenaTRans
                                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                            CadenaTRans = Right(CadenaTRans, objSiniestros.LongitudCampos.Item(Str(i)))
                                        Else
                                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)))
                                        End If
                                    End If
                                Else
                                    CadenaTRans = Rslocal.Fields(i).Value
                                    ' Comprobamos si hay decimales
                                    If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                                        ' Si los hay eliminamos la coma
                                        CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                                    Else
                                        ' Si no los hay añadimos dos ceros al final
                                        CadenaTRans = CadenaTRans & "00"
                                    End If
                                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                    CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                    ' Cambiamos lo espacios en blanco por ceros
                                    CadenaTRans = Replace(CadenaTRans, " ", "0")
                                End If
                            Else
                                If Rslocal.Fields(i).Value > 0 Then
                                    CadenaTRans = Rslocal.Fields(i).Value

                                    ' Comprobamos si hay decimales
                                    ' Si los hay generamos dos cadenas
                                    ' una con la parte entera y otra
                                    ' con la parte decimal, si no
                                    ' añadimos dos ceros a la derecha
                                    '
                                    PunteroCadena = InStr(1, CadenaTRans, ".")
                                    If PunteroCadena = 0 Then
                                        PunteroCadena = InStr(1, CadenaTRans, ",")
                                    End If
                                    If PunteroCadena > 0 Then
                                        tmpCadena1 = ""
                                        tmpCadena2 = ""
                                        tmpCadena1 = Mid(CadenaTRans, 1, PunteroCadena - 1)
                                        tmpCadena2 = Mid(CadenaTRans, PunteroCadena + 1, Len(CadenaTRans) - Len(tmpCadena1))
                                        If Len(tmpCadena2) < 2 Then
                                            tmpCadena2 = tmpCadena2 & "0"
                                        End If
                                        CadenaTRans = tmpCadena1 & tmpCadena2
                                    Else
                                        CadenaTRans = CadenaTRans & "00"
                                    End If
                                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                    CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                    ' Cambiamos lo espacios en blanco por ceros
                                    'CadenaTRans = Replace(CadenaTRans, " ", "0")
                                Else
                                    CadenaTRans = CStr(0)
                                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                    CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                End If
                            End If
                        End If
                    Case "DT" ' Alineación y formato campo fecha
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            'CadenaTRans = VB6.Format(Rslocal.Fields(i).Value, "YYYYMMDD")
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "yyyyMMdd")
                        End If
                    Case "DS" ' Alineación y formato campo fecha española
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "DDMMYYYY")
                        End If
                    Case "ES" ' Campo alfanumerico pero rellenando con ceros
                        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                        CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)))
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Rslocal.Fields(i).Value
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                        End If
                    Case Else ' Otros
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                        End If
                End Select

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i
            ' Grabamos el registro de exportación ya formateado en el fichero
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Garantias procesadas: " & numreg
                Call ActualizarPorcentaje(lngResult)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosGarantias = True
        Else
            Err.Raise(1)
        End If
        Exit Function

Datosgarantias_Err:
        FileClose(CanalFichero)
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich1. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Kill(Fich1)
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich2. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Kill(Fich2)
        DatosGarantias = False
    End Function

    ' Exporta los datos de las asignaciones periciales realizadas en siniestros
    ' de asistencia al fichero de texto para su envío
    '
    Public Function DatosPeritajes_IP() As Boolean

        On Error GoTo DatosPeritajes_IP_Err

        DatosPeritajes_IP = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim FechaPeritaje As Date ' Fecha de la que se ha de obtener la asignación de peritajes
        Dim Ajuste As Short
        Dim strsql As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4054"
        NumRegPeritajes = 0

        ' Establecemos el valor de la fecha de la que se ha de obtener la asignaión de Peritajes
        '
        FechaPeritaje = Today

        ' JLL - 23/11/2009
        ' Como el proceso se lanza diariamente no es necesario preguntar
        ' cuando es fin de semana
        '
        ' If Weekday(FechaPeritaje, vbSunday) = 2 Then
        '    Ajuste = 3
        ' Else
        '    Ajuste = 1
        ' End If
        ' FIN JLL - 23/11/2009

        Ajuste = 0
        strsql = Replace(SelectFich4, "'%Ajuste%'", CStr(Ajuste))

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de peritajes
        '

        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strsql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        Else
            HayPeritajes = True
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich4. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CanalFichero = AbreFichero(Fich4)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Peritajes"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF
            For i = 0 To objSiniestros.NumeroCampos
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                CadenaTRans = Space(objSiniestros.LongitudCampos.Item(Str(i)))
                ' Opciones de alineación y formateo
                '
                Select Case objSiniestros.Alineacion.Item(Str(i))
                    Case "IZ" ' Alineación Izquierda
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.Relleno(Str$(i)). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'

                            ''/*MUL INI */
                            'If Not IsDBNull(objSiniestros.Relleno.Item(Str(i))) and If objSiniestros.Relleno.Item(Str(i)) > 0 Then 
                            If Not IsDBNull(objSiniestros.Relleno.Item(Str(i))) Then
                                If objSiniestros.Relleno.Item(Str(i)) > 0 Then
                                    ''/*MUL FIN */
                                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.Relleno(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                    CadenaTRans = Replace(CadenaTRans, " ", Chr(objSiniestros.Relleno.Item(Str(i))))
                                End If
                            End If
                        End If
                    Case "DR" ' Alineación derecha
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = RSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            End If
                    Case "NU" ' Alineación campo númerico
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = Rslocal.Fields(i).Value
                                ' Comprobamos si hay decimales
                                '
                                If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                                    ' Si los hay eliminamos la coma
                                    CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                                Else
                                    ' Si no los hay añadimos dos ceros al final
                                    CadenaTRans = CadenaTRans & "00"
                                End If
                                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                ' Cambiamos lo espacios en blanco por ceros
                                CadenaTRans = Replace(CadenaTRans, " ", "0")
                            End If
                    Case "DT" ' Alineación y formato campo fecha
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "yyyyMMdd")
                            End If
                    Case "DS" ' Alineación y formato campo fecha española
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "yyyyMMdd")
                            End If
                    Case "ES" ' Campo alfanumerico pero rellenando con ceros
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)))
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(Str$(i)). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                CadenaTRans = Mid(Rslocal.Fields(i).Value, 1, objSiniestros.LongitudCampos.Item(Str(i)))
                                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                                CadenaTRans = Replace(CadenaTRans, " ", "0")
                            End If
                    Case Else ' Otros
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            If Not IsDBNull(Rslocal.Fields(i).Value) Then
                                CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            End If
                End Select

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i
            ' Grabamos el registro de exportación ya formateado en el fichero
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegPeritajes = NumRegPeritajes + 1
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Peritajes procesados: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosPeritajes_IP = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosPeritajes_IP_Err:
        If Err.Number = -2147217871 Then
            Resume
        End If
        If Rslocal.RecordCount = 0 Or lngResult = 0 Then
            FileClose(CanalFichero)
            DatosPeritajes_IP = True
            HayPeritajes = False
        Else
            FileClose(CanalFichero)
            'Kill Fich1
            'Kill Fich2
            'Kill Fich3
            DatosPeritajes_IP = False
        End If
    End Function

    ' Aplica el formato definido en la base de datos
    ' si hay un error retorna el mismo valor que como parametro
    Function AplicarFormatoExportacion(ByVal cadenatransformar As String, _
                                       ByVal strlongitud As Long, _
                                       ByVal strRelleno As String, _
                                       ByVal strFormato As String) As String
        On Error GoTo AplicarFormatoExportacion_ERR

        Dim CadenaTRans As String ' Cadena auxiliar
        CadenaTRans = Space(strlongitud)

        AplicarFormatoExportacion = cadenatransformar

        Select Case strFormato
            Case "IZ" ' Alineación Izquierda
                If Not IsDBNull(cadenatransformar) Then
                    AplicarFormatoExportacion = LSet(cadenatransformar, Len(CadenaTRans))
                    If Not IsDBNull(strRelleno) Then
                        If strRelleno > 0 Then
                            AplicarFormatoExportacion = Replace(AplicarFormatoExportacion, " ", Chr(strRelleno))
                        End If
                    End If
                End If
            Case "DR" ' Alineación derecha
                    If Not IsDBNull(cadenatransformar) Then
                        AplicarFormatoExportacion = RSet(cadenatransformar, Len(CadenaTRans))
                    End If
            Case "NU" ' Alineación campo númerico
                    If Not IsDBNull(cadenatransformar) Then
                        CadenaTRans = cadenatransformar
                        ' Comprobamos si hay decimales
                        If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                            ' Si los hay eliminamos la coma
                            CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                        Else
                            ' Si no los hay añadimos dos ceros al final
                            CadenaTRans = CadenaTRans & "00"
                        End If
                        CadenaTRans = New String("0", strlongitud - Len(CadenaTRans)) & CadenaTRans
                        ' Cambiamos los espacios en blanco por ceros
                        AplicarFormatoExportacion = Replace(CadenaTRans, " ", "0")
                    End If
            Case "DT" ' Alineación y formato campo fecha
                    If Not IsDBNull(cadenatransformar) Then
                        AplicarFormatoExportacion = Format(cadenatransformar, "yyyyMMdd")
                    End If
            Case "DS" ' Alineación y formato campo fecha española
                    If Not IsDBNull(cadenatransformar) Then
                        AplicarFormatoExportacion = Format(cadenatransformar, "yyyyMMdd")
                    End If
            Case "ES" ' Campo alfanumerico pero rellenando con ceros
                    CadenaTRans = New String("0", strlongitud)
                    If Not IsDBNull(cadenatransformar) Then
                        CadenaTRans = Mid(cadenatransformar, 1, strlongitud)
                        CadenaTRans = New String("0", strlongitud - Len(CadenaTRans)) & CadenaTRans
                        AplicarFormatoExportacion = Replace(CadenaTRans, " ", "0")
                    End If
            Case Else ' Otros
                    If Not IsDBNull(cadenatransformar) Then
                        AplicarFormatoExportacion = LSet(cadenatransformar, Len(CadenaTRans))
                    End If
        End Select
        Exit Function

AplicarFormatoExportacion_ERR:
        AplicarFormatoExportacion = cadenatransformar
    End Function


    ' Exporta los datos de las asignaciones periciales realizadas en siniestros
    ' de asistencia al fichero de texto para su envío
    '
    Public Function DatosPeritajesCia(ByVal cia As String) As Boolean

        On Error GoTo DatosPeritajesCia_Err

        DatosPeritajesCia = False

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim FechaPeritaje As Date ' Fecha de la que se ha de obtener la asignación de peritajes
        Dim Ajuste As Short
        Dim strsql As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        strErr = "4054"
        NumRegPeritajes = 0

        ' Establecemos el valor de la fecha de la que se ha de obtener la asignaión de Peritajes
        '
        FechaPeritaje = Today

        ' JLL - 23/11/2009
        ' Como el proceso se lanza diariamente no es necesario preguntar
        ' cuando es fin de semana
        '
        ' If Weekday(FechaPeritaje, vbSunday) = 2 Then
        '    Ajuste = 3
        ' Else
        '    Ajuste = 1
        ' End If
        ' FIN JLL - 23/11/2009

        Ajuste = 0
        strsql = Replace(SelectFich4, "'%Ajuste%'", CStr(Ajuste))

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de peritajes
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strsql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        Else
            HayPeritajes = True
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        CanalFichero = AbreFichero(Fich4)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Peritajes"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF
            For i = 0 To objSiniestros.NumeroCampos
                ' Opciones de alineación y formateo
                CadenaTRans = AplicarFormatoExportacion(IIf(IsDBNull(Rslocal.Fields(i).Value), "", Rslocal.Fields(i).Value), _
                                                        IIf(IsDBNull(objSiniestros.LongitudCampos.Item(Str(i))), "0", objSiniestros.LongitudCampos.Item(Str(i))), _
                                                        IIf(IsDBNull(objSiniestros.Relleno.Item(Str(i))), "0", objSiniestros.Relleno.Item(Str(i))), _
                                                        IIf(IsDBNull(objSiniestros.Alineacion.Item(Str(i))), "IZ", objSiniestros.Alineacion.Item(Str(i))))
                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i
            ' Grabamos el registro de exportación ya formateado en el fichero
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegPeritajes = NumRegPeritajes + 1
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Peritajes procesados: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosPeritajesCia = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosPeritajesCia_Err:
        If Err.Number = -2147217871 Then
            Resume
        End If
        If Rslocal.RecordCount = 0 Or lngResult = 0 Then
            FileClose(CanalFichero)
            DatosPeritajesCia = True
            HayPeritajes = False
        Else
            FileClose(CanalFichero)
            DatosPeritajesCia = False
        End If
    End Function


    ' Descripción:  Asigna el tipo de movimiento producido en cada una de las pólizas
    '               dependiendo de la secuencia de movimientos del histórico
    ' Observaciones El proceso opera directamente con la selección previa de
    '               pólizas registrada en la tabla mpAsisPolSel
    '
    Public Function EstadoPolizas() As Boolean

        On Error GoTo EstadoPolizas_Err

        ' Declaraciones
        '
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim lngResult As Integer ' Lineas afectadas después de la ejecución
        ' de command de un Recordset

        ' ----------------------------------------------------------------------------
        ' El hecho de que una misma póliza pueda tener diferentes estados en la misma
        ' fecha en un margen de segundos nos obliga a realizar el análisis por estado
        ' y no por póliza
        ' ----------------------------------------------------------------------------

        ' Primera casuistica: Altas
        '
        strSqlman = "Update    mpAsisPolSel Set mpAsisPolSel.Estado = 'A', Marca1 = 'A', mpAsisPolSel.Fecest = Polizahist.Fecest " & _
                    "From      Polizahist " & _
                    "Where     mpAsisPolSel.Codram = Polizahist.Codram and " & _
                    "          mpAsisPolSel.Numpol = Polizahist.Numpol and " & _
                    "          ((Polizahist.Servig <> 'S' and Polizahist.Servig <> 'A') OR Polizahist.Servig Is Null) and " & _
                    "          Polizahist.estado = 'A' and mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución mediante objeto Command
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)

        ' Segunda casuistica: Modificaciones sin Alta en la misma fecha
        '
        strSqlman = "Update    mpAsisPolSel Set mpAsisPolSel.Estado = 'M', mpAsisPolSel.Fecest = Polizahist.Fecest " & _
                    "From      Polizahist " & "Where     mpAsisPolSel.Codram = Polizahist.Codram and " & _
                    "          mpAsisPolSel.Numpol = Polizahist.Numpol and " & _
                    "          ((Polizahist.Servig <> 'S' and Polizahist.Servig <> 'A') OR Polizahist.Servig Is Null) and " & _
                    "          Polizahist.estado = 'M' and mpAsisPolSel.Marca1 <> 'A' and mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución mediante objeto Command
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)

        ' Tercera casuistica: Bajas
        '
        strSqlman = "Update    mpAsisPolSel Set mpAsisPolSel.Estado = 'B', Marca1 = 'B', mpAsisPolSel.Fecest = Polizaca.Fecbaj " & _
                    "From      Polizahist, Polizaca " & _
                    "Where     mpAsisPolSel.Codram = Polizahist.Codram and " & _
                    "          mpAsisPolSel.Numpol = Polizahist.Numpol and " & _
                    "          Polizahist.Codram = Polizaca.codram and " & _
                    "          Polizahist.Numpol = Polizaca.numpol and " & _
                    "          ((Polizahist.Servig <> 'S' and Polizahist.Servig <> 'A') OR Polizahist.Servig Is Null) and " & _
                    "          Polizahist.estado = 'B' and mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución mediante objeto Command
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)

        ' Eliminamos de la tabla de selección de pólizas todas aquellas bajas cuya
        ' fecha de baja sea mayor que la fecha de efecto especificada en proceso
        '
        strSqlman = "Delete    mpAsisPolSel Where mpAsisPolSel.Estado = 'B' and Marca1 = 'B' and " & _
                    "          mpAsisPolSel.FecEst > '" & FormatoFechaSQL(mvarFecEfe, False, False) & "' and " & _
                    "          mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución mediante objeto Command
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)

        ' Cuarta casuistica: Rehechos de pólizas sin baja en la misma fecha
        '
        strSqlman = "Update    mpAsisPolSel Set mpAsisPolSel.Estado = 'R', mpAsisPolSel.Fecest = Polizahist.Fecest " & _
                    "From      Polizahist " & _
                    "Where     mpAsisPolSel.Codram = Polizahist.Codram and " & _
                    "          mpAsisPolSel.Numpol = Polizahist.Numpol and " & _
                    "          ((Polizahist.Servig <> 'S' and Polizahist.Servig <> 'A') OR Polizahist.Servig Is Null) and " & _
                    "          Polizahist.estado = 'R' and mpAsisPolSel.Marca1 <> 'B' and mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución mediante objeto Command
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)


        ' Quinta casuistica: Rehechos de pólizas con baja en la misma fecha
        '
        strSqlman = "Update    mpAsisPolSel Set mpAsisPolSel.Estado = 'M', mpAsisPolSel.Fecest = Polizahist.Fecest " & _
                    "From      Polizahist " & _
                    "Where     mpAsisPolSel.Codram = Polizahist.Codram and " & _
                    "          mpAsisPolSel.Numpol = Polizahist.Numpol and " & _
                    "          ((Polizahist.Servig <> 'S' and Polizahist.Servig <> 'A') OR Polizahist.Servig Is Null) and " & _
                    "          Polizahist.estado = 'R' and mpAsisPolSel.Marca1 = 'B' and mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución mediante objeto Command
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)

        ' Abrimos un nuevo RecordSet para contar el total de pólizas
        '
        strSqlman = "Select Count(*) From mpAsisPolSel"

        claseBDExportacion.BDAuxRecord.Open(strSqlman, claseBDExportacion.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        lngPolizas = claseBDExportacion.BDAuxRecord.Fields(0).Value
        claseBDExportacion.BDAuxRecord.Close()
        EstadoPolizas = True

        Exit Function

EstadoPolizas_Err:
        ' Poner Errores
        EstadoPolizas = False
    End Function

    ' Procedure:    ActualizarPorcentaje
    ' Objetivo:     Actualiza barra de progreso y barra de estado del frmInstanciaPrincipal
    '               con porcentajes.
    ' Parametros:   Total = cantidad maxima que va a ver la barra de progreso.
    '
    Private Sub ActualizarPorcentaje(ByVal Total As Double)

        On Error Resume Next

        ' Declaraciones
        '
        Dim intPorcentaje As Short
        Total = Total / 100
        If Not Total = -1 Then ' Actualizar barra de estado, de progreso y porcentaje
            frmInstanciaPrincipal.prbProgreso.Value = frmInstanciaPrincipal.prbProgreso.Value + 1
            intPorcentaje = System.Math.Round(((frmInstanciaPrincipal.prbProgreso.Value / 100) * 100) / Total, 0)
            If CStr(intPorcentaje) & " %" <> frmInstanciaPrincipal.stbEstado.Panels(1).Text Then
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = CStr(intPorcentaje) & " %"
            End If
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub

    ' Procedure:    ContarRegistros
    ' Objetivo:     Cuenta el número de filas de un objeto RecordSet
    ' Parametros:   Objeto RecordSet del que se han de contar los registros
    ' Retorno:      Long con el número de registros contenidos

    Private Function ContarRegistros(ByVal objRs As ADODB.Recordset) As Integer

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

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


    ' Procedure:  Export Diario / Acumulado
    ' Objetivo:   Proceso de selección de polizas tipo Diario / Acumulado para
    '             exportacion de datos
    ' Parametros: Tipo de exportación: "Acumulado" o "Diario"
    '
    Public Function SeleccionPolizas() As Boolean

        On Error GoTo SeleccionPolizas_Err

        ' Declaraciones
        '
        Dim strSql As String ' Cadena que contiene la instrucción SQL para seleccionar
        ' pólizas, leida de la tabla de compañías de asistencia
        Dim strSqlman As String ' Cadena que contiene una instrucción SQL
        Dim intPorcentaje As Short
        Dim lngContador As Integer
        Dim dteFecha As Date ' Fecha de Efecto especificada para el proceso
        Dim dteFechaIni As Date ' Fecha de Ejecución de proceso
        Dim numreg() As Object
        Dim lngResult As Integer ' El command de ADO devuelve el número de registrso afectados
        Dim NumeroFicheros As Short
        Dim itemError As Short ' Indica la iteración de un mismo número de error

        ' Asignación de valores iniciales
        strErr = "Se ha producido un error en el proceso de selección de pólizas para exportación."
        itemError = 0
        dteFecha = frmInstanciaPrincipal.dtpFechaEfecto.Value
        dteFechaIni = Now

        ' Lee la Select de Criterio de selección de polizas para el proceso
        ' de exportación de Angel
        strSql = "SELECT * FROM MPASICIAS WHERE CODCIA = '" & Codcia & "' and CIRCUITOBD = '" & claseBDExportacion.BDManagement & "'"
        claseBDExportacion.BDSystemRecord.Open(strSql, claseBDExportacion.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        strSql = ""

        ' Si no se ha encontrado criterio para la compañía....
        If claseBDExportacion.BDSystemRecord.EOF Then
            strErr = "No existe criterio de selección de pólizas para esta compañía"
            claseBDExportacion.BDSystemRecord.Close()
            Err.Raise(1)
        End If

        ' Montar nombres de los ficheros de exportacion de Datos y Garantias.
        '
        NumeroFicheros = claseBDExportacion.BDSystemRecord.Fields("NumFicherosExp").Value
        '/*MUL INI
        'Fich1 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value & ".txt", "")
        'Fich2 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value & ".txt", "")
        'Fich3 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value & ".txt", "")
        'Fich4 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value & ".txt", "")
        'Fich5 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value & ".txt", "")
        Select Case Codcia
            Case "I"
                Fich1 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value & ".txt", "")
                Fich2 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value & ".txt", "")
                Fich3 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value & ".txt", "")
                Fich4 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value & ".txt", "")
                Fich5 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value & ".txt", "")
            Case "E"
                Fich1 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value), mvarArchivo & "\MDP_" & claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value & "_" & LoteEnviado & "_" & Format(Today(), "yyyyMMdd") & ".txt", "")
                Fich2 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value), mvarArchivo & "\MDP_" & claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value & "_" & LoteEnviado & "_" & Format(Today(), "yyyyMMdd") & ".txt", "")
                Fich3 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value), mvarArchivo & "\MDP_" & claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value & "_" & LoteEnviado & "_" & Format(Today(), "yyyyMMdd") & ".txt", "")
                Fich4 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value), mvarArchivo & "\MDP_" & claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value & "_" & LoteEnviado & "_" & Format(Today(), "yyyyMMdd") & ".txt", "")
                Fich5 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value), mvarArchivo & "\MDP_" & claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value & "_" & LoteEnviado & "_" & Format(Today(), "yyyyMMdd") & ".txt", "")
            Case "M"
                Fich1 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value & ".txt", "")
                Fich2 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value & ".txt", "")
                Fich3 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value & ".txt", "")
                Fich4 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value & ".txt", "")
                Fich5 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value & ".txt", "")
            Case Else
                Fich1 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp1").Value & ".txt", "")
                Fich2 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp2").Value & ".txt", "")
                Fich3 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp3").Value & ".txt", "")
                Fich4 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp4").Value & ".txt", "")
                Fich5 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value), mvarArchivo & "\" & Trim(Descia) & "_" & LoteEnviado & claseBDExportacion.BDSystemRecord.Fields("NombreExp5").Value & ".txt", "")
        End Select

        '/*MUL FIN

        ' Obtenemos la select que asiganada a cada fichero de exportación
        '
        SelectFich1 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("SelectFich1").Value), claseBDExportacion.BDSystemRecord.Fields("SelectFich1").Value, "")
        SelectFich2 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("SelectFich2").Value), claseBDExportacion.BDSystemRecord.Fields("SelectFich2").Value, "")
        SelectFich3 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("SelectFich3").Value), claseBDExportacion.BDSystemRecord.Fields("SelectFich3").Value, "")
        SelectFich4 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("SelectFich4").Value), claseBDExportacion.BDSystemRecord.Fields("SelectFich4").Value, "")
        SelectFich5 = IIf(Not IsDBNull(claseBDExportacion.BDSystemRecord.Fields("SelectFich5").Value), claseBDExportacion.BDSystemRecord.Fields("SelectFich5").Value, "")

        ' Si todo es correcto cogemos la Select especificada en la tabla para esta compañia
        ' asi como los datos para nombrar ficheros
        Select Case mvarTipoExp

            Case "Diario"
                strSql = claseBDExportacion.BDSystemRecord.Fields("CriterioDiario").Value ' Coje la Select de criterio de seleccion para Diario

            Case "Acumulado"
                strSql = claseBDExportacion.BDSystemRecord.Fields("CriterioAcum").Value ' Coje la Select de criterio de seleccion para Acumulado

        End Select
        claseBDExportacion.BDSystemRecord.Close()

        ' Visualiza mensaje en panatlla de inicio de proceso
        '
        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Preparando selección pólizas ..."

        ' Borramos los registros que puedan existir de procesos anteriores
        ' en la tabla de selección de pólizas
        '
        strSqlman = "Delete From mpAsisPolSel"


        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngResult)

        'JCLopez_i Meto el codigo de Suplidos porque de la conversion automatica no ha funcionado
        'claseBDExportacion.BDComand.ActiveConnection.ConnectionString = claseBDExportacion.BDWorkConnect.ConnectionString
        'claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        'claseBDExportacion.BDComand.CommandText = strSqlman
        'claseBDExportacion.BDComand.Execute(lngResult)
        'JCLopez_f

        ' Reemplaza los parametros de la SQL, indicado como %Fecha% y %Cia% , por los fecha
        ' de efecto y compañia de aistencia
        '
        strSql = Replace(strSql, "%FECHA%", "'" & frmInstanciaPrincipal.dtpFechaEfecto.Value.Month & "/" & frmInstanciaPrincipal.dtpFechaEfecto.Value.Day & "/" & frmInstanciaPrincipal.dtpFechaEfecto.Value.Year & " 23:59:59'")
        strSql = Replace(strSql, "%Cia%", "'" & Codcia & "'")
        FechaProceso = FecEfe

        ' Ejecutar sentencia SQL de seleccion de polizas para su exportación
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute(lngPolizas)

        ' Si no hay polizas a procesar...
        '
        'Esto es para pruebas lngPolizas = 0

        If lngPolizas = 0 Then
            strErr = "No existen pólizas a exportar con los criterios asignados."
            Err.Raise(1)
        End If

        ' Si no se ha podido obtener el número de pólizas a procesar emitimos un
        ' mensaje.  si todo  es correcto  inicializamos  os valores de la status
        ' bar y de la progres bar
        '
        If lngPolizas = -1 Then
            frmInstanciaPrincipal.stbEstado.Panels(1).Text = "NA"
        End If

        ' Si el modo de exportación es acumulado no es necesario actualizar el estado
        ' y la fecha de movimiento en la tabla de selección de polizas ya que en este
        ' el estado = 'A' y la fecha será la de efecto
        '
        If mvarTipoExp = "Diario" Then
            If EstadoPolizas() Then
                SeleccionPolizas = True
            Else
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                strErr = "Se ha producido un error en el proceso de selección de pólizas para exportación."
                Err.Raise(25002)
            End If
        Else
            SeleccionPolizas = True
            'If Not DepuraBajas Then
            '    strErr = "4032"
            '    Err.Raise 25002
            'End If
        End If

        ' Una vez obtenido el número definitivo de polizas a informar actualizamos
        ' la barra de estado y la progres bar.
        '
        If lngPolizas = 0 Then
            ' Quitar pantalla cuando error. Escribir fichero de texto con error
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            strErr = "No existen pólizas a exportar con los criterios asignados."
            Err.Raise(1)
        Else
            frmInstanciaPrincipal.prbProgreso.Visible = True
            frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
            frmInstanciaPrincipal.prbProgreso.Value = 1
            frmInstanciaPrincipal.stbEstado.Panels(1).Text = CStr(lngPolizas) & " Pólizas"
            frmInstanciaPrincipal.stbEstado.Panels(1).Text = "0 %"
        End If
        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Pólizas Seleccionadas: " & lngPolizas

        Exit Function

SeleccionPolizas_Err:
        frmInstanciaPrincipal.Cursor = Cursors.Default
        If TipoEjecucion = "P" Then
            itemError = 0
        End If
        If Err.Number = -2147217871 Then
            claseBDExportacion.BDWorkConnect.Errors.Clear()
            Resume
        Else
            SeleccionPolizas = False
        End If
    End Function

    ' Esta función elimina de la tabla de selección de pólizas todas aquellas
    ' que tengan fecha de baja superior a la fecha de proceso. Serán remitidas
    ' a la compañía de asistencia cuando la baja sea efectiva
    '
    Private Function DepuraBajas() As Boolean

        On Error GoTo DepuraBajas_Err

        ' Declaraciones
        '
        Dim strSql As String ' Instruccionde sql
        Dim sfecha As String

        sfecha = mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year
        'sfecha = VB6.Format(mvarFecEfe, "mm/dd/yyyy") & " 23:59:59"

        strSql = "Delete mpAsisPolsel " & "From   Polizaca " & "Where  mpAsisPolSel.Codram = Polizaca.Codram and " & "       mpAsisPolSel.Numpol = Polizaca.Numpol and " & "       (Polizaca.Polanu = 'S' and Polizaca.Fecbaj > '" & sfecha & "') and " & "       mpAsisPolSel.Cia = '" & Codcia & "'"

        ' Ejecución
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.Execute()

        DepuraBajas = True
        Exit Function

DepuraBajas_Err:
        DepuraBajas = False
    End Function

    ' Esta función ejecuta la actualización en el historico de pólizas de los
    ' movimiento exportados a la compañía
    '
    Public Function ActualizaHistorico() As Boolean

        On Error GoTo ActualizaHistorico_Err

        ' Declaraciones
        '
        Dim strSqlman As String
        Dim lngResult As Double
        Dim Vez As Short

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4035"
        Vez = 0

        ' Construimos la instrucción Sql según el tipo de exportación
        '
        Select Case mvarTipoExp
            Case "Acumulado"
                strSqlman = "Update Polizahist Set Servig = 'S', Fservig = '" & Now.Month & "/" & Now.Day & "/" & Now.Year & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second & "' " & "From   mpAsisPolSel, Polizaca " & "Where  Servig is null and Fservig is null and " & "       mpAsisPolSel.Codram = Polizaca.Codram and mpAsisPolSel.Numpol = Polizaca.Numpol and " & "       Polizahist.Fecest <= '" & mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & mvarFecEfe.Minute & ":" & mvarFecEfe.Second & "' and " & "       mpAsisPolSel.Codram = Polizahist.Codram and mpAsisPolSel.Numpol = Polizahist.Numpol and" & "       ((Polizahist.Estado <> 'B') or (Polizahist.Estado = 'B' and Polizaca.Fecbaj <= '" & mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & mvarFecEfe.Minute & ":" & mvarFecEfe.Second & "'))"

                ' Abrimos transaccción
                '
                claseBDExportacion.BDWorkConnect.BeginTrans()

                claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
                claseBDExportacion.BDComand.CommandText = strSqlman
                claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
                claseBDExportacion.BDComand.Execute(lngResult)

                claseBDExportacion.BDWorkConnect.CommitTrans()

                ' 22/03/2007 JLL

                ' Debemos marcar también en el histórico las bajas que tengan cuya fecha
                ' de baja sea igual o inferior a la fecha de efecto del proceso y que no
                ' hayan sido marcadas antes.

                strSqlman = "Update Polizahist Set Servig = 'S', Fservig = '" & mvarFecEfe.Month & "/" & _
                            mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & _
                            mvarFecEfe.Minute & ":" & mvarFecEfe.Second & "' " & _
                            "Where  Servig is null and Fservig is null and " & _
                            "       Polizahist.Fecest <= '" & mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & mvarFecEfe.Minute & ":" & mvarFecEfe.Second & _
                            "' and " & "       ((Polizahist.Estado <> 'B') or (Polizahist.Estado = 'B' and Polizahist.Fecbaj <= '" & mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & mvarFecEfe.Minute & ":" & mvarFecEfe.Second & "'))"


                ' Abrimos transaccción
                '
                claseBDExportacion.BDWorkConnect.BeginTrans()

                claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
                claseBDExportacion.BDComand.CommandText = strSqlman
                claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
                claseBDExportacion.BDComand.Execute(lngResult)

                claseBDExportacion.BDWorkConnect.CommitTrans()

                ' Fín 22/03/2007 JLL

                ' MUL MAN-3243)26/06/2015 ini
                ' Si hay una rehabilitación y la baja todavía no ha sido efectiva se ha de marcar como ya tratada la baja
                ' porque se ha rehabilitado
                strSqlman = "Update polizahist Set Servig = 'S', Fservig = '" & mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & _
                            mvarFecEfe.Minute & ":" & mvarFecEfe.Second & "' " & _
                            " from (select numhis,codram,numpol,estado from polizahist phr where phr.estado = 'R') as phr " & _
                            " where Servig is null and Fservig is null " & _
                            " and polizahist.estado = 'B' " & _
                            " and polizahist.codram = phr.codram " & _
                            " and polizahist.numpol = phr.numpol " & _
                            " and polizahist.numhis < phr.numhis " & _
                            " and polizahist.fecbaj >= '" & mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & mvarFecEfe.Hour & ":" & mvarFecEfe.Minute & ":" & mvarFecEfe.Second & "'"
                claseBDExportacion.BDWorkConnect.BeginTrans()

                claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
                claseBDExportacion.BDComand.CommandText = strSqlman
                claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
                claseBDExportacion.BDComand.Execute(lngResult)

                claseBDExportacion.BDWorkConnect.CommitTrans()
                ' MUL MAN-3243)26/06/2015 Fin

                ActualizaHistorico = True

            Case "Diario"
                strSqlman = "Update Polizahist Set Servig = 'S', Fservig = '" & Now.Month & "/" & Now.Day & "/" & _
                            Now.Year & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second & _
                            "' From mpAsisPolSel Where Servig is null and Fservig is null and Polizahist.Fecest <= '" & _
                            mvarFecEfe.Month & "/" & mvarFecEfe.Day & "/" & mvarFecEfe.Year & " " & _
                            mvarFecEfe.Hour & ":" & mvarFecEfe.Minute & ":" & mvarFecEfe.Second & _
                            "' and mpAsisPolSel.Codram = Polizahist.Codram and mpAsisPolSel.Numpol = Polizahist.Numpol"

                ' Abrimos transaccción
                '
                claseBDExportacion.BDWorkConnect.BeginTrans()

                claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
                claseBDExportacion.BDComand.CommandText = strSqlman
                claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
                claseBDExportacion.BDComand.Execute(lngResult)

                claseBDExportacion.BDWorkConnect.CommitTrans()

                ActualizaHistorico = True

            Case Else
                Err.Raise(1)
        End Select

        Exit Function

ActualizaHistorico_Err:
        If Err.Number = -2147217871 Then
            Vez = Vez + 1
            If Vez = 3 Then
                claseBDExportacion.BDWorkConnect.RollbackTrans()
                ActualizaHistorico = False
            Else
                Resume
            End If
        End If
    End Function


    Public Function ObtenerModoEjecucion() As String

        On Error GoTo ObtenerModoEjecucion_Err

        Dim Rslocal As ADODB.Recordset
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command

        Rslocal = New ADODB.Recordset

        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = "Select * From EjecucionProgramadaAsistencia"
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            ObtenerModoEjecucion = UCase("DIARIO")
        Else
            Rslocal.MoveFirst()
            ObtenerModoEjecucion = Rslocal.Fields("Modo").Value
        End If

        Rslocal.Close()
        Exit Function

ObtenerModoEjecucion_Err:
        ObtenerModoEjecucion = UCase("DIARIO")
    End Function


    ' 25-06-2007. Eloi función para guardar el log del resultado de la ejecución
    Public Function InsertarLog(ByRef cadenaLog As String, ByRef terminarProceso As Boolean) As Boolean

        Dim lCanal As Short
        Dim lNombreFichero As String

        InsertarLog = False
        On Error GoTo InsertarLog_Err

        ' Obtenemos el nombre del fichero log
        ' lNombreFichero = "c:\" & Format(Now, "yyyyMMdd") & "_Exportacion_Asistencia_" & UCase(Modo) & ".log"
        lNombreFichero = FicheroLog

        ' Abrimos el fichero log
        lCanal = AbreFicheroLog(lNombreFichero)

        ' Grabamos en el fichero especificado el vector de datos
        '
        cadenaLog = obtenerTextoError(cadenaLog)
        If GrabaFichero(cadenaLog, lCanal) Then
            ' Cerramos el fichero una vez que se ha guardado el log
            If CierraFichero(lCanal) Then
                InsertarLog = True
            End If
        End If
        ' En el caso de que se tenga que acabar el programa la función despúes de insertar un log
        If terminarProceso Then End
        Exit Function

InsertarLog_Err:
        FileClose(lCanal)
        'Kill(lNombreFichero)
        MsgBox("Error al grabar el fichero de log: " & lNombreFichero)
    End Function

    ' 2/3/2011 - JLL
    ' Obtiene la ruta y el nombre del fichero log
    '
    Public Function RutaLog() As String
        Dim i As Short
        Dim strParametroApp As String
        Dim lCanal As Short

        strParametroApp = Microsoft.VisualBasic.Command
        PosCom = CStr(InStr(strParametroApp, "*"))

        If CDbl(PosCom) > 0 Then
            RutaLog = Mid(strParametroApp, CDbl(PosCom) + 1, Len(strParametroApp) - CDbl(PosCom))
            lCanal = AbreFicheroLog(RutaLog)
            If lCanal = -1 Then
                RutaLog = "C:\Asistencia_Log_" & Now.Year & Now.Month & Now.Day
                PosCom = CStr(1)
            Else
                CierraFichero(lCanal)
            End If
        Else
            RutaLog = "C:\Asistencia_Log_" & Now.Year & Now.Month & Now.Day
            PosCom = CStr(1)
        End If
    End Function


    Private Function AbreFicheroLog(ByVal Fichero As String) As Short
        On Error GoTo AbreFicheroLog_Err

        ' Declaraciones
        Dim Canal As Short

        ' Obtenemos el numero de canal por el que abriremos el fichero
        Canal = FreeFile()

        ' Apertura del fichero
        FileOpen(Canal, Fichero, OpenMode.Append, OpenAccess.Write)
        AbreFicheroLog = Canal
        Exit Function

AbreFicheroLog_Err:
        AbreFicheroLog = -1
    End Function

    Private Function obtenerTextoError(ByRef lstrErr As String) As String
        Dim lstrSqlman As String ' Cadena que contiene la instrucción SQL
        Dim lngResult As Integer

        lstrSqlman = "select Mensaje from mdpplus.dbo.mdpSysError where coderr = '" & Replace(Trim(lstrErr), "'", "''") & "'"
        'lstrSqlman = "Select Count(*) From mpAsisPolSel"
        claseBDExportacion.BDAuxRecord.Open(lstrSqlman, claseBDExportacion.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not claseBDExportacion.BDAuxRecord.EOF Then
            lstrErr = claseBDExportacion.BDAuxRecord.Fields(0).Value
            '    mdpbd.BDAuxRecord.Close
        End If
        claseBDExportacion.BDAuxRecord.Close()
        obtenerTextoError = lstrErr
    End Function
    ' Fi Eloi

    ' Actualiza en la tabla MpAsicias el número de ordinal del lote
    '
    Public Function ActualizaLote(ByVal cia As String) As Boolean

        On Error GoTo ActualizaLote_Err

        ' Declaraciones
        Dim strSqlman As String
        Dim lngResult As Double
        Dim Vez As Short

        ' Valores iniciales
        strErr = "4035"
        Vez = 0

        strSqlman = "Update MpAsicias Set Lote = " & (LoteLeido + 1) & _
                    " Where  MpAsicias.Codcia = '" & cia & "'"

        ' Abrimos transaccción
        '
        ' mdpbd.BDWorkConnect.BeginTrans
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSqlman


        '/*MUL INI en VB6 tenenmos la BDSYSTEMCONNECT
        'claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDSystemConnect
        '/*MUL FIN

        claseBDExportacion.BDComand.Execute(lngResult)
        'mdpbd.BDWorkConnect.CommitTrans
        ActualizaLote = True

        Exit Function

ActualizaLote_Err:
        If Err.Number = -2147217871 Then
            Vez = Vez + 1
            If Vez = 3 Then
                claseBDExportacion.BDWorkConnect.RollbackTrans()
                ActualizaLote = False
            Else
                Resume
            End If
        End If

    End Function

    ' Obtención de los registros de polizas ya formateados para InterPartner
    '
    Public Function DatosCabecera_IP() As Boolean

        On Error GoTo Exportacion_Datos_IP_Err

        DatosCabecera_IP = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim strsql As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4016"
        NumRegPolizas = 0

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        ''/*MUL INI */
        strsql = Replace(SelectFich1, "%Cia%", "'" & Codcia & "'")
        ''/*MUL FIN */

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strsql 'SelectFich1
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich1. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CanalFichero = AbreFichero(Fich1)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Polizas"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF

            ' Leemos y acumulamos los valoes que forman el registro
            '
            For i = 0 To objSiniestros.NumeroCampos
                CadenaTRans = Rslocal.Fields(i).Value

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i

            ' Grabamos el registro de exportación ya formateado en el fichero
            '
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegPolizas = NumRegPolizas + 1
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Pólizas procesadas: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosCabecera_IP = True
        Else
            Err.Raise(1)
        End If
        Exit Function

Exportacion_Datos_IP_Err:
        If Err.Number = -2147217871 Then
            Resume
        ElseIf Err.Number = 94 Then
            Resume Next
        Else
            FileClose(CanalFichero)
            DatosCabecera_IP = False
        End If
    End Function

    ' Obtención de los registros de polizas ya formateados para 
    ' la compañia que se pasa como parametro: cia
    '
    Public Function DatosPolizasCia(ByVal cia As String) As Boolean

        On Error GoTo DatosPolizasCia_Err

        DatosPolizasCia = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim strSql As String
        Dim strGarantias As String
        Dim numGarantias As Double

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        strErr = "4016"
        NumRegPolizas = 0
        strGarantias = ""

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        strSql = Replace(SelectFich1, "%Cia%", "'" & Codcia & "'")

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich1. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CanalFichero = AbreFichero(Fich1)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Polizas"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF
            ' Leemos y acumulamos los valores que forman el registro
            For i = 0 To objSiniestros.NumeroCampos
                CadenaTRans = Rslocal.Fields(i).Value

                If cia = "M" And i = 21 Then 'calcular el número de garantias
                    numGarantias = DatosGarantiasMA(cia, Trim(Left(Rslocal.Fields(0).Value, 9)), _
                                                    Trim(Rslocal.Fields(4).Value), strGarantias)
                    CadenaTRans = Format(numGarantias, "000")
                    CadenaTRans = CadenaTRans + Left(strGarantias + Space(2900), 2900)

                    ' La llamada a DatosGarantiasMA configura el fichero 
                    ' para Garantias, se ha de volver a cargar el formato de Polizas
                    objSiniestros.ParaFichero = "Polizas"
                    objSiniestros.CodCiaAsist = Codcia
                End If

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i
            strVectorDatos = strVectorDatos + Space(2) ' chr(13) + chr(10)
            ' Grabamos el registro de exportación ya formateado en el fichero
            '
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegPolizas = NumRegPolizas + 1
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Pólizas procesadas: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosPolizasCia = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosPolizasCia_Err:
        If Err.Number = -2147217871 Then
            Resume
        ElseIf Err.Number = 94 Then
            Resume Next
        Else
            FileClose(CanalFichero)
            DatosPolizasCia = False
        End If
    End Function


    Public Function DatosRiesgo_IP() As Boolean

        On Error GoTo DatosRiesgos_IP_Err

        DatosRiesgo_IP = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim strsql As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4016"
        NumRegRiesgos = 0

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        ''/*MUL INI */
        strsql = Replace(SelectFich2, "%Cia%", "'" & Codcia & "'")
        ''/*MUL FIN */

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strsql 'SelectFich2
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich2. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CanalFichero = AbreFichero(Fich2)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Riesgos"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        Do While Not Rslocal.EOF

            ' Leemos y acumulamos los valoes que forman el registro
            '
            For i = 0 To objSiniestros.NumeroCampos
                CadenaTRans = Rslocal.Fields(i).Value

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i

            ' Grabamos el registro de exportación ya formateado en el fichero
            '
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegRiesgos = NumRegRiesgos + 1
                Rslocal.MoveNext()
                numreg = numreg + 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Riesgos Procesados: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosRiesgo_IP = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosRiesgos_IP_Err:
        If Err.Number = -2147217871 Then
            Resume
        Else
            FileClose(CanalFichero)
            'Kill Fich2
            DatosRiesgo_IP = False
        End If
    End Function

    Public Function DatosGarantias_IP() As Boolean

        On Error GoTo DatosGarantias_IP_Err

        DatosGarantias_IP = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim PolizaAct As String ' Poliza que se acaba de tratar
        Dim PolizaSig As String ' Poliza que se va a tratar
        Dim Vuelta As Short
        Dim X As Short
        Dim Filler As String ' Para completar hasta los 2000 caracteres
        Dim strsql As String

        'Dim contadorJCLopez As Double

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4016"
        NumRegGarantias = 0
        'contadorJCLopez = 0

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        ''/*MUL INI */
        strsql = Replace(SelectFich3, "%Cia%", "'" & Codcia & "'")
        ''/*MUL FIN */
        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strsql 'SelectFich3
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        CanalFichero = AbreFichero(Fich3)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "Garantias"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        PolizaAct = Rslocal.Fields(1).Value
        PolizaSig = PolizaAct
        Vuelta = 1
        Do While Not Rslocal.EOF

            ' Leemos y acumulamos los valoes que forman el registro
            '
            Do While PolizaAct = PolizaSig

                ' Establecemos a partir de que columna hay qye leer dependiendo
                ' de que estemos leyendo la misma póliza
                '
                If Vuelta = 1 Then
                    X = 0
                Else
                    X = 5
                End If

                ' Vamos leyendo los datos correspondiente a cada columna
                ' y generando un strinf acumulado con todos los datos
                '
                For i = X To objSiniestros.NumeroCampos
                    CadenaTRans = Rslocal.Fields(i).Value

                    strVectorDatos = strVectorDatos & CadenaTRans
                    CadenaTRans = ""
                Next i

                ' Si ha cambiado la póliza grabamos el string acumulado y
                ' volvemos a empezar con la garantias de la póliza siguiente
                ' si no seguimos leyendo las garantias de la póliza.
                '
                Rslocal.MoveNext()
                If Rslocal.EOF Then Exit Do
                numreg = numreg + 1
                PolizaSig = Rslocal.Fields(1).Value
                Vuelta = 2
            Loop

            ' Grabamos el registro de exportación ya formateado en el fichero
            '
            Filler = CStr(2000 - Len(strVectorDatos))

            '' !!! MUL codram = 510 INI
            '' MAXIMO 50 garantias 
            ''strVectorDatos = strVectorDatos & Space(CShort(Filler))
            If Filler < 0 Then Filler = 0
            strVectorDatos = Mid(strVectorDatos & Space(CShort(Filler)), 1, 2000)
            '' !!! MUL codram = 510 FIN


            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegGarantias = NumRegGarantias + 1
                If Not Rslocal.EOF Then PolizaAct = Rslocal.Fields(1).Value
                Vuelta = 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Garantias Procesadas: " & numreg - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
            'contadorJCLopez = contadorJCLopez + 1
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosGarantias_IP = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosGarantias_IP_Err:
        'MsgBox(contadorJCLopez, MsgBoxStyle.Critical)
        If Err.Number = -2147217871 Then
            Resume
        Else
            FileClose(CanalFichero)
            'Kill Fich3
            DatosGarantias_IP = False
        End If
    End Function


    Public Function DatosGarantiasCia(ByVal cia As String) As Boolean
        On Error GoTo DatosGarantiasCia_Err

        DatosGarantiasCia = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSql As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim PolizaAct As String ' Poliza que se acaba de tratar
        Dim PolizaSig As String ' Poliza que se va a tratar
        Dim Vuelta As Short
        Dim X As Short
        Dim Filler As String ' Para completar hasta los 2000 caracteres

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        strErr = "4016"
        NumRegGarantias = 0

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        strSql = Replace(SelectFich3, "%Cia%", "'" & Codcia & "'")

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Abrimos el fichero secuencial de exportación
        CanalFichero = AbreFichero(Fich3)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        objSiniestros.ParaFichero = "Garantias"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1
        'PolizaAct = Rslocal.Fields(1).Value
        'PolizaSig = PolizaAct
        'Vuelta = 1

        Do While Not Rslocal.EOF
            ' Leemos y acumulamos los valores que forman el registro
            'Do While PolizaAct = PolizaSig

            ' Establecemos a partir de que columna hay qye leer dependiendo
            ' de que estemos leyendo la misma póliza

            'If Vuelta = 1 Then
            X = 0
            ' Else
            '     X = 5
            ' End If

            ' Vamos leyendo los datos correspondiente a cada columna
            ' y generando un strinf acumulado con todos los datos
            For i = X To objSiniestros.NumeroCampos
                CadenaTRans = Rslocal.Fields(i).Value
                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i

            ' Si ha cambiado la póliza grabamos el strinf acumulado y
            ' volvemos a empezar con la garantias de la póliza siguiente
            ' si no seguimos leyendo las garantias de la póliza.
            'Rslocal.MoveNext()
            'If Rslocal.EOF Then Exit Do
            'numreg = numreg + 1
            ' PolizaSig = Rslocal.Fields(1).Value
            'Vuelta = 2
            'Loop

            ' Grabamos el registro de exportación ya formateado en el fichero
            '
            Filler = CStr(38 - Len(strVectorDatos))
            strVectorDatos = strVectorDatos & Space(CShort(Filler))
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                NumRegGarantias = NumRegGarantias + 1
                'If Not Rslocal.EOF Then PolizaAct = Rslocal.Fields(1).Value
                'Vuelta = 1
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Garantias Procesadas: " & numreg '- 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
            '/* MUL ini
            Rslocal.MoveNext()
            numreg = numreg + 1
            '/* MUL fin
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosGarantiasCia = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosGarantiasCia_Err:
        If Err.Number = -2147217871 Then
            Resume
        Else
            FileClose(CanalFichero)
            'Kill Fich3
            DatosGarantiasCia = False
        End If
    End Function


    Public Function DatosGarantiasMA(ByVal cia As String, ByVal numpol As String, _
                                     ByVal codram As String, ByRef strGarantias As String) As Double
        On Error GoTo DatosGarantiasMA_Err

        DatosGarantiasMA = 0

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        'Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSql As String ' Cadena que contiene la instrucción SQL
        'Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numreg As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        'Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        '  Dim PolizaAct As String ' Poliza que se acaba de tratar
        '  Dim PolizaSig As String ' Poliza que se va a tratar
        ' Dim Vuelta As Short
        'Dim X As Short
        'Dim Filler As String ' Para completar hasta los 2000 caracteres

        ' Creación de objetos
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        strErr = "4016"
        NumRegGarantias = 0
        strGarantias = ""

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        strSql = Replace(SelectFich2, "%Cia%", "'" & cia & "'")
        strSql = Replace(strSql, "%numpol%", "'" & numpol & "'")
        strSql = Replace(strSql, "%codram%", "'" & codram & "'")

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de cabecera
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strSql
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        End If

        ' Mostrar barra de progreso
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        objSiniestros.ParaFichero = "Garantias"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numreg = 1

        'recorremos todas las garantias de la poliza pasada como 
        'parametro y retornamos el string de todas las garantias
        Do While Not Rslocal.EOF

            ' Leemos y acumulamos los valoes que forman el registro
            '
            For i = 0 To objSiniestros.NumeroCampos
                strGarantias = strGarantias & Rslocal.Fields(i).Value
            Next i
            DatosGarantiasMA += 1

            frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Garantias Procesadas: " & numreg - 1
            Call ActualizarPorcentaje(lngPolizas)

            Rslocal.MoveNext()
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        Rslocal.Close()
        Exit Function

DatosGarantiasMA_Err:
        If Err.Number = -2147217871 Then
            Resume
        Else
            DatosGarantiasMA = -1
        End If
    End Function

    ' Esta función une en un solo fichero los ficheros de Polizas, Garantías y Riesgos
    ' para enviarloa InterPartner de acuerdo con su formato de ficheros y registros
    '
    Public Function FusionFicheros_IP() As Boolean

        On Error GoTo FusionFicheros_IP_Err

        FusionFicheros_IP = True

        ' Declaraciones
        '
        Dim CanalFichero As Short
        Dim Canalfichero2 As Short
        Dim CadenaLeida As String

        ' Abrimos el fichero para la union de todos los registros
        '
        FicheroFusion = mvarArchivo & "\" & Trim(Descia) & "_Fichero_" & LoteEnviado & ".txt"
        CanalFichero = AbreFichero(FicheroFusion)

        ' Primero grabamos el registro de cabecera
        '
        If Not GrabaRegistroCabecera_IP(CanalFichero) Then
            Err.Raise(1)
        End If

        ' Unimos fichero de Polizas
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich1. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Canalfichero2 = AbreFicheroLectura(Fich1)

        ' Bucle de lectura/escritura
        '
        Do While Not EOF(Canalfichero2)

            CadenaLeida = LineInput(Canalfichero2)
            If Not GrabaFichero(CadenaLeida, CanalFichero) Then
                Err.Raise(1)
            End If
        Loop
        FileClose(Canalfichero2)

        ' Unimos fichero de Riesgos
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fich2. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Canalfichero2 = AbreFicheroLectura(Fich2)

        ' Bucle de lectura/escritura
        '
        Do While Not EOF(Canalfichero2)

            CadenaLeida = LineInput(Canalfichero2)
            If Not GrabaFichero(CadenaLeida, CanalFichero) Then
                Err.Raise(1)
            End If
        Loop
        FileClose(Canalfichero2)

        ' Unimos fichero de Garantias
        '
        Canalfichero2 = AbreFicheroLectura(Fich3)

        ' Bucle de lectura/escritura
        '
        Do While Not EOF(Canalfichero2)

            CadenaLeida = LineInput(Canalfichero2)
            If Not GrabaFichero(CadenaLeida, CanalFichero) Then
                Err.Raise(1)
            End If
        Loop
        FileClose(Canalfichero2)

        ' Por ultio grabamos registro de totales
        '
        If Not GrabaRegistroTotales_IP(CanalFichero) Then
            Err.Raise(1)
        End If

        Exit Function

FusionFicheros_IP_Err:
        FusionFicheros_IP = False
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4056"
        FileClose(CanalFichero)
        FileClose(Canalfichero2)
    End Function


    ' función que concatena los ficheros de Polizas, Garantías y Riesgos
    ' para enviarlos a la compañia que se le pasa como parametro
    ' de acuerdo con su formato de ficheros y registros
    '
    Public Function FusionFicherosCia(ByVal cia As String) As Boolean
        On Error GoTo FusionFicherosCia_Err

        FusionFicherosCia = True

        ' Declaraciones
        Dim CanalFichero As Short
        Dim Canalfichero2 As Short
        Dim CadenaLeida As String

        ' Abrimos el fichero para la union de todos los registros
        Select Case cia
            Case "I"
                FicheroFusion = mvarArchivo & "\" & Trim(Descia) & "_Fichero_" & LoteEnviado & ".txt"
            Case "E"
                FicheroFusion = mvarArchivo & "\" & "MDP_polizas_" & LoteEnviado & "_" & Format(Today, "yyyyMMdd") & ".txt"
            Case "EG" 'Europ Asistance garantias
                FicheroFusion = mvarArchivo & "\" & "MDP_garantias_" & LoteEnviado & "_" & Format(Today, "yyyyMMdd") & ".txt"
            Case "M"
                FicheroFusion = mvarArchivo & "\" & Trim(Descia) & "_Fichero_" & LoteEnviado & ".txt"
        End Select

        CanalFichero = AbreFichero(FicheroFusion)

        ' Grabar registro de cabecera
        Select Case cia
            Case "I"
                If Not GrabaRegistroCabecera_IP(CanalFichero) Then
                    Err.Raise(1)
                End If
            Case "E"
                If Not GrabaRegistroCabeceraPolizaGarantiaEA(CanalFichero, "POL") Then
                    Err.Raise(1)
                End If
            Case "EG"
                If Not GrabaRegistroCabeceraPolizaGarantiaEA(CanalFichero, "GAR") Then
                    Err.Raise(1)
                End If
            Case "M"
                If Not GrabaRegistroCabeceraMA(CanalFichero) Then
                    Err.Raise(1)
                End If
        End Select

        ' Unimos fichero de Polizas
        Select Case cia
            Case "I", "E", "M"
                Canalfichero2 = AbreFicheroLectura(Fich1)
                Do While Not EOF(Canalfichero2)
                    CadenaLeida = LineInput(Canalfichero2)
                    If Not GrabaFichero(CadenaLeida, CanalFichero) Then
                        Err.Raise(1)
                    End If
                Loop
                FileClose(Canalfichero2)
            Case Else
                'nada
        End Select

        ' Unimos fichero de Riesgos
        Select Case cia
            Case "I"
                Canalfichero2 = AbreFicheroLectura(Fich2)
                Do While Not EOF(Canalfichero2)
                    CadenaLeida = LineInput(Canalfichero2)
                    If Not GrabaFichero(CadenaLeida, CanalFichero) Then
                        Err.Raise(1)
                    End If
                Loop
                FileClose(Canalfichero2)

            Case "E", "M"
                'No se utiliza
        End Select

        ' Unimos fichero de Garantias
        Select Case cia
            Case "I", "EG"
                Canalfichero2 = AbreFicheroLectura(Fich3)
                Do While Not EOF(Canalfichero2)
                    CadenaLeida = LineInput(Canalfichero2)
                    If Not GrabaFichero(CadenaLeida, CanalFichero) Then
                        Err.Raise(1)
                    End If
                Loop
                FileClose(Canalfichero2)

            Case "E", "M"
                'No se utiliza
        End Select

        ' Por ultimo grabamos registro de totales
        Select Case cia
            Case "I"
                If Not GrabaRegistroTotales_IP(CanalFichero) Then
                    Err.Raise(1)
                End If
            Case "E", "M"
                'No se utiliza
        End Select
        CierraFichero(CanalFichero)

        Exit Function

FusionFicherosCia_Err:
        FusionFicherosCia = False
        strErr = "4056"
        FileClose(CanalFichero)
        FileClose(Canalfichero2)
    End Function


    ' Esta función graba el registro de cabecera para el fichero fusionado de datos
    ' a enviar a InterPartner
    '
    Public Function GrabaRegistroCabecera_IP(ByRef CanalFichero As Short) As Boolean

        On Error GoTo GrabaRegistroCabecera_IP_Err

        GrabaRegistroCabecera_IP = True

        ' Declaraciones
        '
        Dim MontaCadena As String
        Dim Formateo As String
        Dim CadenaTemporal As Object
        Dim TipoEnvio As String

        ' Montamos el registro de cabecera a grabar en el fichero
        '
        MontaCadena = "HO0300"
        MontaCadena = MontaCadena & LoteEnviado

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CadenaTemporal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CadenaTemporal = Trim(Str(NumRegPolizas + NumRegRiesgos + NumRegGarantias + 2))
        Formateo = "00000000"
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CadenaTemporal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Formateo = Formateo + CadenaTemporal
        Formateo = Right(Formateo, 8)

        MontaCadena = MontaCadena & Formateo

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CadenaTemporal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CadenaTemporal = Format(Today, "yyyyMMdd")

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CadenaTemporal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        MontaCadena = MontaCadena + CadenaTemporal

        If UCase(mvarTipoExp) = UCase("acumulado") Then
            TipoEnvio = "R"
        ElseIf UCase(mvarTipoExp) = UCase("Diario") Then
            TipoEnvio = "P"
        Else
            TipoEnvio = " "
        End If

        MontaCadena = MontaCadena & TipoEnvio

        MontaCadena = MontaCadena & Space(1971)

        ' Grabamos el registro
        '
        If Not GrabaFichero(MontaCadena, CanalFichero) Then
            Err.Raise(1)
        End If

        Exit Function

GrabaRegistroCabecera_IP_Err:
        FileClose(CanalFichero)
        GrabaRegistroCabecera_IP = False
    End Function

    ' Función que Graba el registro de cabecera del fichero de datos
    ' a enviar a Europ Asistance
    '
    Public Function GrabaRegistroCabeceraPolizaGarantiaEA(ByRef CanalFichero As Short, ByVal tipo As String) As Boolean
        On Error GoTo GrabaRegistroCabeceraPolizaGarantiaEA_Err

        GrabaRegistroCabeceraPolizaGarantiaEA = True

        ' Declaraciones
        '
        Dim MontaCadena As String
        Dim Formateo As String
        Dim CadenaTemporal As Object
        Dim TipoEnvio As String

        ' Montamos el registro de cabecera a grabar en el fichero
        '
        MontaCadena = tipo ' POL = poliza / GAR = garantia 
        MontaCadena = MontaCadena & LoteEnviado

        Select Case tipo
            Case "POL"
                CadenaTemporal = Trim(Str(NumRegPolizas + 1))
            Case "GAR"
                CadenaTemporal = Trim(Str(NumRegGarantias + 1))
            Case Else
                CadenaTemporal = ""
        End Select

        Formateo = "00000000"
        Formateo = Formateo + CadenaTemporal
        Formateo = Right(Formateo, 8)

        MontaCadena = MontaCadena & Formateo
        CadenaTemporal = Format(Today, "yyyyMMdd")
        MontaCadena = MontaCadena + CadenaTemporal

        If UCase(mvarTipoExp) = UCase("acumulado") Then
            TipoEnvio = "R"
        ElseIf UCase(mvarTipoExp) = UCase("Diario") Then
            TipoEnvio = "P"
        Else
            TipoEnvio = " "
        End If

        MontaCadena = MontaCadena & TipoEnvio
        Select Case tipo
            Case "POL"
                MontaCadena = MontaCadena & Space(287)
            Case "GAR"
                MontaCadena = MontaCadena & Space(12)
            Case Else
                'nada
        End Select

        ' Grabamos el registro
        '
        If Not GrabaFichero(MontaCadena, CanalFichero) Then
            Err.Raise(1)
        End If
        Exit Function

GrabaRegistroCabeceraPolizaGarantiaEA_Err:
        FileClose(CanalFichero)
        GrabaRegistroCabeceraPolizaGarantiaEA = False
    End Function

    ' Función que Graba el registro de cabecera del fichero de datos
    ' a enviar a Multiasistencia
    '
    Public Function GrabaRegistroCabeceraMA(ByRef CanalFichero As Short) As Boolean
        On Error GoTo GrabaRegistroCabeceraMA_Err

        GrabaRegistroCabeceraMA = True

        ' Declaraciones
        '
        Dim MontaCadena As String
        Dim Formateo As String
        Dim CadenaTemporal As Object
        Dim TipoEnvio As String

        ' Montamos el registro de cabecera a grabar en el fichero
        ' Lote
        MontaCadena = "*"
        MontaCadena = MontaCadena & LoteEnviado

        'fecha generación
        CadenaTemporal = Format(Today, "yyyyMMdd")
        MontaCadena = MontaCadena + CadenaTemporal

        'número de registros
        CadenaTemporal = Trim(Str(NumRegPolizas + 1))
        Formateo = "0000000000"
        Formateo = Formateo + CadenaTemporal
        Formateo = Right(Formateo, 10)
        MontaCadena = MontaCadena & Formateo

        'Tipo de fichero de transmision + código compañia aseguradora
        MontaCadena = MontaCadena + "01" + "1212"

        'Tipo envio 0 -> actualización 1 -> toda la cartera
        If UCase(mvarTipoExp) = UCase("acumulado") Then
            TipoEnvio = "1"
        ElseIf UCase(mvarTipoExp) = UCase("Diario") Then
            TipoEnvio = "0"
        Else
            TipoEnvio = " "
        End If
        MontaCadena = MontaCadena & TipoEnvio
        MontaCadena = MontaCadena & Space(3453)

        'Final de linia CR + LF
        MontaCadena = MontaCadena + Space(2) 'Chr(13) + Chr(10)

        ' Grabamos el registro
        '
        If Not GrabaFichero(MontaCadena, CanalFichero) Then
            Err.Raise(1)
        End If
        Exit Function

GrabaRegistroCabeceraMA_Err:
        FileClose(CanalFichero)
        GrabaRegistroCabeceraMA = False
    End Function

    ' Esta función graba el registro de totales para el fichero fusionado de datos
    ' a enviar a InterPartner
    '
    Public Function GrabaRegistroTotales_IP(ByRef CanalFichero As Short) As Boolean

        On Error GoTo GrabaRegistroTotales_IP_Err

        GrabaRegistroTotales_IP = True

        ' Declaraciones
        '
        Dim MontaCadena As String
        Dim Formateo As String
        Dim CadenaTemporal As String
        Dim TipoEnvio As String

        ' Montamos el registro de cabecera a grabar en el fichero
        '
        MontaCadena = "HO0399"
        MontaCadena = MontaCadena & LoteEnviado

        CadenaTemporal = Trim(Str(NumRegPolizas + NumRegRiesgos + NumRegGarantias + 2))
        Formateo = "00000000"
        Formateo = Formateo + CadenaTemporal
        Formateo = Right(Formateo, 8)

        MontaCadena = MontaCadena & Formateo

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CadenaTemporal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CadenaTemporal = Today.Year & Today.Month & Today.Day

        MontaCadena = MontaCadena + CadenaTemporal

        If UCase(mvarTipoExp) = UCase("Acumulado") Then
            TipoEnvio = "R"
        ElseIf UCase(mvarTipoExp) = UCase("Diario") Then
            TipoEnvio = "P"
        Else
            TipoEnvio = " "
        End If
        MontaCadena = MontaCadena & TipoEnvio

        MontaCadena = MontaCadena & Space(1971)

        ' Grabamos el registro
        '
        If Not GrabaFichero(MontaCadena, CanalFichero) Then
            Err.Raise(1)
        End If

        Exit Function

GrabaRegistroTotales_IP_Err:
        FileClose(CanalFichero)
        GrabaRegistroTotales_IP = False
    End Function
    ' Esta función abre un fichero secuencial y devuelve el número de canal
    ' por el que ha sido abierto
    '
    Private Function AbreFicheroLectura(ByVal Fichero As String) As Short

        On Error GoTo AbreFicheroLectura_Err

        ' Declaraciones
        '
        Dim Canal As Short

        ' Obtenemos el numero de canal por el que abriremos el fichero
        '
        Canal = FreeFile()

        ' Apertura del fichero
        '
        FileOpen(Canal, Fichero, OpenMode.Input, OpenAccess.Read)
        AbreFicheroLectura = Canal

        Exit Function

AbreFicheroLectura_Err:
        AbreFicheroLectura = -1
    End Function

    ' Exporta los datos de las asignaciones periciales realizadas en siniestros
    ' de asistencia al fichero de texto para su envío
    '
    Public Function DatosCruceReferencias_IP() As Boolean

        On Error GoTo DatosCruceReferencias_IP_Err

        DatosCruceReferencias_IP = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numref As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim strsql As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4058"
        numref = 0


        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        ''/*MUL INI */
        strsql = Replace(SelectFich5, "%Cia%", "'" & Codcia & "'")
        ''/*MUL FIN */

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de peritajes
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = strsql 'SelectFich3
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        Else
            HayCruce = True
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        CanalFichero = AbreFichero(Fich5)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "CruceRef"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numref = 1
        Do While Not Rslocal.EOF
            For i = 0 To objSiniestros.NumeroCampos
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                CadenaTRans = Space(objSiniestros.LongitudCampos.Item(Str(i)))
                ' Opciones de alineación y formateo
                '
                Select Case objSiniestros.Alineacion.Item(Str(i))
                    Case "IZ" ' Alineación Izquierda
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.Relleno(Str$(i)). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                            ''/*MUL INI */
                            'If Not IsDBNull(objSiniestros.Relleno.Item(Str(i))) And objSiniestros.Relleno.Item(Str(i)) > 0 Then
                            If Not IsDBNull(objSiniestros.Relleno.Item(Str(i))) Then
                                If objSiniestros.Relleno.Item(Str(i)) > 0 Then
                                    ''/*MUL FIN */
                                    'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.Relleno(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                                    CadenaTRans = Replace(CadenaTRans, " ", Chr(objSiniestros.Relleno.Item(Str(i))))
                                End If
                            End If
                        End If
                    Case "DR" ' Alineación derecha
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = RSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                        End If
                    Case "NU" ' Alineación campo númerico
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Rslocal.Fields(i).Value
                            ' Comprobamos si hay decimales
                            '
                            If objUtilidades.ItemCaracter(CadenaTRans, ",") > 0 Then
                                ' Si los hay eliminamos la coma
                                CadenaTRans = "0" & Replace(CadenaTRans, ",", "")
                            Else
                                ' Si no los hay añadimos dos ceros al final
                                CadenaTRans = CadenaTRans & "00"
                            End If
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                            ' Cambiamos lo espacios en blanco por ceros
                            CadenaTRans = Replace(CadenaTRans, " ", "0")
                        End If
                    Case "DT" ' Alineación y formato campo fecha
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "yyyyMMdd")
                        End If
                    Case "DS" ' Alineación y formato campo fecha española
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = Format(Rslocal.Fields(i).Value, "yyyyMMdd")
                        End If
                    Case "ES" ' Campo alfanumerico pero rellenando con ceros
                        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                        CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)))
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(Str$(i)). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            CadenaTRans = Mid(Rslocal.Fields(i).Value, 1, objSiniestros.LongitudCampos.Item(Str(i)))
                            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objSiniestros.LongitudCampos(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            CadenaTRans = New String("0", objSiniestros.LongitudCampos.Item(Str(i)) - Len(CadenaTRans)) & CadenaTRans
                            CadenaTRans = Replace(CadenaTRans, " ", "0")
                        End If
                    Case Else ' Otros
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If Not IsDBNull(Rslocal.Fields(i).Value) Then
                            CadenaTRans = LSet(Rslocal.Fields(i).Value, Len(CadenaTRans))
                        End If
                End Select

                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i

            ' Grabamos el registro de exportación ya formateado en el fichero
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                numref = numref + 1
                Rslocal.MoveNext()
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Cruces procesados: " & numref - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosCruceReferencias_IP = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosCruceReferencias_IP_Err:
        If Err.Number = -2147217871 Then
            Resume
        End If
        If Rslocal.RecordCount = 0 Or lngResult = 0 Then
            FileClose(CanalFichero)
            DatosCruceReferencias_IP = True
            HayCruce = False
        Else
            FileClose(CanalFichero)
            DatosCruceReferencias_IP = False
        End If
    End Function


    ' Exporta los datos de las asignaciones periciales realizadas en siniestros
    ' de asistencia al fichero de texto para su envío
    '
    Public Function DatosCruceReferenciasCia(ByVal cia As String) As Boolean

        On Error GoTo DatosCruceReferenciasCia_Err

        DatosCruceReferenciasCia = True

        ' Declaraciones
        '
        Dim Rslocal As ADODB.Recordset ' Recordset local para el conjunto de datos de cabecera
        Dim strVectorDatos As String ' Contiene el valor formateado del campo a exportar
        Dim strSqlman As String ' Cadena que contiene la instrucción SQL
        Dim CadenaTRans As String ' Cadena auxiliar
        Dim lngResult As Double ' Número de registros afectados por la ejecución Command de Ado
        Dim numref As Integer ' Número de registro del recordset leido
        Dim i As Short ' Contador para bucles
        Dim CanalFichero As Short ' Canal de apertura del fichero de exportación
        Dim strSql As String

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Valores iniciales
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strErr. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strErr = "4058"
        numref = 0

        ' Reemplaza los parametros de la SQL, indicado como %Cia% , 
        ' por la compañia de aistencia
        'strSql = Replace(SelectFich5, "%Cia%", "'" & Codcia & "'")

        ' Lanzamos la select para construir el recordset con todos los datos
        ' de exportación de registro de fichero de peritajes
        '
        claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDExportacion.BDComand.CommandText = SelectFich5
        claseBDExportacion.BDComand.ActiveConnection = claseBDExportacion.BDWorkConnect
        Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

        If lngResult = 0 Then
            Err.Raise(1)
        Else
            HayCruce = True
        End If

        ' Abrimos el fichero secuencial de exportación
        '
        CanalFichero = AbreFichero(Fich5)
        If CanalFichero <= 0 Then Err.Raise(1)

        ' Mostrar barra de progreso
        '
        frmInstanciaPrincipal.prbProgreso.Visible = True
        frmInstanciaPrincipal.prbProgreso.Maximum = lngPolizas
        frmInstanciaPrincipal.prbProgreso.Value = 1

        ' Asignamos el código de la compañia para cargar los
        ' datos de formato que le correspondan
        '
        objSiniestros.ParaFichero = "CruceRef"
        objSiniestros.CodCiaAsist = Codcia

        Rslocal.MoveFirst()
        numref = 1
        Do While Not Rslocal.EOF
            For i = 0 To objSiniestros.NumeroCampos
                ' Opciones de alineación y formateo
                CadenaTRans = AplicarFormatoExportacion(IIf(IsDBNull(Rslocal.Fields(i).Value), "", Rslocal.Fields(i).Value), _
                                                        IIf(IsDBNull(objSiniestros.LongitudCampos.Item(Str(i))), "0", objSiniestros.LongitudCampos.Item(Str(i))), _
                                                        IIf(IsDBNull(objSiniestros.Relleno.Item(Str(i))), "0", objSiniestros.Relleno.Item(Str(i))), _
                                                        IIf(IsDBNull(objSiniestros.Alineacion.Item(Str(i))), "IZ", objSiniestros.Alineacion.Item(Str(i))))
                strVectorDatos = strVectorDatos & CadenaTRans
                CadenaTRans = ""
            Next i

            ' Grabamos el registro de exportación ya formateado en el fichero
            If GrabaFichero(strVectorDatos, CanalFichero) Then
                numref = numref + 1
                Rslocal.MoveNext()
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Cruces procesados: " & numref - 1
                Call ActualizarPorcentaje(lngPolizas)
                strVectorDatos = ""
            Else
                Err.Raise(1)
            End If
        Loop

        ' Cerramos el fichero secuencial de exportación y el Recordset local
        '
        Rslocal.Close()
        If CierraFichero(CanalFichero) Then
            DatosCruceReferenciasCia = True
        Else
            Err.Raise(1)
        End If
        Exit Function

DatosCruceReferenciasCia_Err:
        If Err.Number = -2147217871 Then
            Resume
        End If
        If Rslocal.RecordCount = 0 Or lngResult = 0 Then
            FileClose(CanalFichero)
            DatosCruceReferenciasCia = True
            HayCruce = False
        Else
            FileClose(CanalFichero)
            DatosCruceReferenciasCia = False
        End If
    End Function


    ' Esta función graba los datos de la tabla 'RegistroComunicacionesAsistencia'
    ' que es la tabla donde se guardan los datos enviados a la compañía de
    ' asistencia diariamente por ramo y modo de proceso (Acumulado o Diario)
    '
    Public Function RegistroComunicaciones() As Boolean

        On Error GoTo RegistroComunicaciones_Err

        claseBDExportacion.BDWorkRecord.Open("RegistroComunicacionAsistencia", claseBDExportacion.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        mvarFecEfe = CDate("28/10/2009")

        With claseBDExportacion.BDWorkRecord
            .AddNew()
            .Fields("TipoProceso").Value = Modo
            .Fields("FechaEnvio").Value = Now
            .Fields("FechaProceso").Value = FechaProceso
            .Fields("Hogar_Altas").Value = DatosRegistro("Hogar", "A")
            'If .Fields("Hogar_Altas").Value = "Error" Then Err.Raise(500)

            .Fields("Hogar_Bajas").Value = DatosRegistro("Hogar", "B")
            'If .Fields("Hogar_Bajas").Value = "Error" Then Err.Raise(500)

            .Fields("Hogar_Modificaciones").Value = DatosRegistro("Hogar", "M")
            'If .Fields("Hogar_Modificaciones").Value = "Error" Then Err.Raise(500)

            .Fields("Hogar_Rehabilitaciones").Value = DatosRegistro("Hogar", "R")
            'If .Fields("Hogar_Rehabilitaciones").Value = "Error" Then Err.Raise(500)

            .Fields("Hogar_Otros").Value = DatosRegistro("Hogar", "O")
            'If .Fields("Hogar_Otros").Value = "Error" Then Err.Raise(500)

            .Fields("Edificio_Altas").Value = DatosRegistro("Edificio", "A")
            'If .Fields("Edificio_Altas").Value = "Error" Then Err.Raise(500)

            .Fields("Edificio_Bajas").Value = DatosRegistro("Edificio", "B")
            'If .Fields("Edificio_Bajas").Value = "Error" Then Err.Raise(500)

            .Fields("Edificio_Modificaciones").Value = DatosRegistro("Edificio", "M")
            'If .Fields("Edificio_Modificaciones").Value = "Error" Then Err.Raise(500)

            .Fields("Edificio_Rehabilitaciones").Value = DatosRegistro("Edificio", "R")
            'If .Fields("Edificio_Rehabilitaciones").Value = "Error" Then Err.Raise(500)

            .Fields("Edificio_Otros").Value = DatosRegistro("Edificio", "O")
            'If .Fields("Edificio_Otros").Value = "Error" Then Err.Raise(500)

            .Fields("Usuario").Value = UsuaApli
            .Update()
            .Close()
        End With
        RegistroComunicaciones = True

        Exit Function

RegistroComunicaciones_Err:
        RegistroComunicaciones = False
    End Function

    ' Esta función devuelve el número de registros pertenecientes a la fecha de
    ' proceso especificada, al ramo especificado y al tipo de movimiento especificado
    '
    Public Function DatosRegistro(ByRef Ramo As String, ByRef Movimiento As String) As Object

        On Error GoTo DatosRegistro_Err

        Dim strSql As String
        Dim resultado As Integer

        If Ramo = "Hogar" Then
            Ramo = "('600','610','620','630','640','50')"
        End If
        If Ramo = "Edificio" Then
            ''MUL nuevo ramo 510 confort
            'Ramo = "('400','200','500','150')"
            Ramo = "('400','200','500','150','510')"
        End If

        If Movimiento = "O" Then
            strSql = "Select Count(*) From MpAsisPolSel Where Codram in " & Ramo & " and Estado not in ('A','B','M','R')"
        Else
            strSql = "Select Count(*) From MpAsisPolSel Where Codram in " & Ramo & " and Estado = '" & Movimiento & " '"
        End If

        claseBDExportacion.BDAuxRecord.Open(strSql, claseBDExportacion.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        claseBDExportacion.BDAuxRecord.MoveFirst()
        With claseBDExportacion.BDAuxRecord
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto DatosRegistro. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            DatosRegistro = .Fields(0).Value
            .Close()
        End With

        Exit Function

DatosRegistro_Err:
        DatosRegistro = "Error"
    End Function

End Class


