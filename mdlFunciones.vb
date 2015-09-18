Module mdlFunciones
    ' Public Sub main()
    Function Main(ByVal CmdArgs() As String) As Integer
        Dim msgUso As String

        'parse de parametros
        'asistprj -aplicacion -parametro -log 
        ' - aplicacion valores posibles:
        '       - aperturas
        '       - cierres
        '       - pagos
        '       - suplidos
        '       - importacion
        '       - exportacion
        ' - parametros depende de la aplicacion seleccionada
        ' - log path y nombre de fichero de log
        '
        msgUso = "Uso: asistprj -aplicacion -parametro -log" & vbCrLf & " - aplicacion valores posibles:  "

        If CmdArgs.Length = 0 Then
            MsgBox(msgUso)
        End If
        ' aplicación a ejecutar

        Select Case CmdArgs(0).Substring(1).ToUpper.Trim
            Case "APERTURAS"
                Dim frm As New frmPrincipalAperturas

                frm.ShowDialog()

            Case "CIERRRES"
                Dim frm As New frmPrincipalCierres

                frm.ShowDialog()

            Case "PAGOS"
                Dim frm As New frmPrincipalPagos

                frm.ShowDialog()

            Case "SUPLIDOS"
                Dim frm As New frmPrincipalSuplidos

                frm.ShowDialog()

            Case "IMPORT"
                Dim frm As New frmPrincipalImportacion

                frm.ShowDialog()

            Case "EXPORT"
                Dim frm As New frmPrincipalExportacion

                frm.ShowDialog()
            Case Else
                MsgBox(msgUso)
                Return -1
        End Select

        Return 0
    End Function

    Public Function FormatoFechaSQL(ByRef dtmFecha As Date, Optional ByRef blnComodin As Boolean = True, Optional ByRef blnHora As Boolean = True) As String

        ' Declaraciones
        '
        Dim strResultadoFecha As String
        Dim strFecha As String
        Dim strHora As String

        strResultadoFecha = ""
        strFecha = dtmFecha.Month & "/" & dtmFecha.Day & "/" & dtmFecha.Year
        strHora = Hour(dtmFecha) & ":" & Minute(dtmFecha) & ":" & Second(dtmFecha)
        If blnComodin Then strResultadoFecha = "#"
        strResultadoFecha = strResultadoFecha & strFecha
        If blnHora Then strResultadoFecha = strResultadoFecha & Space(1) & strHora
        If blnComodin Then strResultadoFecha = strResultadoFecha & "#"
        FormatoFechaSQL = strResultadoFecha

    End Function

    ' Esta función devuelve en variables globales ( mdlMain ) los datos de la compañia
    ' de asistencia especificada necesariso para crear el perjudicado
    '
    Public Function DatosCiaAsistencia(ByRef Codcia As String) As Boolean

        On Error GoTo DatosCiaAsistencia_Err

        ' Declaraciones
        '
        Dim strSql As String ' Instrucción Sql
        Dim directori As String 'directori inici
        Dim bd As New clsBD_NET

        '/*MUL T-19908 INI el FTP ha de estar en función de la compañia de asistencia
        'strSql = "SELECT Codcia, Nombre = Isnull(Nombre,''), Nif = Isnull(Nif,''), Direccion = Isnull(Direccion,''), " & _
        '         "       Poblacion = Isnull(Poblacion,''), CodPob = Isnull(CodPob,''), Numcia = isnull(Numcia, ''),  " & _
        '         "       IdRefer = isnull(Idrefer,'') " & _
        '         "From   MPASICIAS " & "Where  CodCia = '" & Codcia & "'"

        strSql = "SELECT Codcia, Nombre = Isnull(Nombre,''), Nif = Isnull(Nif,''), Direccion = Isnull(Direccion,''), " & _
                 "       Poblacion = Isnull(Poblacion,''), CodPob = Isnull(CodPob,''), Numcia = isnull(Numcia, ''),  " & _
                 "       Descripcion = Isnull(Descripcion,''), lote, IdRefer = isnull(Idrefer,''), FtpServer = Isnull(FtpServer,''), FtpUser = Isnull(FtpUser,''), " & _
                 "       FtpPwd = Isnull(FtpPwd,''), FtpPathin  = Isnull(FtpPathin,''), FtpPathout  = Isnull(FtpPathout,'') " & _
                 "From   MPASICIAS " & "Where  CodCia = '" & Codcia & "'"
        '/*MUL T-19908 FIN
        bd.BDSystemRecord.Open(strSql, bd.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If bd.BDSystemRecord.EOF Then
            Err.Raise(1)
        Else
            NIFCompa = bd.BDSystemRecord.Fields("NIF").Value
            DirecCompa = bd.BDSystemRecord.Fields("Direccion").Value
            PoblaCompa = bd.BDSystemRecord.Fields("Poblacion").Value
            CodPobCompa = bd.BDSystemRecord.Fields("codPob").Value
            NombreCompa = bd.BDSystemRecord.Fields("Nombre").Value
            Numcompa = bd.BDSystemRecord.Fields("Numcia").Value
            Descia = bd.BDSystemRecord.Fields("Descripcion").Value
            LoteLeido = bd.BDSystemRecord.Fields("Lote").Value ' Número de lote
            FtpCiaImport = bd.BDSystemRecord.Fields("ftppathin").Value
            FtpCiaExport = bd.BDSystemRecord.Fields("ftppathout").Value
            IdReferCompa = bd.BDSystemRecord.Fields("IdRefer").Value

            ''' aperturas
            strNIFCompa = NIFCompa
            strDirecCompa = DirecCompa
            strPoblaCompa = PoblaCompa
            strCodPobCompa = CodPobCompa
            strNombreCompa = NombreCompa
            strNumCompa = Numcompa
            strIdReferCompa = IdReferCompa
            '''


            '/*MUL T-19908 INI el FTP ha de estar en función de la compañia de asistencia
            'DiscoFTP = "195.77.230.7"
            DiscoFTP = bd.BDSystemRecord.Fields("FtpServer").Value
            'UsuarioFTP = "mprop@"
            UsuarioFTP = bd.BDSystemRecord.Fields("FtpUser").Value
            'PasswordFTP = "-pw ultpg.2"
            PasswordFTP = "-pw " & bd.BDSystemRecord.Fields("FtpPwd").Value
            'directori = "K:\Siniestros\Asistencia\Import"
            directori = bd.BDSystemRecord.Fields("FtpPathin").Value
            'PathImportacion = "K:\Siniestros\Asistencia\Import"
            PathImportacion = directori
            'ConfigFTP = "K:\Siniestros\Asistencia\FTP"
            ConfigFTP = directori & "\FTP"
            'DatosFTPApe = "K:\Siniestros\Asistencia\FTP\Aperturas"
            DatosFTPApe = ConfigFTP & "\Aperturas"
            'DatosFTPPag = "K:\Siniestros\Asistencia\FTP\Pagos"
            DatosFTPPag = ConfigFTP & "\Pagos"
            'MUL T-19908 FIN */
        End If
        bd.BDSystemRecord.Close()
        DatosCiaAsistencia = True
        Exit Function

DatosCiaAsistencia_Err:
        DatosCiaAsistencia = False
        If bd.BDSystemRecord.State = 1 Then bd.BDSystemRecord.Close()
        GlobalNumErr = "4004"
        '' en aperturas la linea de abajo estaba sin comentar
        'MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática.", MsgBoxStyle.Exclamation)
    End Function



    ' Este procedimiento genera el númeo de lote para los ficheros y el email de envio
    '
    Public Sub ObtenerLote()

        ' Declaraciones
        '
        Dim Formateo As String ' Cadena para el formateado del numero de lote
        Dim LoteIncremento As Double
        Dim LoteInc As String

        On Error GoTo ObtenerLote_Err

        If (LoteLeido + 1) <= 0 Then Err.Raise(1)

        LoteEnviado = ""

        LoteIncremento = LoteLeido + 1
        LoteInc = Trim(Str(LoteIncremento))
        Formateo = "000000"
        Formateo = Formateo & LoteInc
        Formateo = Right(Formateo, 6)
        LoteInc = Formateo

        LoteEnviado = LoteInc

        Exit Sub

ObtenerLote_Err:
        LoteEnviado = "Error"
        GlobalNumErr = "4055"
    End Sub

    Friend Function ObtenerCiaForm(ByVal cbxCompania As ComboBox, ByVal lstCompania As ListBox) As String
        ' Obtiene la compañia asistencia con la que se trabaja
        ObtenerCiaForm = lstCompania.Items.Item(cbxCompania.SelectedIndex)
    End Function

    'Friend Function ObtenerCiaForm(ByVal cbxCompania As ComboBox, ByVal lstCompania As ListBox) As String
    '    ' Obtiene la compañia asistencia con la que se trabaja
    '    ObtenerCiaForm = lstCompania.Items.Item(cbxCompania.SelectedIndex)
    'End Function


    ' 30/09/2009 - JLL
    '
    ' Esta función sustituye en el fichero de comandos de FTP el literal '1?'
    ' por el nombre del fichero a copiar
    '
    Public Function ComandosFTP(ByRef Path As Object, ByRef FicheroFTP1 As String, ByRef FicheroFTP As String, ByRef FicheroFTP3 As String) As Boolean

        On Error GoTo ComandosFTP_Err

        Dim CanalFTP As Short

        ' Borramos el último fichero de comandos ejecutado
        '
        Kill("C:\Comandos.scr")

        ' Obtenemos el numero de canal por el que abriremos el fichero
        '
        CanalFTP = FreeFile()

        ' Generamos un nuevo fichero de comandos con los nombres de ficheros a copiar
        '
        FileOpen(CanalFTP, "C:\Comandos.scr", OpenMode.Append, OpenAccess.Write)
        PrintLine(CanalFTP, "cd entrada/carteras")
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Path. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        PrintLine(CanalFTP, "put " + Path + "\" + FicheroFTP1)
        If HayPeritajes Then
            PrintLine(CanalFTP, "cd")
            PrintLine(CanalFTP, "cd entrada/peritos")
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Path. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            PrintLine(CanalFTP, "put " + Path + "\" + FicheroFTP)
        End If
        If HayCruce Then
            PrintLine(CanalFTP, "cd")
            PrintLine(CanalFTP, "cd entrada/referencias")
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Path. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            PrintLine(CanalFTP, "put " + Path + "\" + FicheroFTP3)
        End If
        FileClose(CanalFTP)

        ComandosFTP = True
        Exit Function

ComandosFTP_Err:
        If Err.Number = 53 Then Resume Next
        ComandosFTP = False
    End Function



    Public Function LlenarComboCias(ByVal objCombo As ComboBox, ByVal objlista As ListBox, ByVal pclsbd_net As clsBD_NET) As Boolean

        ' Declaraciones
        '
        Dim num As Integer
        Dim indiceLista, indiceCombo As Integer

        On Error GoTo LlenarComboCias_Err

        ' Inicialimos los objetos combo y list
        '
        objCombo.Items.Clear()
        objlista.Items.Clear()


        pclsbd_net.BDSystemRecord.Open("SELECT Codcia, Nombre, Inicio From MPASICIAS where activa = 'S' ", _
                                               pclsbd_net.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, _
                                               ADODB.LockTypeEnum.adLockOptimistic)
        num = 0
        pclsbd_net.BDSystemRecord.MoveFirst()
        Do While Not pclsbd_net.BDSystemRecord.EOF
            NombreCompa = Trim(claseBDExportacion.BDSystemRecord.Fields("Nombre").Value)
            strNombreCompa = Trim(claseBDPagos.BDSystemRecord.Fields("Nombre").Value)

            indiceLista = objlista.Items.Add(pclsbd_net.BDSystemRecord.Fields("Codcia").Value)             ' ListBox guarda el Codcia
            indiceCombo = objCombo.Items.Add(NombreCompa)             ' Descripcion compañia
            'objCombo.ItemData(objCombo.NewIndex) = objlista.NewIndex  ' Asocia listbox a combobox
            objCombo.Items.Item(indiceCombo) = NombreCompa
            If pclsbd_net.BDSystemRecord.Fields("Inicio").Value = "S" Then CiaDefault = num
            pclsbd_net.BDSystemRecord.MoveNext()
            num = num + 1
        Loop
        pclsbd_net.BDSystemRecord.Close()
        LlenarComboCias = True
        Exit Function

LlenarComboCias_Err:
        LlenarComboCias = False
        If pclsbd_net.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then pclsbd_net.BDWorkRecord.Close()
        'GlobalNumErr = "4004"
        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")
    End Function

    Public Sub LlenarComboProcesos(ByRef objCombo As System.Windows.Forms.ComboBox)

        objCombo.Items.Clear()

        claseBDImportar.BDSystemRecord.Open("SELECT * From mpAsiCiasProcImp Order By Orden", claseBDImportar.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        claseBDImportar.BDSystemRecord.MoveFirst()
        ProcesoIni = claseBDImportar.BDSystemRecord.Fields(4).Value
        Do While Not claseBDImportar.BDSystemRecord.EOF
            objCombo.Items.Add(claseBDImportar.BDSystemRecord.Fields("Descripcion").Value) ' Descripcion compañia
            claseBDImportar.BDSystemRecord.MoveNext()
        Loop
        claseBDImportar.BDSystemRecord.Close()

        ' La importación de suplidos de momento se añade manualmente para
        ' no interferir con Real, cuando se implante, se deberá añadir este
        ' registro a la tabla mpAsiciasProcImp como los demás
        objCombo.Items.Add("Datos Suplidos")

    End Sub

    Public Sub Log(ByRef Archivo As String, ByRef Mensaje As String, ByRef numlin As Integer, ByRef Fichero As String)

        Dim lFile As Integer
        Dim sLogFile As String
        Dim lExtension As Integer

        On Error Resume Next

        lFile = FreeFile()
        FileOpen(lFile, Archivo, OpenMode.Append, , OpenShare.Shared)
        PrintLine(lFile, "Fichero: " & Fichero & " Linea " & numlin & "   Fecha: " & Today & " " & TimeOfDay & ",")
        PrintLine(lFile, "Descripción Error: " & Mensaje & ",")
        PrintLine(lFile, New String("*", 5) & ",")
        PrintLine(lFile, "")
        FileClose(lFile)

    End Sub

    ''''''''''' aperturas
    Public Function CargarListView_aperturas(ByRef lvwControl As ListView, _
                                            ByRef strsql As String, _
                                            ByRef CampoKey As String, _
                                            ByVal ParamArray Campos() As Object) As Boolean

        ' Declaraciones
        '
        Dim objListItem As ListViewItem
        Dim intContador As Integer
        Dim numFila As Long
        Dim strCodSiniestro As String

        On Error GoTo CargarListView_aperturasError

        CargarListView_aperturas = True

        Cursor.Current = Cursors.WaitCursor

        numFila = 0

        lvwControl.Items.Clear()
        claseBDAperturas.BDWorkConnect.Errors.Clear()
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        claseBDAperturas.BDWorkRecord.Open(strsql, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDAperturas.BDWorkRecord.EOF Then
            Do
                For intContador = LBound(Campos) To UBound(Campos)
                    If intContador = LBound(Campos) Then
                        strCodSiniestro = claseBDAperturas.BDWorkRecord.Fields(Campos(intContador)).Value
                        lvwControl.Items.Add(strCodSiniestro)
                    Else
                        lvwControl.Items(numFila).SubItems.Add("" & claseBDAperturas.BDWorkRecord.Fields(Campos(intContador)).Value)
                    End If
                Next
                lvwControl.Items(numFila).Selected = False
                If Not claseBDAperturas.BDWorkRecord.EOF Then claseBDAperturas.BDWorkRecord.MoveNext()

                numFila = numFila + 1
            Loop Until claseBDAperturas.BDWorkRecord.EOF

        Else
            MsgBox("No existen datos para mostrar con los criterios de selección asignados.", MsgBoxStyle.Exclamation)
        End If

        claseBDAperturas.BDWorkRecord.Close()
        Cursor.Current = Cursors.Default

        Exit Function

CargarListView_aperturasError:
        Cursor.Current = Cursors.Default
        CargarListView_aperturas = False
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    'Public Sub ColorListItem(ByVal objlistitem As ListViewItem, ByVal Color As Color)

    '    Dim objListSubItem As ListViewItem.ListViewSubItem

    '    If IsNothing(objlistitem) Then Exit Sub

    '    objlistitem.ForeColor = Color

    '    For Each objListSubItem In objlistitem.SubItems
    '        objListSubItem.ForeColor = Color
    '    Next
    'End Sub
    ' Devuelve el código de aviso de una referencia
    '
    Public Function CodigoAviso(ByRef strRefer As String) As String

        Dim strsql As Object

        strsql = "Select    Top 1 Coderr " & "From      mpAsiHistError " & "Where     Referencia = '" & strRefer & "' And " & "          Proceso = 'A' and CodCia in ('R','I') " & "Order By  Fecgra"

        claseBDAperturas.BDWorkRecord.Open(strsql, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDAperturas.BDWorkRecord.EOF Then
            CodigoAviso = ""
        Else
            CodigoAviso = claseBDAperturas.BDWorkRecord.Fields("Coderr").Value
        End If

        claseBDAperturas.BDWorkRecord.Close()

    End Function

    '    Public Function LlenarComboProducto(ByRef objCombo As ComboBox, ByRef objLista As ListBox) As Boolean
    '        ' Declaraciones
    '        Dim strsql As String
    '        Dim indiceLista, indiceCombo As Short

    '        ' Instrucción Sql
    '        strsql = "select codram, descri from ramos where ramo1 in( '6', '4')"

    '        ' Reiniciamos el objeto combo
    '        objCombo.Items.Clear()

    '        ' Abrimos RecordSet con la instrucción Sql
    '        '
    '        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, claseBDPagos.BDWorkRecord.CursorType.adOpenDynamic, _
    '                                     ADODB.LockTypeEnum.adLockOptimistic)

    '        claseBDPagos.BDWorkRecord.MoveFirst()
    '        Do While Not claseBDPagos.BDWorkRecord.EOF
    '            strNomProd = Trim(claseBDPagos.BDWorkRecord.Fields("Descri").Value)
    '            indiceLista = objLista.Items.Add(claseBDPagos.BDWorkRecord.Fields("Codram").Value)
    '            indiceCombo = objCombo.Items.Add(strNomProd)

    '            objCombo.Items.Item(indiceCombo) = strNomProd

    '            claseBDPagos.BDWorkRecord.MoveNext()
    '        Loop

    '        objCombo.Items.Add("Todos los Productos")

    '        LlenarComboProducto = True
    '        claseBDPagos.BDWorkRecord.Close()
    '        Exit Function

    'LlenarComboProducto_Err:
    '        LlenarComboProducto = False
    '        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()
    '        'GlobalNumErr = "4004"
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")

    '    End Function


    Public Function LlenarComboProducto(ByRef objCombo As ComboBox, _
                                        ByRef objLista As ListBox, _
                                        ByRef prbdnet As clsBD_NET, _
                                        ByVal psOpcion As String) As Boolean
        On Error GoTo LlenarComboProducto_Err

        ' Declaraciones
        Dim strsql As String
        Dim indiceLista, indiceCombo As Short

        ' Instrucción Sql
        strsql = "select codram, descri from ramos where ramo1 in( '6', '4')"

        ' Reiniciamos el objeto combo
        objCombo.Items.Clear()

        ' Abrimos RecordSet con la instrucción Sql
        prbdnet.BDWorkRecord.Open(strsql, prbdnet.BDWorkConnect, _
                                  prbdnet.BDWorkRecord.CursorType.adOpenDynamic, _
                                  ADODB.LockTypeEnum.adLockOptimistic)

        prbdnet.BDWorkRecord.MoveFirst()
        Do While Not prbdnet.BDWorkRecord.EOF
            strNomProd = Trim(prbdnet.BDWorkRecord.Fields("Descri").Value)
            indiceLista = objLista.Items.Add(prbdnet.BDWorkRecord.Fields("Codram").Value)
            indiceCombo = objCombo.Items.Add(strNomProd)

            objCombo.Items.Item(indiceCombo) = strNomProd

            prbdnet.BDWorkRecord.MoveNext()
        Loop

        If psOpcion = "TODOS" Then
            objCombo.Items.Add("Todos los Productos")
        End If
        LlenarComboProducto = True
        prbdnet.BDWorkRecord.Close()
        Exit Function

LlenarComboProducto_Err:
        LlenarComboProducto = False
        If prbdnet.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then prbdnet.BDWorkRecord.Close()
        'GlobalNumErr = "4004"
        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")
    End Function


    '    Public Function LlenarComboProducto(ByRef objCombo As ComboBox, ByRef objLista As ListBox) As Boolean
    '        On Error GoTo LlenarComboProducto_Err

    '        ' Declaraciones
    '        Dim strsql As String
    '        Dim indiceLista, indiceCombo As Short

    '        ' Instrucción Sql
    '        strsql = "select codram, descri from ramos where ramo1 in( '6', '4')"

    '        ' Reiniciamos el objeto combo
    '        objCombo.Items.Clear()

    '        ' Abrimos RecordSet con la instrucción Sql
    '        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, claseBDCierres.BDWorkRecord.CursorType.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '        claseBDCierres.BDWorkRecord.MoveFirst()
    '        Do While Not claseBDCierres.BDWorkRecord.EOF
    '            strNomProd = Trim(claseBDCierres.BDWorkRecord.Fields("Descri").Value)
    '            indiceLista = objLista.Items.Add(claseBDCierres.BDWorkRecord.Fields("Codram").Value)
    '            indiceCombo = objCombo.Items.Add(strNomProd)

    '            objCombo.Items.Item(indiceCombo) = strNomProd

    '            claseBDCierres.BDWorkRecord.MoveNext()
    '        Loop

    '        objCombo.Items.Add("Todos los Productos")

    '        LlenarComboProducto = True
    '        claseBDCierres.BDWorkRecord.Close()
    '        Exit Function

    'LlenarComboProducto_Err:
    '        LlenarComboProducto = False
    '        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")

    '    End Function

    '    Public Function LlenarComboProducto(ByRef objCombo As ComboBox, ByRef objLista As ListBox) As Boolean
    '        On Error GoTo LlenarComboProducto_Err

    '        ' Declaraciones
    '        Dim strsql As String
    '        Dim indiceLista, indiceCombo As Short

    '        ' Instrucción Sql
    '        strsql = "select codram, descri from ramos where ramo1 in( '6', '4')"

    '        ' Reiniciamos el objeto combo
    '        objCombo.Items.Clear()

    '        ' Abrimos RecordSet con la instrucción SQL
    '        claseBDAperturas.BDWorkRecord.Open(strsql, claseBDAperturas.BDWorkConnect, claseBDAperturas.BDWorkRecord.CursorType.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '        claseBDAperturas.BDWorkRecord.MoveFirst()
    '        Do While Not claseBDAperturas.BDWorkRecord.EOF
    '            strNomProd = Trim(claseBDAperturas.BDWorkRecord.Fields("Descri").Value)
    '            indiceLista = objLista.Items.Add(claseBDAperturas.BDWorkRecord.Fields("Codram").Value)
    '            indiceCombo = objCombo.Items.Add(strNomProd)

    '            objCombo.Items.Item(indiceCombo) = strNomProd

    '            claseBDAperturas.BDWorkRecord.MoveNext()
    '        Loop

    '        objCombo.Items.Add("Todos los Productos")

    '        LlenarComboProducto = True
    '        claseBDAperturas.BDWorkRecord.Close()
    '        Exit Function

    'LlenarComboProducto_Err:
    '        LlenarComboProducto = False
    '        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")

    '    End Function


    Public Function ReferenciaExternaDuplicada(ByRef RefReparalia As String, ByRef Codsin As String) As Integer

        On Error GoTo ReferenciaExternaDuplicada_Err

        ' Declaraciones
        '
        Dim lngSinRefExtDup As Integer ' Número de siniestros con referencia externa duplicada
        Dim strsql As String ' instrucción sql

        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        lngSinRefExtDup = 0
        strsql = "SELECT Codsin " & "FROM Snsinies " & "WHERE Snsinies.refext = 'RP" & RefReparalia & "'"

        claseBDAperturas.BDWorkRecord.Open(strsql, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDAperturas.BDWorkRecord.EOF Then
            lngSinRefExtDup = 0
            Codsin = ""
        Else
            claseBDAperturas.BDWorkRecord.MoveFirst()
            lngSinRefExtDup = claseBDAperturas.BDWorkRecord.RecordCount
            Codsin = claseBDAperturas.BDWorkRecord.Fields(0).Value
        End If

        ReferenciaExternaDuplicada = lngSinRefExtDup
        Exit Function

ReferenciaExternaDuplicada_Err:
        ReferenciaExternaDuplicada = -1
        Codsin = ""

    End Function

    ' Esta función devuelve el número de siniestros de una misma póliza
    ' en el plazo de mas menos una semana
    '
    Public Function SiniestrosPoliza(ByRef Poliza As String, ByRef Ramo As String, ByRef FechaSiniestro As Date, ByRef Codsin As String) As Integer

        On Error GoTo SiniestrosPoliza_Err

        ' Declaraciones
        '
        Dim lngSiniestros As Integer ' Número de siniestros
        Dim strsql As String ' instrucción sql

        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        lngSiniestros = 0
        strsql = "SELECT Codsin, Dessin " & "FROM Snsinies " & "WHERE Snsinies.Numpol = '" & Poliza & "' AND " & "      Snsinies.Codram = '" & Ramo & "' And " & "      (Snsinies.Feccas >= '" & claseUtilidadesAperturas.FormatoFechaSQL(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -7, FechaSiniestro), False, False) & "' AND " & "      Snsinies.Feccas <= '" & claseUtilidadesAperturas.FormatoFechaSQL(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 7, FechaSiniestro), False, False) & "') " & "ORDER BY Snsinies.Fecrec"

        claseBDAperturas.BDWorkRecord.Open(strsql, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDAperturas.BDWorkRecord.EOF Then
            lngSiniestros = 0
            Codsin = ""
        Else
            claseBDAperturas.BDWorkRecord.MoveFirst()
            lngSiniestros = claseBDAperturas.BDWorkRecord.RecordCount
            Codsin = claseBDAperturas.BDWorkRecord.Fields(0).Value
        End If

        SiniestrosPoliza = lngSiniestros
        Exit Function

SiniestrosPoliza_Err:
        SiniestrosPoliza = -1
        Codsin = ""
    End Function

    Public Sub Imprimir_aperturas()

        On Error GoTo GeneraTMP_Err

        ' Declaraciones
        '
        Dim strSeleccion As String

        Cursor.Current = Cursors.WaitCursor

        '/*MUL INI
        ' Si no se ha realizado la busqueda no se ha generado la query no se puede imprimir
        If Trim(strSQLCR) = "" Then
            MsgBox("No existen datos para imprimir", MsgBoxStyle.Information)
            Exit Sub
        End If
        '/*MUL FIN

        'Generación de nuevos registros temporales en mdpImpAperAsis
        '
        With claseBDAperturas.BDWorkConnect
            boolTransaccion = True
            If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
            .BeginTrans()
            .Execute("DELETE FROM mdpImpAperAsis WHERE IDComponente = '" & strIDComp & "'")
            .Execute("INSERT INTO mdpImpAperAsis (IDComponente,REFER,PRODUCTO,CODSIN,POLIZA,AVISO,FAPER,FSINI,DESCR,CCAUSA,FECGRA,ESTADO,FECHAPROCESO,APERTURAPOR)" & strSQLCR)
            .CommitTrans()
            boolTransaccion = False
        End With

        ' Parametrización y lanzamiento del informe
        '
        With frmInstAperturas
            If .FiltroAviso.Checked Then
                strSeleccion = "Se han seleccionado los siniestros con avisos"
            End If

            If .FiltroErrores.Checked Then
                strSeleccion = "Se han seleccionado los siniestros con errores"
            End If

            If .FiltroNoPagados.Checked Then
                strSeleccion = "Se han seleccionado los siniestros pendientes de pago"
            End If
            If .FiltroPagados.Checked Then
                strSeleccion = "Se han seleccionado los siniestros pagados"
            End If
            If .FiltroTodos.Checked Then
                strSeleccion = "Se han seleccionado todos los siniestros"
            End If

            .CR2.set_ParameterFields(1, "FecDesde;" & CStr(.dtpDesde.Value) & ";TRUE")
            .CR2.set_ParameterFields(2, "FecHasta;" & CStr(.dtpHasta.Value) & ";TRUE")
            .CR2.set_ParameterFields(3, "Producto;" & CStr(.cbxProducto.Text) & ";TRUE")
            .CR2.set_ParameterFields(4, "Compañia;" & CStr(.cbxCompania.Text) & ";TRUE")
            .CR2.set_ParameterFields(5, "Seleccion;" & strSeleccion & ";TRUE")
            .CR2.Destination = Crystal.DestinationConstants.crptToWindow
            'Conectar al entorno de pruebas o producción

            'claseBDAperturas.ConfigurarImpesion()

            .CR2.Action = 1
        End With
        Cursor.Current = Cursors.Default
        Exit Sub

GeneraTMP_Err:
        frmInstanciaPrincipal.Cursor = Cursors.Default
        If claseBDAperturas.BDWorkConnect.Errors.Count > 0 Then
            If boolTransaccion Then claseBDAperturas.BDWorkConnect.RollbackTrans()
            claseBDAperturas.BDWorkConnect.Errors.Clear()
            MsgBox("El proceso de transacción de la Base de Datos ha devuelto un error. La transacción será abortada", MsgBoxStyle.Information)
        Else
            MsgBox("Se ha producido un error al intentar imprimir la selección de referencias de asistencia. Llame a informática.", MsgBoxStyle.Critical)
        End If
    End Sub

    ''''''''''pagos
    'Public Sub ColorListItem(ByVal objlistitem As ListViewItem, ByVal Color As Color)

    '    Dim objListSubItem As ListViewItem.ListViewSubItem

    '    If IsNothing(objlistitem) Then Exit Sub

    '    objlistitem.ForeColor = Color

    '    For Each objListSubItem In objlistitem.SubItems
    '        objListSubItem.ForeColor = Color
    '    Next
    'End Sub

    Public Function CargarListView_pagos(ByRef lvwControl As ListView, _
                                         ByVal strSelect As String, _
                                         ByVal strFromwhere As String, _
                                         ByVal strOrderBy As String, _
                                         ByVal CampoKey As String, _
                                         ByVal ParamArray Campos() As Object) As Boolean
        ' Declaraciones
        Dim intContador As Integer
        Dim numFila As Long
        Dim Importe, Iva, ImporteTotal As Double
        Dim CadenaSQL, strCodSiniestro As String
        Dim strsql As String
        Dim lcampos_ini As Integer
        Dim lcampos_fin As Integer
        Dim lstviewietm As ListViewItem

        On Error GoTo CargarListView_pagosError

        ' Ponemos a 0 el recuadro de resumen
        Importe = 0
        Iva = 0
        ImporteTotal = 0
        intContador = 0
        lcampos_ini = LBound(Campos)
        lcampos_fin = UBound(Campos)
        strsql = strSelect & strFromwhere & strOrderBy

        frmInstPagos.lbImporte.Text = Format(Importe, "##,##0.00")
        frmInstPagos.lbIVA.Text = Format(Iva, "##,##0.00")
        frmInstPagos.lbTotal.Text = Format(ImporteTotal, "##,##0.00")
        frmInstPagos.lbResumenSiniestros.Text = ""
        frmInstPagos.lbResumenReferencias.Text = ""
        frmInstPagos.stbEstado.Panels(2).Text = "0"

        CargarListView_pagos = True

        frmInstanciaPrincipal.Cursor = Cursors.WaitCursor

        lvwControl.Items.Clear()
        claseBDPagos.BDWorkConnect.Errors.Clear()
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDPagos.BDWorkRecord.EOF Then
            numFila = 0
            Do
                Importe = Importe + claseBDPagos.BDWorkRecord.Fields("t2_impor").Value
                Iva = Iva + claseBDPagos.BDWorkRecord.Fields("t2_imptva").Value
                'With lvwControl.Items.Add
                'objListItem = lvwControl.Items.Add("" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value)
                'With lvwControl.Items.Add("" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value)

                'objListItem = lvwControl.Items.Add(
                For intContador = lcampos_ini To lcampos_fin
                    'For intContador = LBound(Campos) To UBound(Campos)
                    ' "" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value
                    'If intContador = LBound(Campos) Then
                    If intContador = lcampos_ini Then
                        lvwControl.Items.Add("" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value)
                        'objListItem.Text = "" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value
                    Else
                        If claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Type = ADODB.DataTypeEnum.adNumeric And Campos(intContador) <> "T2_CODRAM" And Campos(intContador) <> "T2_NUMORD" Then
                            lvwControl.Items(numFila).SubItems.Add("" & Format(claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value, "##,##0.00"))
                        Else
                            lvwControl.Items(numFila).SubItems.Add("" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value)
                        End If
                    End If
                Next
                'If Not Trim(CampoKey) = "" Then
                '    objListItem.Key = claseBDPagos.BDWorkRecord.Fields(CampoKey).Value
                'End If
                lvwControl.Items(numFila).Selected = False
                'End With
                claseBDPagos.BDWorkRecord.MoveNext()
                numFila = numFila + 1
            Loop Until claseBDPagos.BDWorkRecord.EOF
        Else
            MsgBox("No existen datos para mostrar con los criterios de selección asignados.", MsgBoxStyle.Exclamation)
            'GlobalNumErr = "4005"
            'Err.Raise(Val(GlobalNumErr))
        End If
        claseBDPagos.BDWorkRecord.Close()

        ' Suma y formateo de totales para el recuadro de resumen
        ImporteTotal = Importe + Iva

        frmInstPagos.lbImporte.Text = Format(Importe, "##,##0.00")
        frmInstPagos.lbIVA.Text = Format(Iva, "##,##0.00")
        frmInstPagos.lbTotal.Text = Format(ImporteTotal, "##,##0.00")

        ' Contamos los siniestros diferentes del RecordSet
        strsql = "Select count(distinct T2_CODSIN) " & strFromwhere
        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDPagos.BDWorkRecord.EOF Then
            frmInstPagos.lbResumenSiniestros.Text = claseBDPagos.BDWorkRecord.Fields(0).Value
        End If
        claseBDPagos.BDWorkRecord.Close()
        '/* MUL ineficiente lo sustituyo por el select count
        'frmInstanciaPrincipal.lbResumenSiniestros.Text = claseUtilidadesPagos.ContarRegistrosDistinct("T2_CODSIN", claseBDPagos.BDWorkRecord, Val(frmInstanciaPrincipal.stbEstado.Panels(2).Text))

        ' Contamos las referencias diferentes del RecordSet
        strsql = "Select count(distinct T2_REFER) " & strFromwhere
        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDPagos.BDWorkRecord.EOF Then
            frmInstPagos.lbResumenReferencias.Text = claseBDPagos.BDWorkRecord.Fields(0).Value
        End If
        claseBDPagos.BDWorkRecord.Close()
        '/* MUL ineficiente lo sustituyo por el select count
        'frmInstanciaPrincipal.lbResumenReferencias.Text = claseUtilidadesPagos.ContarRegistrosDistinct("T2_REFER", claseBDPagos.BDWorkRecord, Val(frmInstanciaPrincipal.stbEstado.Panels(2).Text))

        ' No dejar ningún elemento seleccionado
        'lvwControl.SelectedItem = Nothing

        frmInstanciaPrincipal.Cursor = Cursors.Default

        'JCLopez_f

        Exit Function

CargarListView_pagosError:
        'If Err.Number = -2147217871 Then
        '    Resume
        'End If
        frmInstanciaPrincipal.Cursor = Cursors.Default
        CargarListView_pagos = False
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()
        'JCLopez_i  
        'objError.Tipo = Pantalla
        'objError.Ver(IdProceso, "4005", , Codcia)
        'MsgBox("No existen datos para mostrar con los criterios de selección asignados.", MsgBoxStyle.Exclamation)
        'JCLopez_f
        frmInstPagos.lbImporte.Text = Format(0, "##,##0.00")
        frmInstPagos.lbIVA.Text = Format(0, "##,##0.00")
        frmInstPagos.lbTotal.Text = Format(0, "##,##0.00")

    End Function

    '    Public Function DatosCiaAsistencia(ByVal Codcia As String) As Boolean

    '        On Error GoTo DatosCiaAsistencia_Err

    '        ' Declaraciones
    '        '
    '        Dim strsql As String                ' Instrucción Sql

    '        strsql = "SELECT Codcia, Nombre = Isnull(Nombre,''), Nif = Isnull(Nif,''), Direccion = Isnull(Direccion,''), " & _
    '                 "       Poblacion = Isnull(Poblacion,''), CodPob = Isnull(CodPob,''), Numcia = isnull(Numcia, ''), IdRefer = Isnull(IdRefer,'') " & _
    '                 "From  MPASICIAS " & _
    '                 "Where CodCia = '" & Codcia & "'"
    '        claseBDPagos.BDSystemRecord.Open(strsql, claseBDPagos.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '        If claseBDPagos.BDSystemRecord.EOF Then
    '            'Err.Raise(1)
    '            DatosCiaAsistencia = False
    '        Else
    '            strNIFCompa = claseBDPagos.BDSystemRecord.Fields("NIF").Value
    '            strDirecCompa = claseBDPagos.BDSystemRecord.Fields("Direccion").Value
    '            strPoblaCompa = claseBDPagos.BDSystemRecord.Fields("Poblacion").Value
    '            strCodPobCompa = claseBDPagos.BDSystemRecord.Fields("codPob").Value
    '            strNombreCompa = claseBDPagos.BDSystemRecord.Fields("Nombre").Value
    '            strNumCompa = claseBDPagos.BDSystemRecord.Fields("Numcia").Value
    '            strIdReferCompa = claseBDPagos.BDSystemRecord.Fields("IdRefer").Value
    '            DatosCiaAsistencia = True
    '        End If
    '        claseBDPagos.BDSystemRecord.Close()
    '        Exit Function

    'DatosCiaAsistencia_Err:
    '        DatosCiaAsistencia = False
    '        If claseBDPagos.BDSystemRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDSystemRecord.Close()
    '        'JCLopez_i
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática.", MsgBoxStyle.Exclamation)
    '        'GlobalNumErr = "4004"
    '        'JCLopez_f
    '    End Function


    '    Public Function LlenarComboCias(ByVal objCombo As ComboBox, ByVal objlista As ListBox) As Boolean

    '        ' Declaraciones
    '        '
    '        Dim num As Integer
    '        Dim indiceLista, indiceCombo As Integer

    '        On Error GoTo LlenarComboCias_Err

    '        ' Inicialimos los objetos combo y list
    '        '
    '        objCombo.Items.Clear()
    '        objlista.Items.Clear()


    '        claseBDPagos.BDSystemRecord.Open("SELECT Codcia, Nombre, Inicio From MPASICIAS  where activa = 'S' ", claseBDPagos.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '        num = 0
    '        claseBDPagos.BDSystemRecord.MoveFirst()
    '        Do While Not claseBDPagos.BDSystemRecord.EOF
    '            strNombreCompa = Trim(claseBDPagos.BDSystemRecord.Fields("Nombre").Value)
    '            indiceLista = objlista.Items.Add(claseBDPagos.BDSystemRecord.Fields("Codcia").Value)             ' ListBox guarda el Codcia
    '            indiceCombo = objCombo.Items.Add(strNombreCompa)             ' Descripcion compañia
    '            'objCombo.ItemData(objCombo.NewIndex) = objlista.NewIndex  ' Asocia listbox a combobox
    '            objCombo.Items.Item(indiceCombo) = strNombreCompa
    '            If claseBDPagos.BDSystemRecord.Fields("Inicio").Value = "S" Then strCiaDefault = num
    '            claseBDPagos.BDSystemRecord.MoveNext()
    '            num = num + 1
    '        Loop

    '        claseBDPagos.BDSystemRecord.Close()
    '        LlenarComboCias = True
    '        Exit Function

    'LlenarComboCias_Err:
    '        LlenarComboCias = False
    '        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()
    '        'GlobalNumErr = "4004"
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")
    '    End Function



    Public Sub Imprimir_pagos()

        On Error GoTo Imprimir_Error

        ' Declaraciones
        '
        Dim strSeleccion As String
        Dim intentos As Short

        '/*MUL INI
        ' Si no se ha realizado la busqueda no se ha generado la query no se puede imprimir
        If Trim(strSQLCR) = "" Then
            MsgBox("No existen datos para imprimir", MsgBoxStyle.Information)
            Exit Sub
        End If
        '/*MUL FIN


        frmInstanciaPrincipal.Cursor = Cursors.WaitCursor
        intentos = 0

        'Generación de nuevos registros temporales en mdpImpAperAsis
        '
        With claseBDPagos.BDWorkConnect
            claseBDPagos.BDWorkConnect.CommandTimeout = 0
            boolTransaccion = True
            If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
            .BeginTrans()
            .Execute("DELETE FROM mdpImpPagosAsis ")
            .Execute("INSERT INTO mdpImpPagosAsis (Codcia,Refer,Numord,Fpago,Importe,Ultpag,Numpol,CauPer,Iva,Fimportacion,Codsin,Pagado,Indgas,Estado,Festado,Situacion,Perito,Codram,Total,Fproceso, FechaSiniestro, FechaExportacion, Factura, Codmod, Codgru)" & strSQLCR$)
            .CommitTrans()
            boolTransaccion = False
        End With

        ' Parametrización y lanzamiento del informe
        '
        With frmInstPagos

            If .FiltroAviso.Checked Then
                strSeleccion = "Se han seleccionado los siniestros con avisos"
            End If

            If .FiltroErrores.Checked Then
                strSeleccion = "Se han seleccionado los siniestros con errores"
            End If

            If .FiltroNoPagados.Checked Then
                strSeleccion = "Se han seleccionado los siniestros pendientes de pago"
            End If
            If .FiltroPagados.Checked Then
                strSeleccion = "Se han seleccionado los siniestros pagados"
            End If
            If .FiltroTodos.Checked Then
                strSeleccion = "Se han seleccionado todos los siniestros"
            End If

            .CR2.set_ParameterFields(1, "FecDesde;" & CStr(.dtpDesde.Value) & ";TRUE")
            .CR2.set_ParameterFields(2, "FecHasta;" & CStr(.dtpHasta.Value) & ";TRUE")
            .CR2.set_ParameterFields(3, "Producto;" & CStr(.cbxProducto.Text) & ";TRUE")
            .CR2.set_ParameterFields(4, "Compañia;" & CStr(.cbxCompania.Text) & ";TRUE")
            .CR2.set_ParameterFields(5, "Seleccion;" & strSeleccion & ";TRUE")
            .CR2.Destination = Crystal.DestinationConstants.crptToWindow
            .CR2.Action = 2
        End With
        frmInstanciaPrincipal.Cursor = Cursors.Default
        Exit Sub

Imprimir_Error:
        If Err.Number = -2147217871 Then
            If intentos < 5 Then
                intentos = intentos + 1
                Resume
            End If
        End If
        frmInstanciaPrincipal.Cursor = Cursors.Default
        If claseBDPagos.BDWorkConnect.Errors.Count > 0 Then
            If boolTransaccion Then claseBDPagos.BDWorkConnect.RollbackTrans()
            claseBDPagos.BDWorkConnect.Errors.Clear()
            MsgBox("El proceso de transacción de la Base de Datos ha devuelto un error. La transacción será abortada", MsgBoxStyle.Information)
        Else
            MsgBox("Se ha producido un error al intentar imprimir la selección de referencias de asistencia. Llame a informática.", MsgBoxStyle.Critical)
        End If
        'objError.Tipo = Pantalla
        'objError.Ver(IdProceso, gstrError, , Codcia)
    End Sub
    ''''''''''' suplidos
    Public Sub ColorListItem(ByVal objlistitem As ListViewItem, ByVal Color As Color)

        Dim objListSubItem As ListViewItem.ListViewSubItem

        If IsNothing(objlistitem) Then Exit Sub

        objlistitem.ForeColor = Color

        For Each objListSubItem In objlistitem.SubItems
            objListSubItem.ForeColor = Color
        Next
    End Sub

    '    Public Function DatosCiaAsistencia(ByVal Codcia As String) As Boolean

    '        On Error GoTo DatosCiaAsistencia_Err

    '        ' Declaraciones
    '        '
    '        Dim strsql As String                ' Instrucción Sql

    '        strsql = "SELECT Codcia, Nombre = Isnull(Nombre,''), Nif = Isnull(Nif,''), Direccion = Isnull(Direccion,''), " & _
    '                 "       Poblacion = Isnull(Poblacion,''), CodPob = Isnull(CodPob,''), Numcia = isnull(Numcia, ''), IdRefer = Isnull(IdRefer,'') " & _
    '                 "From  MPASICIAS " & _
    '                 "Where CodCia = '" & Codcia & "'"
    '        claseBDSuplidos.BDSystemRecord.Open(strsql, claseBDSuplidos.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '        If claseBDSuplidos.BDSystemRecord.EOF Then
    '            'Err.Raise(1)
    '            DatosCiaAsistencia = False
    '        Else
    '            strNIFCompa = claseBDSuplidos.BDSystemRecord.Fields("NIF").Value
    '            strDirecCompa = claseBDSuplidos.BDSystemRecord.Fields("Direccion").Value
    '            strPoblaCompa = claseBDSuplidos.BDSystemRecord.Fields("Poblacion").Value
    '            strCodPobCompa = claseBDSuplidos.BDSystemRecord.Fields("codPob").Value
    '            strNombreCompa = claseBDSuplidos.BDSystemRecord.Fields("Nombre").Value
    '            strNumCompa = claseBDSuplidos.BDSystemRecord.Fields("Numcia").Value
    '            strIdReferCompa = claseBDSuplidos.BDSystemRecord.Fields("IdRefer").Value
    '            DatosCiaAsistencia = True
    '        End If
    '        claseBDSuplidos.BDSystemRecord.Close()
    '        Exit Function

    'DatosCiaAsistencia_Err:
    '        DatosCiaAsistencia = False
    '        If claseBDSuplidos.BDSystemRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDSystemRecord.Close()
    '        'JCLopez_i
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática.", MsgBoxStyle.Exclamation)
    '        'GlobalNumErr = "4004"
    '        'JCLopez_f
    '    End Function


    '    Public Function LlenarComboCias(ByVal objCombo As ComboBox, ByVal objlista As ListBox) As Boolean

    '        ' Declaraciones
    '        '
    '        Dim num As Integer
    '        Dim indiceLista, indiceCombo As Integer

    '        On Error GoTo LlenarComboCias_Err

    '        ' Inicialimos los objetos combo y list
    '        '
    '        objCombo.Items.Clear()
    '        objlista.Items.Clear()

    '        claseBDSuplidos.BDSystemRecord.Open("SELECT Codcia, Nombre, Inicio From MPASICIAS  where activa = 'S' ", claseBDSuplidos.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '        num = 0
    '        claseBDSuplidos.BDSystemRecord.MoveFirst()
    '        Do While Not claseBDSuplidos.BDSystemRecord.EOF
    '            strNombreCompa = Trim(claseBDSuplidos.BDSystemRecord.Fields("Nombre").Value)
    '            indiceLista = objlista.Items.Add(claseBDSuplidos.BDSystemRecord.Fields("Codcia").Value)             ' ListBox guarda el Codcia
    '            indiceCombo = objCombo.Items.Add(strNombreCompa)             ' Descripcion compañia
    '            'objCombo.ItemData(objCombo.NewIndex) = objlista.NewIndex  ' Asocia listbox a combobox
    '            objCombo.Items.Item(indiceCombo) = strNombreCompa
    '            If claseBDSuplidos.BDSystemRecord.Fields("Inicio").Value = "S" Then strCiaDefault = num
    '            claseBDSuplidos.BDSystemRecord.MoveNext()
    '            num = num + 1
    '        Loop
    '        claseBDSuplidos.BDSystemRecord.Close()
    '        LlenarComboCias = True
    '        Exit Function

    'LlenarComboCias_Err:
    '        LlenarComboCias = False
    '        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()
    '        'GlobalNumErr = "4004"
    '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")
    '    End Function


    Public Function CargarListView_suplidos(ByRef lvwControl As ListView, ByRef strsql As String, ByRef CampoKey As String, ByVal ParamArray Campos() As Object) As Boolean

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim intContador As Short
        Dim CadenaSql As String
        Dim strCodSiniestro As String
        Dim numFila As Short

        On Error GoTo CargarListView_suplidosError

        CargarListView_suplidos = True

        Cursor.Current = Cursors.WaitCursor

        numFila = 0

        lvwControl.Items.Clear()
        claseBDSuplidos.BDWorkConnect.Errors.Clear()
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()

        claseBDSuplidos.BDWorkRecord.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDSuplidos.BDWorkRecord.EOF Then
            claseBDSuplidos.BDWorkRecord.MoveFirst()
            Do
                For intContador = LBound(Campos) To UBound(Campos)
                    ' "" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value
                    If intContador = LBound(Campos) Then
                        strCodSiniestro = claseBDSuplidos.BDWorkRecord.Fields(Campos(intContador)).Value
                        lvwControl.Items.Add(strCodSiniestro)
                        'objListItem.Text = "" & claseBDPagos.BDWorkRecord.Fields(Campos(intContador)).Value
                    Else
                        If claseBDSuplidos.BDWorkRecord.Fields(Campos(intContador)).Type = ADODB.DataTypeEnum.adNumeric And Campos(intContador) <> "T7_CODRAM" And Campos(intContador) <> "T2_NUMORD" Then
                            lvwControl.Items(numFila).SubItems.Add("" & Format(claseBDSuplidos.BDWorkRecord.Fields(Campos(intContador)).Value, "##,##0.00"))
                        Else
                            lvwControl.Items(numFila).SubItems.Add("" & claseBDSuplidos.BDWorkRecord.Fields(Campos(intContador)).Value)
                        End If
                    End If
                Next
                claseBDSuplidos.BDWorkRecord.MoveNext()
                numFila = numFila + 1
            Loop Until claseBDSuplidos.BDWorkRecord.EOF
        Else
            GlobalNumErr = "4005"
            Err.Raise(Val(GlobalNumErr))
        End If

        Cursor.Current = Cursors.Default

        Exit Function

CargarListView_suplidosError:
        Cursor.Current = Cursors.Default
        CargarListView_suplidos = False
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()
        MsgBox("No existen datos para mostrar con los criterios de selección asignados", MsgBoxStyle.Exclamation)
        'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
        'objError.Ver(IdProceso, "4005", , strCodcia)
    End Function

    Public Function CruceReferencias(ByRef RefCompa As String, ByRef IdCompa As String) As Boolean

        On Error GoTo CruceReferencias_Err

        Dim strsql As String
        Dim objcmd As ADODB.Command
        Dim lngRegistros As Integer

        strsql = "UPDATE SuplidosAsistencia " & "SET    T7_CODSIN = snsinies.CODSIN " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & RefCompa & "' + SuplidosAsistencia.T7_REFER AND " & "       (SuplidosAsistencia.T7_CODSIN = '' OR SuplidosAsistencia.T7_CODSIN IS NULL OR SuplidosAsistencia.T7_CODSIN = 'No Existe') and " & "       T7_Codcia = '" & IdCompa & "'"

        objcmd = New ADODB.Command
        objcmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        objcmd.CommandText = strsql
        objcmd.ActiveConnection = claseBDSuplidos.BDWorkConnect
        objcmd.Execute(lngRegistros)

        ' Después marcamos los que no han sido encontrados como 'No Existe'
        '
        strsql = "UPDATE SuplidosAsistencia " & "SET    T7_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & RefCompa & "' + SuplidosAsistencia.T7_REFER AND " & "       ((SuplidosAsistencia.T7_CODSIN = '' OR SuplidosAsistencia.T7_CODSIN IS NULL) and SuplidosAsistencia.T7_CODSIN <> 'No Existe') and " & "       T7_Codcia = '" & IdCompa & "'"

        objcmd = New ADODB.Command
        objcmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        objcmd.CommandText = strsql
        objcmd.ActiveConnection = claseBDSuplidos.BDWorkConnect
        objcmd.Execute(lngRegistros)

        CruceReferencias = True
        Exit Function

CruceReferencias_Err:
        CruceReferencias = False
    End Function

    Public Sub Imprimir_suplidos()
        Dim strError As String
        On Error GoTo GeneraTMP_Err

        ' Declaraciones
        '
        Dim strSeleccion As String

        '/*MUL INI
        ' Si no se ha realizado la busqueda no se ha generado la query no se puede imprimir
        If Trim(strSQLCR) = "" Then
            MsgBox("No existen datos para imprimir", MsgBoxStyle.Information)
            Exit Sub
        End If
        '/*MUL FIN

        Cursor.Current = Cursors.WaitCursor

        'Generación de nuevos registros temporales en mdpImpAperAsis
        '
        With claseBDSuplidos.BDWorkConnect
            Transaccion = True
            If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
            .BeginTrans()
            .Execute("DELETE FROM mdpImpSuplidosAsis ")
            .Execute("INSERT INTO mdpImpSuplidosAsis (Codcia,Codsin,Refer,Codram,Numpol,Estado,FechaGrab,Fichero,FechaProceso,NumFactura,FecFactura)" & strSQLCR)
            .CommitTrans()
            Transaccion = False
        End With

        ' Parametrización y lanzamiento del informe
        '
        With frmInstSuplidos

            If .FiltroAviso.Checked Then
                strSeleccion = "Se han seleccionado los suplidos con avisos"
            End If

            If .FiltroErrores.Checked Then
                strSeleccion = "Se han seleccionado los suplidos con errores"
            End If

            If .FiltroNoPagados.Checked Then
                strSeleccion = "Se han seleccionado los suplidos pendientes de pago"
            End If
            If .FiltroPagados.Checked Then
                strSeleccion = "Se han seleccionado los suplidos pagados"
            End If
            If .FiltroTodos.Checked Then
                strSeleccion = "Se han seleccionado todos los suplidos"
            End If

            .CR2.set_ParameterFields(1, "FecDesde;" & CStr(.dtpDesde.Value) & ";TRUE")
            .CR2.set_ParameterFields(2, "FecHasta;" & CStr(.dtpHasta.Value) & ";TRUE")
            .CR2.set_ParameterFields(4, "Compañia;" & CStr(.cbxCompania.Text) & ";TRUE")
            .CR2.set_ParameterFields(5, "Seleccion;" & strSeleccion$ & ";TRUE")
            .CR2.Destination = Crystal.DestinationConstants.crptToWindow
            .CR2.Action = 2

            .Show()
        End With
        Cursor.Current = Cursors.Default
        Exit Sub

GeneraTMP_Err:
        Cursor.Current = Cursors.Default
        If claseBDSuplidos.BDWorkConnect.Errors.Count > 0 Then
            If Transaccion Then claseBDSuplidos.BDWorkConnect.RollbackTrans()
            gstrError = "4009"
            strError = "El proceso de transacción de la Base de Datos ha devuelto un error. La transacción será abortada"
            claseBDSuplidos.BDWorkConnect.Errors.Clear()
        Else
            gstrError = "4010"
            strError = "Se ha producido un error al intentar imprimir la selección de referencias de asistencia. Llame a informática."
        End If

        MsgBox(strError, MsgBoxStyle.Critical)
    End Sub
    ''''' cierres
    'Public Sub ColorListItem(ByVal objlistitem As ListViewItem, ByVal Color As Color)

    '    Dim objListSubItem As ListViewItem.ListViewSubItem

    '    If IsNothing(objlistitem) Then Exit Sub

    '    objlistitem.ForeColor = Color

    '    For Each objListSubItem In objlistitem.SubItems
    '        objListSubItem.ForeColor = Color
    '    Next
    'End Sub

    'Public Sub ActualizarPorcentaje(ByVal Total As Long)

    '    Dim intPorcentaje As Integer
    '    Dim lbprbProgreso As ProgressBar

    '    On Error Resume Next

    '    lbprbProgreso = frmInstanciaPrincipal.prbProgreso

    '    If Not Total = -1 Then ' Actualizar barra de estado, de progreso y porcentaje
    '        lbprbProgreso.Visible = True
    '        If Val(lbprbProgreso.Value) <= Val(lbprbProgreso.Maximum) Then
    '            lbprbProgreso.Value = lbprbProgreso.Value + 1
    '        End If

    '        intPorcentaje = Math.Round((lbprbProgreso.Value * 100) / Total, 0)
    '        If CStr(intPorcentaje) & " %" <> frmInstanciaPrincipal.stbEstado.Panels(2).Text Then
    '            frmInstanciaPrincipal.stbEstado.Panels(2).Text = CStr(intPorcentaje) & " %"
    '        End If
    '    Else
    '        lbprbProgreso.Visible = False
    '    End If

    'End Sub


    'Public Sub ActualizarPorcentaje(ByVal Total As Long)

    '    Dim intPorcentaje As Integer
    '    Dim lbprbProgreso As ProgressBar

    '    On Error Resume Next

    '    lbprbProgreso = frmInstSuplidos.prbProgreso

    '    If Not Total = -1 Then ' Actualizar barra de estado, de progreso y porcentaje
    '        lbprbProgreso.Visible = True
    '        If Val(lbprbProgreso.Value) <= Val(lbprbProgreso.Maximum) Then
    '            lbprbProgreso.Value = lbprbProgreso.Value + 1
    '        End If

    '        intPorcentaje = Math.Round((lbprbProgreso.Value * 100) / Total, 0)
    '        If CStr(intPorcentaje) & " %" <> frmInstSuplidos.stbEstado.Panels(2).Text Then
    '            frmInstSuplidos.stbEstado.Panels(2).Text = CStr(intPorcentaje) & " %"
    '        End If
    '    Else
    '        lbprbProgreso.Visible = False
    '    End If

    'End Sub

    Public Sub ActualizarPorcentaje(ByVal Total As Long, _
                                    ByRef lbprbProgreso As ProgressBar, _
                                    ByRef pStatusBar As StatusBar)

        Dim intPorcentaje As Integer
        'Dim lbprbProgreso As ProgressBar

        On Error Resume Next

        'lbprbProgreso = frmInstanciaPrincipal.prbProgreso

        If Not Total = -1 Then ' Actualizar barra de estado, de progreso y porcentaje
            lbprbProgreso.Visible = True
            If Val(lbprbProgreso.Value) <= Val(lbprbProgreso.Maximum) Then
                lbprbProgreso.Value = lbprbProgreso.Value + 1
            End If

            intPorcentaje = Math.Round((lbprbProgreso.Value * 100) / Total, 0)
            If CStr(intPorcentaje) & " %" <> pStatusBar.Panels(2).Text Then
                pStatusBar.Panels(2).Text = CStr(intPorcentaje) & " %"
            End If
        Else
            lbprbProgreso.Visible = False
        End If

    End Sub

    'Friend Function FiltrarRegistros(ByVal iIndice As Integer)
    '    ' Declaraciones
    '    Dim rbBoton As RadioButton

    '    ' Asignación del parametro a buscar según el filtro seleccionado
    '    Select Case iIndice
    '        Case 0
    '            strFiltro = "T"
    '        Case 1
    '            strFiltro = "P"
    '        Case 2
    '            strFiltro = "X"
    '        Case 3
    '            strFiltro = "A"
    '        Case 4
    '            strFiltro = "E"
    '    End Select

    '    If bwflag Then

    '        ' Antes de ejeutar el filtro comprobamos que
    '        ' se ha seleccionado una compañía
    '        '
    '        If frmInstAperturas.cbxCompania.Text = "" Then
    '            MsgBox("No se ha seleccionado ninguna compañía de asistencia", MsgBoxStyle.Information)
    '            Exit Function
    '        End If

    '        ' Ejecución del filtro
    '        frmInstAperturas.RefrescarGrid(frmInstAperturas.dtpDesde.Value, frmInstAperturas.dtpHasta.Value, frmInstAperturas.cbxProducto.Text, frmInstanciaPrincipal.lbxCompania.Items.Item(frmInstanciaPrincipal.cbxCompania.SelectedIndex))
    '    End If
    '    bwflag = True
    'End Function

    'Friend Function FiltrarRegistros(ByVal iIndice As Integer)
    '    ' Declaraciones
    '    Dim rbBoton As RadioButton

    '    ' Asignación del parametro a buscar según el filtro seleccionado
    '    '
    '    Select Case iIndice
    '        Case 0
    '            strFiltro = "T"
    '        Case 1
    '            strFiltro = "P"
    '        Case 2
    '            strFiltro = "X"
    '        Case 3
    '            strFiltro = "A"
    '        Case 4
    '            strFiltro = "E"
    '    End Select

    '    If bwflag Then

    '        ' Antes de ejeutar el filtro comprobamos que
    '        ' se ha seleccionado una compañía
    '        '
    '        If frmInstanciaPrincipal.cbxCompania.Text = "" Then
    '            MsgBox("No se ha seleccionado ninguna compañía de asistencia", MsgBoxStyle.Information)
    '            'gstrError = "4011"
    '            'objError.Tipo = Pantalla
    '            'objError.Ver(IdProceso, gstrError, , Codcia)
    '            Exit Function
    '        End If

    '        ' Ejecución del filtro
    '        frmInstPagos.RefrescarGrid(frmInstPagos.dtpDesde.Value, frmInstPagos.dtpHasta.Value, frmInstPagos.cbxProducto.Text, _
    '                                   frmInstPagos.lbxCompania.Items.Item(frmInstPagos.cbxCompania.SelectedIndex))
    '    End If
    '    bwflag = True
    'End Function

    Friend Function FiltrarRegistros(ByVal iIndice As Integer, _
                                     ByRef pCmbCompania As ComboBox) As Boolean
        ' Declaraciones
        'Dim rbBoton As RadioButton
        FiltrarRegistros = False

        ' Asignación del parametro a buscar según el filtro seleccionado
        Select Case iIndice
            Case 0
                strFiltro = "T"
            Case 1
                strFiltro = "P"
            Case 2
                strFiltro = "X"
            Case 3
                strFiltro = "A"
            Case 4
                strFiltro = "E"
        End Select

        If bwflag Then
            ' Antes de ejeutar el filtro comprobamos que se ha seleccionado una compañía
            If pCmbCompania.Text = "" Then
                MsgBox("No se ha seleccionado ninguna compañía de asistencia", MsgBoxStyle.Information)
                'gstrError = "4011"
                'objError.Tipo = Pantalla
                'objError.Ver(IdProceso, gstrError, , Codcia)
                Exit Function
            End If

            ' Ejecución del filtro
            'strSigCompa = pLbxCompania.Items.Item(pCmbCompania.SelectedIndex)

            'frmInstSuplidos.RefrescarGrid(frmInstSuplidos.dtpDesde.Value, frmInstSuplidos.dtpHasta.Value, strSigCompa)
        End If
        bwflag = True
        FiltrarRegistros = True
    End Function


    'Friend Function FiltrarRegistros(ByVal iIndice As Integer)
    '    ' Declaraciones
    '    Dim rbBoton As RadioButton

    '    If bwflag Then

    '        ' Antes de ejeutar el filtro comprobamos que
    '        ' se ha seleccionado una compañía
    '        '
    '        If frmInstCierres.cbxCompania.Text = "" Then
    '            MsgBox("No se ha seleccionado ninguna compañía de asistencia", MsgBoxStyle.Information)
    '            'gstrError = "4011"
    '            'objError.Tipo = Pantalla
    '            'objError.Ver(IdProceso, gstrError, , Codcia)
    '            Exit Function
    '        End If

    '        ' Asignación del parametro a buscar según el filtro seleccionado
    '        '
    '        Select Case iIndice
    '            Case 0
    '                strFiltro = "T"
    '            Case 1
    '                strFiltro = "P"
    '            Case 2
    '                strFiltro = "X"
    '            Case 3
    '                strFiltro = "A"
    '            Case 4
    '                strFiltro = "E"
    '        End Select

    '        ' Ejecución del filtro
    '        frmInstCierres.RefrescarGrid(frmInstCierres.dtpDesde.Value, frmInstCierres.dtpHasta.Value, frmInstCierres.cbxProducto.Text, frmInstCierres.lbxCompania.Items.Item(frmInstanciaPrincipal.cbxCompania.SelectedIndex))
    '    End If
    '    bwflag = True
    'End Function


    Public Function LlenarComboCias(ByVal objCombo As ComboBox, ByVal objlista As ListBox) As Boolean

        ' Declaraciones
        '
        Dim num As Integer
        Dim indiceLista, indiceCombo As Integer

        On Error GoTo LlenarComboCias_Err

        ' Inicialimos los objetos combo y list
        '
        objCombo.Items.Clear()
        objlista.Items.Clear()


        claseBDCierres.BDSystemRecord.Open("SELECT Codcia, Nombre, Inicio From MPASICIAS  where activa = 'S'", claseBDCierres.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        num = 0
        claseBDCierres.BDSystemRecord.MoveFirst()
        Do While Not claseBDCierres.BDSystemRecord.EOF
            strNombreCompa = Trim(claseBDCierres.BDSystemRecord.Fields("Nombre").Value)
            indiceLista = objlista.Items.Add(claseBDCierres.BDSystemRecord.Fields("Codcia").Value)             ' ListBox guarda el Codcia
            indiceCombo = objCombo.Items.Add(strNombreCompa)             ' Descripcion compañia
            objCombo.Items.Item(indiceCombo) = strNombreCompa
            If claseBDCierres.BDSystemRecord.Fields("Inicio").Value = "S" Then strCiaDefault = num
            claseBDCierres.BDSystemRecord.MoveNext()
            num = num + 1
        Loop
        claseBDCierres.BDSystemRecord.Close()
        LlenarComboCias = True
        Exit Function

LlenarComboCias_Err:
        LlenarComboCias = False
        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        'GlobalNumErr = "4004"
        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática")
    End Function


    Public Function CargarListView_cierres(ByRef lvwControl As ListView, ByRef strsql As String, ByRef CampoKey As String, ByVal ParamArray Campos() As Object) As Boolean

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim intContador As Short
        Dim Importe, Iva As Double
        Dim ImporteTotal As Double
        Dim TiempoEspera As Short
        Dim CadenaSQL, strCodSiniestro As String
        Dim numFila As Long

        On Error GoTo CargarListView_cierresError

        Importe = 0
        Iva = 0
        ImporteTotal = 0

        CargarListView_cierres = True

        Cursor.Current = Cursors.WaitCursor

        lvwControl.Items.Clear()
        claseBDCierres.BDWorkConnect.CommandTimeout = 0
        claseBDCierres.BDWorkConnect.Errors.Clear()
        If claseBDCierres.BDAuxRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDAuxRecord.Close()

        claseBDCierres.BDAuxRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDCierres.BDAuxRecord.EOF Then
            Do

                For intContador = LBound(Campos) To UBound(Campos)
                    If intContador = LBound(Campos) Then
                        strCodSiniestro = claseBDCierres.BDAuxRecord.Fields(Campos(intContador)).Value
                        lvwControl.Items.Add(strCodSiniestro)
                    Else
                        lvwControl.Items(numFila).SubItems.Add("" & claseBDCierres.BDAuxRecord.Fields(Campos(intContador)).Value)
                    End If

                Next
                lvwControl.Items(numFila).Selected = False
                numFila = numFila + 1

                If Not claseBDCierres.BDAuxRecord.EOF Then claseBDCierres.BDAuxRecord.MoveNext()

            Loop Until claseBDCierres.BDAuxRecord.EOF
        Else
            strGlobalNumErr = "4005"
            Err.Raise(Val(strGlobalNumErr))
        End If

        Cursor.Current = Cursors.Default

        Exit Function

CargarListView_cierresError:
        If Err.Number = -2147217871 Then
            If TiempoEspera < 5 Then
                TiempoEspera = TiempoEspera + 1
                System.Windows.Forms.Application.DoEvents()
                Resume
            End If
        End If
        If Err.Number = -2147217887 Then
            If TiempoEspera < 5 Then
                TiempoEspera = TiempoEspera + 1
                System.Windows.Forms.Application.DoEvents()
                claseBDCierres.BDWorkConnect.Errors.Clear()
                If claseBDCierres.BDAuxRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDAuxRecord.Close()
                Resume
            End If
        End If
        Cursor.Current = Cursors.Default
        CargarListView_cierres = False
        If claseBDCierres.BDAuxRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDAuxRecord.Close()

        If TiempoEspera >= 2 Then
            MsgBox("Por algún motivo el tiempo de acceso a la Base de Datos se encuentra anormalmente ralentizado. Si el resultado de la consulta no es satisfactorio intentelo un poco más tarde.", MsgBoxStyle.Critical)
        Else
            MsgBox("No existen datos para mostrar con los criterios de selección asignados")
        End If
    End Function

    '    Public Function DatosCiaAsistencia(ByRef Codcia As String) As Boolean

    '        On Error GoTo DatosCiaAsistencia_Err

    '        ' Declaraciones
    '        '
    '        Dim strsql As String ' Instrucción Sql

    '        strsql = "SELECT Codcia, Nombre = Isnull(Nombre,''), Nif = Isnull(Nif,''), Direccion = Isnull(Direccion,''), " & "       Poblacion = Isnull(Poblacion,''), CodPob = Isnull(CodPob,''), Numcia = isnull(Numcia, ''), IdRefer = Isnull(IdRefer,'') " & "From  MPASICIAS " & "Where CodCia = '" & Codcia & "'"
    '        claseBDCierres.BDSystemRecord.Open(strsql, claseBDCierres.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '        If claseBDCierres.BDSystemRecord.EOF Then
    '            Err.Raise(1)
    '        Else
    '            strNIFCompa = claseBDCierres.BDSystemRecord.Fields("NIF").Value
    '            strDirecCompa = claseBDCierres.BDSystemRecord.Fields("Direccion").Value
    '            strPoblaCompa = claseBDCierres.BDSystemRecord.Fields("Poblacion").Value
    '            strCodPobCompa = claseBDCierres.BDSystemRecord.Fields("codPob").Value
    '            strNombreCompa = claseBDCierres.BDSystemRecord.Fields("Nombre").Value
    '            strNumCompa = claseBDCierres.BDSystemRecord.Fields("Numcia").Value
    '            strIdReferCompa = claseBDCierres.BDSystemRecord.Fields("IdRefer").Value
    '        End If
    '        claseBDCierres.BDSystemRecord.Close()
    '        DatosCiaAsistencia = True
    '        Exit Function

    'DatosCiaAsistencia_Err:
    '        DatosCiaAsistencia = False
    '        If claseBDCierres.BDSystemRecord.State = 1 Then claseBDCierres.BDSystemRecord.Close()
    '        strGlobalNumErr = "4004"
    '    End Function

    ' Esta función informa si el siniestros pasado en el parametro esta
    ' pendiente de procesar del ultimo proceso lanzado
    '
    Public Function EstaPendiente(ByRef sCodsin As ListViewItem) As Boolean

        Dim i As Integer

        EstaPendiente = False

        For i = 1 To colSiniestrosPendientes.Count()
            If sCodsin.Text = colSiniestrosPendientes.Item(i) Then
                EstaPendiente = True
                Exit For
            End If
        Next

    End Function

    'Funcion que muestra el listado con los cierres
    Public Sub Imprimir_cierres()

        On Error GoTo GeneraTMP_Err

        ' Declaraciones
        '
        Dim strSeleccion, strCierre, strReferencia As String
        Dim i As Short

        '/*MUL INI
        ' Si no se ha realizado la busqueda no se ha generado la query no se puede imprimir
        If Trim(strSQLCR) = "" Then
            MsgBox("No existen datos para imprimir", MsgBoxStyle.Information)
            Exit Sub
        End If
        '/*MUL FIN


        Cursor.Current = Cursors.WaitCursor

        If frmInstCierres.rbSiniestros.Checked Then

            'Generación de nuevos registros temporales en mdpImpCierreAsis
            '
            With claseBDCierres.BDWorkConnect
                boolTransaccion = True
                If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()

                .BeginTrans()
                .Execute("DELETE FROM mdpImpCierreAsis ")
                .Execute("INSERT INTO mdpImpCierreAsis (Codcia,Codsin,Refer,Codram,Numpol,Estado,Situacion,Cierre)" & strSQLCR)
                .Execute("UPDATE mdpImpCierreAsis  Set Nombrecia = '" & strNombreCompa & "' Where Codcia = '" & strCodCia & "'")
                For i = 0 To frmInstCierres.lvwCierres.Items.Count - 1
                    .Execute("UPDATE mdpImpCierreAsis  Set Cierre = '" & frmInstCierres.lvwCierres.Items.Item(i).SubItems(6).Text & "' Where Codcia = '" & strCodCia & "' and Refer ='" & frmInstCierres.lvwCierres.Items.Item(i).SubItems(1).Text & "'")
                Next
                .CommitTrans()
                boolTransaccion = False
            End With

            ' Parametrización y lanzamiento del informe
            '
            With frmInstCierres
                If .FiltroAviso.Checked Then
                    strSeleccion = "Se han seleccionado los siniestros con avisos"
                End If

                If .FiltroErrores.Checked Then
                    strSeleccion = "Se han seleccionado los siniestros con errores"
                End If

                If .FiltroNoPagados.Checked Then
                    strSeleccion = "Se han seleccionado los siniestros pendientes de pago"
                End If
                If .FiltroPagados.Checked Then
                    strSeleccion = "Se han seleccionado los siniestros pagados"
                End If
                If .FiltroTodos.Checked Then
                    strSeleccion = "Se han seleccionado todos los siniestros"
                End If
                .CR2.ReportFileName = PathReports & "siniCierres.rpt"
                .CR2.set_ParameterFields(1, "FecDesde;" & CStr(.dtpDesde.Value) & ";TRUE")
                .CR2.set_ParameterFields(2, "FecHasta;" & CStr(.dtpHasta.Value) & ";TRUE")
                .CR2.set_ParameterFields(3, "Producto;" & .cbxProducto.Text & ";TRUE")
                .CR2.set_ParameterFields(6, "Seleccion;" & strSeleccion & ";TRUE")
                .CR2.set_ParameterFields(4, "Tipo;" & frmInstCierres.cbxTipoFecha.Text & ";TRUE")
                .CR2.set_ParameterFields(5, "FecCierre;" & CStr(FechaCierre) & ";TRUE")
                .CR2.Destination = Crystal.DestinationConstants.crptToWindow
                .CR2.Action = 1
            End With
        Else
            ' Generación de nuevos registros temporales en mdpImpAnulacionesAsis
            '
            With claseBDCierres.BDWorkConnect
                ' Construimos la sentencia Sql en base a la opción determinada en el menú
                '
                If frmInstCierres.rbAnuProvisionales.Checked Then
                    strSQLCR = "Select T5_Codsin Codsin , T5_Refer Referencia, T5_Estado Estado, T5_CodRechazo CodRechazo, T5_Codcia Codcia, T5_Descripcion Descripcion, T5_Comentarios Observaciones, T5_Fecanula FecAnulacion " & _
                               "From   AnulacionesAsistencia " & _
                               "Where  T5_Codcia = '" & strCodCia & _
                               "' and (T5_Estado <> 'P' and T5_Denega <> 'S')  " & _
                               "  and     AnulacionesAsistencia.T5_Fgraba = '" & claseUtilidadesCierres.FormatoFechaSQL(frmInstCierres.dtpFechaAnulaciones.Value, False, False) & _
                               "' and " & "       T5_Tipmov = '5' and T5_Codcia = '" & strCodCia & "'" & _
                               "Order By T5_Codsin"
                Else
                    strSQLCR = "Select T5_Codsin Codsin , T5_Refer Referencia, T5_Estado Estado, T5_CodRechazo CodRechazo, T5_Codcia Codcia, T5_Descripcion Descripcion, T5_Comentarios Observaciones, T5_Fecanula FecAnulacion " & _
                               "From   AnulacionesAsistencia " & _
                               "Where  T5_Codcia = '" & strCodCia & "' and (T5_Estado <> 'P' and T5_Denega <> 'S')  " & _
                               "  and     AnulacionesAsistencia.T5_Fgraba = '" & claseUtilidadesCierres.FormatoFechaSQL(frmInstCierres.dtpFechaAnulaciones.Value, False, False) & "' and " & "       T5_Tipmov = '6' and T5_Codcia = '" & strCodCia & "' " & "Order By T5_Codsin"
                End If

                '   strSQLCR$ = "Select T5_Codsin Codsin , T5_Refer Referencia, T5_Estado Estado, T5_CodRechazo CodRechazo, T5_Codcia Codcia, T5_Descripcion Descripcion, T5_Comentarios Observaciones, T5_Fecanula FecAnulacion " & _
                ''               "From   AnulacionesAsistencia " & _
                ''               "Where  T5_Codcia ='" & Codcia & "' and " & _
                ''               "       (T5_Estado <> 'P' and T5_Denega <> 'S') and " & _
                ''                "       AnulacionesAsistencia.T5_Fgraba = '" & objUtiles.FormatoFechaSQL(frmCierres.FechaAnula, False, False) & "' " & _
                ''                "Order By T5_Codsin"
                boolTransaccion = True
                If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
                .BeginTrans()
                .Execute("DELETE FROM mdpImpAnulacionesAsis")
                .Execute("INSERT INTO mdpImpAnulacionesAsis (Codsin, Referencia, Estado, CodRechazo, Codcia, Descripcion, Observaciones, FecAnulacion)" & strSQLCR)
                .CommitTrans()
                boolTransaccion = False
            End With
            frmInstCierres.CR3.ReportFileName = PathReports & "siniAnulaciones.rpt"
            frmInstCierres.CR3.set_ParameterFields(1, "Cia;" & Trim(strNombreCompa) & ";TRUE")
            frmInstCierres.CR3.Destination = Crystal.DestinationConstants.crptToWindow
            frmInstCierres.CR3.Action = 1
        End If
        Cursor.Current = Cursors.Default
        Exit Sub

GeneraTMP_Err:
        Cursor.Current = Cursors.Default
        If claseBDCierres.BDWorkConnect.Errors.Count > 0 Then
            If boolTransaccion Then claseBDCierres.BDWorkConnect.RollbackTrans()
            MsgBox("El proceso de transacción de la Base de Datos ha devuelto un error. La transacción será abortada", MsgBoxStyle.Information)
            claseBDCierres.BDWorkConnect.Errors.Clear()
        Else
            MsgBox("Se ha producido un error al intentar imprimir la selección de referencias de asistencia. Llame a informática." & vbCrLf & _
                    frmInstCierres.CR2.LastErrorNumber & " - " & frmInstCierres.CR2.LastErrorString, MsgBoxStyle.Critical)
        End If
    End Sub

    '    Friend Sub ImprimirNuevo()
    '        On Error GoTo ImprimirNuevo_Error

    '        ' Declaraciones
    '        '
    '        Dim strSeleccion As String
    '        Dim i As Short

    '        Cursor.Current = Cursors.WaitCursor

    '        If frmInstanciaPrincipal.rbSiniestros.Checked Then

    '            'Generación de nuevos registros temporales en mdpImpCierreAsis
    '            '
    '            With mdpbd.BDWorkConnect
    '                Transaccion = True
    '                If .State = adStateClosed Then .Open()
    '                .BeginTrans()
    '                .Execute("DELETE FROM mdpImpCierreAsis ")
    '                .Execute("INSERT INTO mdpImpCierreAsis (Codcia,Codsin,Refer,Codram,Numpol,Estado,Situacion,Cierre)" & strSQLCR$)
    '                .Execute("UPDATE mdpImpCierreAsis  Set Nombrecia = '" & NombreCompa & "' Where Codcia = '" & Codcia & "'")
    '                For i = 1 To frmCierres.lvwPagos.ListItems.Count
    '                    .Execute("UPDATE mdpImpCierreAsis  Set Cierre = '" & frmCierres.lvwPagos.ListItems.Item(i).ListSubItems(6).Text & "' Where Codcia = '" & Codcia & "' and Refer ='" & frmCierres.lvwPagos.ListItems.Item(i).ListSubItems(1).Text & "'")
    '                Next
    '                .CommitTrans()
    '                Transaccion = False
    '            End With

    '            ' Parametrización y lanzamiento del informe
    '            '
    '            With frmCierres.DefInstance
    '                For A = 0 To .Option1.Count - 1
    '                    If .Option1(A).Checked = True Then
    '                        Select Case A
    '                            Case 0 'Todos
    '                                strSeleccion = "Se han seleccionado todos los siniestros"
    '                            Case 1 'Pagados
    '                                strSeleccion = "Se han seleccionado los siniestros pagados"
    '                            Case 2 'Pendientes de Pago
    '                                strSeleccion = "Se han seleccionado los siniestros pendientes de pago"
    '                            Case 3 'Avisos
    '                                strSeleccion = "Se han seleccionado los siniestros con avisos"
    '                            Case 4 'Errores
    '                                strSeleccion = "Se han seleccionado los siniestros con errores"
    '                        End Select
    '                    End If
    '                Next A
    '                frmCierres.DefInstance.CR2.ReportFileName = PathReports & "siniCierres.rpt"
    '                .CR2.set_ParameterFields(1, "FecDesde;" & CStr(.dtpDesde.Value) & ";TRUE")
    '                .CR2.set_ParameterFields(2, "FecHasta;" & CStr(.dtpHasta.Value) & ";TRUE")
    '                .CR2.set_ParameterFields(3, "Producto;" & CStr(.cboProducto.Text) & ";TRUE")
    '                .CR2.set_ParameterFields(6, "Seleccion;" & strSeleccion & ";TRUE")
    '                .CR2.set_ParameterFields(4, "Tipo;" & CStr(.txtTipoFecha.Text) & ";TRUE")
    '                .CR2.set_ParameterFields(5, "FecCierre;" & CStr(FechaCierre) & ";TRUE")
    '                .CR2.Destination = Crystal.DestinationConstants.crptToWindow
    '                .CR2.Action = 1
    '            End With
    '        Else
    '            ' Generación de nuevos registros temporales en mdpImpAnulacionesAsis
    '            '
    '            With mdpBDclsmdpBD_definst.BDWorkConnect
    '                ' Construimos la sentencia Sql en base a la opción determinada en el menú
    '                '
    '                If frmCierres.DefInstance.Provisionales.Checked Then
    '                    strSQLCR = "Select T5_Codsin Codsin , T5_Refer Referencia, T5_Estado Estado, T5_CodRechazo CodRechazo, T5_Codcia Codcia, T5_Descripcion Descripcion, T5_Comentarios Observaciones, T5_Fecanula FecAnulacion " & "From   AnulacionesAsistencia " & "Where  T5_Codcia = '" & Codcia & "' and (T5_Estado <> 'P' and T5_Denega <> 'S') and " & "       AnulacionesAsistencia.T5_Fgraba = '" & objUtiles.FormatoFechaSQL(CDate(frmCierres.DefInstance.FechaAnula.Text), False, False) & "' and " & "       T5_Tipmov = '5' and T5_Codcia = '" & Codcia & "'" & "Order By T5_Codsin"
    '                Else
    '                    strSQLCR = "Select T5_Codsin Codsin , T5_Refer Referencia, T5_Estado Estado, T5_CodRechazo CodRechazo, T5_Codcia Codcia, T5_Descripcion Descripcion, T5_Comentarios Observaciones, T5_Fecanula FecAnulacion " & "From   AnulacionesAsistencia " & "Where  T5_Codcia = '" & Codcia & "' and (T5_Estado <> 'P' and T5_Denega <> 'S') and " & "       AnulacionesAsistencia.T5_Fgraba = '" & objUtiles.FormatoFechaSQL(CDate(frmCierres.DefInstance.FechaAnula.Text), False, False) & "' and " & "       T5_Tipmov = '6' and T5_Codcia = '" & Codcia & "' " & "Order By T5_Codsin"
    '                End If

    '                '   strSQLCR$ = "Select T5_Codsin Codsin , T5_Refer Referencia, T5_Estado Estado, T5_CodRechazo CodRechazo, T5_Codcia Codcia, T5_Descripcion Descripcion, T5_Comentarios Observaciones, T5_Fecanula FecAnulacion " & _
    '                ''               "From   AnulacionesAsistencia " & _
    '                ''               "Where  T5_Codcia ='" & Codcia & "' and " & _
    '                ''               "       (T5_Estado <> 'P' and T5_Denega <> 'S') and " & _
    '                ''                "       AnulacionesAsistencia.T5_Fgraba = '" & objUtiles.FormatoFechaSQL(frmCierres.FechaAnula, False, False) & "' " & _
    '                ''                "Order By T5_Codsin"
    '                Transaccion = True
    '                If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
    '                .BeginTrans()
    '                .Execute("DELETE FROM mdpImpAnulacionesAsis")
    '                .Execute("INSERT INTO mdpImpAnulacionesAsis (Codsin, Referencia, Estado, CodRechazo, Codcia, Descripcion, Observaciones, FecAnulacion)" & strSQLCR)
    '                .CommitTrans()
    '                Transaccion = False
    '            End With
    '            frmCierres.DefInstance.CR2.ReportFileName = PathReports & "siniAnulaciones.rpt"
    '            frmCierres.DefInstance.CR2.Destination = Crystal.DestinationConstants.crptToWindow
    '            frmCierres.DefInstance.CR2.Action = 1
    '        End If

    '        'UPGRADE_WARNING: Screen propiedad Screen.MousePointer tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        Exit Sub

    'ImprimirNuevo_Error:
    '        'UPGRADE_WARNING: Screen propiedad Screen.MousePointer tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        If mdpBDclsmdpBD_definst.BDWorkConnect.Errors.Count > 0 Then
    '            If Transaccion Then mdpBDclsmdpBD_definst.BDWorkConnect.RollbackTrans()
    '            gstrError = "4009"
    '            mdpBDclsmdpBD_definst.BDWorkConnect.Errors.Clear()
    '        Else
    '            gstrError = "4010"
    '        End If
    '        objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
    '        objError.Ver(IdProceso, gstrError, , Codcia)
    '    End Sub


End Module
