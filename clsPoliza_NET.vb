Public Class clsPoliza_NET
    ' Variables locales para almacenar los valores de las propiedades
    '

    Private pPoliza As String ' Número e Póliza
    Private pRamo As String ' Código de Ramo
    Private pCapitalCte As Integer ' Capital Continente
    Private pCapitalCdo As Integer ' Capital Contenido
    Private pCapitalRc As Integer ' Capital de Resposabilidad Civil
    Private pFechaEfecto As Date ' Fecha de Efecto
    Private pFechaVto As Date ' Fecha de Vencimiento
    Private pFechaPoliza As Date ' Fecha de la Póliza
    Private pFechaRegistro As Date ' Fecha de Grabación de la Póliza
    Private pFechaEmision As Date ' Fecha de Emisión
    Private pMutualista As String ' Código del Mutualista
    Private pNomTomador As String ' Nombre del tomador
    Private pNifTomador As String ' Nif del tomador
    Private pDomicTomador As String ' Domicilio del tomador
    Private pTelefTomador As String ' Telefono del tomador
    Private pCPTomador As String ' Codigo Postal del tomador
    Private pPobTomador As String ' Población del tomador
    Private pRiesgo As String ' Riesgo
    Private pAnulacion As String ' Indicador de anulación
    Private pPrimaMinima As String ' Indicador de prima mínima

    ' Asigna el número de póliza de la cual se facilitarán los datos
    '

    ' Lee el número de póliza a la que pertenecen los datos
    '
    Public Property Poliza() As String
        Get
            Poliza = pPoliza
        End Get
        Set(ByVal Value As String)
            pPoliza = Value
        End Set
    End Property

    ' Asigna el ramo de la póliza
    '

    ' Lee el ramo de la póliza especificada en la propiedad poliza
    '
    Public Property Ramo() As String
        Get
            Ramo = pRamo
        End Get
        Set(ByVal Value As String)
            pRamo = Value
        End Set
    End Property

    ' Lee el valor del Capital de Continente de la Póliza
    '
    Public ReadOnly Property CapitalContinente() As String
        Get
            CapitalContinente = CStr(pCapitalCte)
        End Get
    End Property

    ' Lee el valor del Capital de Contenido de la Póliza
    '
    Public ReadOnly Property CapitalContenido() As String
        Get
            CapitalContenido = CStr(pCapitalCdo)
        End Get
    End Property

    ' Lee el valor del Capital de RC de la Póliza
    '
    Public ReadOnly Property CapitalRC() As String
        Get
            CapitalRC = CStr(pCapitalRc)
        End Get
    End Property

    ' Lee la fecha de efecto de la póliza
    '
    Public ReadOnly Property FechaEfecto() As Date
        Get
            FechaEfecto = pFechaEfecto
        End Get
    End Property

    ' Lee la fecha de Vencimiento de la Póliza
    '
    Public ReadOnly Property FechaVencimiento() As Date
        Get
            FechaVencimiento = pFechaVto
        End Get
    End Property

    ' Lee la fecha de póliza de la Póliza
    '
    Public ReadOnly Property FechaPoliza() As Date
        Get
            FechaPoliza = pFechaPoliza
        End Get
    End Property

    ' Lee la fecha de registro de la Póliza
    '
    Public ReadOnly Property FechaRegistro() As Date
        Get
            FechaRegistro = pFechaRegistro
        End Get
    End Property

    ' Lee la fecha de emisión de la Póliza
    '
    Public ReadOnly Property FechaEmision() As Date
        Get
            FechaEmision = pFechaEmision
        End Get
    End Property

    ' Lee el código del mutualista de la Póliza
    '
    Public ReadOnly Property Mutualista() As String
        Get
            Mutualista = pMutualista
        End Get
    End Property

    ' Lee el código del Riesgo de la Póliza
    '
    Public ReadOnly Property Riesgo() As String
        Get
            Riesgo = pRiesgo
        End Get
    End Property

    ' Lee el indicador de anulación de la Póliza
    '
    Public ReadOnly Property Anulada() As String
        Get
            Anulada = pAnulacion
        End Get
    End Property

    ' Lee el indicador de prima mínima de la póliza
    '
    Public ReadOnly Property PrimaMinima() As String
        Get
            PrimaMinima = pPrimaMinima
        End Get
    End Property

    ' Lee el nombre del tomador de la póliza
    '
    Public ReadOnly Property NombreTomador() As String
        Get
            NombreTomador = pNomTomador
        End Get
    End Property

    ' Lee la dirección del tomador de la póliza
    '
    Public ReadOnly Property DomicilioTomador() As String
        Get
            DomicilioTomador = pDomicTomador
        End Get
    End Property

    ' Lee el Nif del tomador de la póliza
    '
    Public ReadOnly Property NifTomador() As String
        Get
            NifTomador = pNifTomador
        End Get
    End Property

    ' Lee el Código Postal del tomador de la póliza
    '
    Public ReadOnly Property CPTomador() As String
        Get
            CPTomador = pCPTomador
        End Get
    End Property

    ' Lee la población del tomador de la póliza
    '
    Public ReadOnly Property PoblacionTomador() As String
        Get
            PoblacionTomador = pPobTomador
        End Get
    End Property

    ' Lee el telefono del tomador de la póliza
    '
    Public ReadOnly Property TelefonoTomador() As String
        Get
            TelefonoTomador = pTelefTomador
        End Get
    End Property

    Public Function BuscaPoliza(ByRef nRamo As String, ByRef nPoliza As String) As clsPoliza_NET

        ' Declaraciones
        '
        Dim Sql As String

        ' Asignación inicial
        '
        Sql = "Select * From Polizaca Where Codram = '" & nRamo & "' and Numpol = '" & nPoliza

        ' Ejecutamos la select sobre el RecordSet
        '
        claseBDLibrerias.BDWorkRecord.Open(Sql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        claseBDLibrerias.BDWorkRecord.MoveFirst()
        Do While Not claseBDLibrerias.BDWorkRecord.EOF
            With claseBDLibrerias.BDWorkRecord.Fields
                pRamo = .Item("Codram").Value
                pPoliza = .Item("Numpol").Value
                pCapitalCte = .Item("Capcte").Value
                pCapitalCdo = .Item("Capcde").Value
                pCapitalRc = .Item("CapiRc").Value
                pFechaEfecto = .Item("Fecefe").Value
                pFechaVto = .Item("Fecvto").Value
                pFechaPoliza = .Item("Fecpol").Value
                pFechaRegistro = .Item("Fecgra").Value
                pFechaEmision = .Item("Fecemi").Value
                pMutualista = .Item("Numtom").Value
                pNomTomador = .Item("Nomase").Value
                pNifTomador = .Item("Nifase").Value
                pDomicTomador = .Item("Domase").Value & .Item("Domase2").Value
                pTelefTomador = .Item("Telase").Value
                pCPTomador = .Item("Copase").Value
                pPobTomador = .Item("Pobbase").Value
                pRiesgo = .Item("Codris").Value
                pAnulacion = .Item("Polanu").Value
                pPrimaMinima = .Item("Prmisn").Value
            End With
        Loop
    End Function

    Public Sub New()

    End Sub
End Class
