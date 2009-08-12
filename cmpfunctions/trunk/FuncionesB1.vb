Public Class FuncionesB1

    ''' <summary>
    ''' Versión de la biblioteca de funciones
    ''' </summary>
    ''' <remarks>Indica la fecha de liberación de la versión de la biblioteca en formato YYYYMMDD</remarks>
    Public Const VersionDeFunctions As String = "20090507"

#Region ">>> Variables Globales de Opciones <<<"


    Friend fCompany As SAPbobsCOM.Company
    Friend fPappl As SAPbouiCOM.Application



    ' CARPETAS

    ''' <summary>
    ''' Carpeta dentro de la ruta del add-on que contiene los archivos XML de los formularios
    ''' </summary>
    ''' <remarks></remarks>
    Public carpetaFormularios As String = "Forms"
    ''' <summary>
    ''' Carpeta dentro de la ruta del add-on que contiene los archivos XML de los reportes
    ''' </summary>
    ''' <remarks></remarks>
    Public carpetaReportes As String = "Forms"
    ''' <summary>
    ''' Carpeta dentro de la ruta del add-on que contiene los archivos de imágen a utilizar
    ''' </summary>
    ''' <remarks></remarks>
    Public carpetaImagenes As String = "Forms"
    ''' <summary>
    ''' Nombre del grupo de query en el que se guardarán los querys de Búsquedas Formateadas
    ''' </summary>
    ''' <remarks></remarks>
    Public grupoQueryBusqF As String = "Busquedas Formateadas"


    ' MENSAJES

    ''' <summary>
    ''' Establece si se deben registrar errores en un archivo de texto en C:"
    ''' </summary>
    ''' <remarks></remarks>
    Public mantenerLogErrores As Boolean = False
    ''' <summary>
    ''' Establece el nombre del archivo para el log de errores
    ''' </summary>
    ''' <remarks></remarks>
    Public logErroresArchivo As String = "CMP_Functions_Log.txt"
    ''' <summary>
    ''' Si está activo, se mostrarán mensajes al tener éxito en las operaciones.
    ''' </summary>
    ''' <remarks></remarks>
    Public mostrarMensajesExito As Boolean = False
    ''' <summary>
    ''' Si está activo, se mostrarán mensajes al presentarse errores en las operaciones.
    ''' </summary>
    ''' <remarks></remarks>
    Public mostrarMensajesError As Boolean = False
    ''' <summary>
    ''' Indica el tipo de mensaje a mostrar en caso de error
    ''' </summary>
    ''' <remarks></remarks>
    Public mostrarMensajesTipo As enumTipoMensaje = enumTipoMensaje.MessageBox

    ' PRECONSTRUIDOS

    ''' <summary>
    ''' Arreglo con los valores "Y" y "N"
    ''' </summary>
    ''' <remarks>Para ser usado en creación de campos, etc</remarks>
    Public YNvalues As String() = {"Y", "N"}
    ''' <summary>
    ''' Arreglo con los valores "Sí" y "No"
    ''' </summary>
    ''' <remarks>Para ser usado en creación de campos, etc</remarks>
    Public YNdescription As String() = {"Sí", "No"}


#End Region

#Region ">>> Enumeradores <<<"

    ''' <summary>
    ''' Enumerador de tipo de documento. Indica el tipo de documento que se debe evaluar.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum T_Doc
        FacturaReserva = 1
        FacturaInmediata = 2
        NotaDebito = 3
        Cotizacion = 4
        NotaCredito = 5
        DocumentoPreliminar = 6
    End Enum
    ''' <summary>
    ''' Enumerador de tipo de formulario. indica si un formulario es Estándar o de Usuario.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum T_Form
        Standard = 1
        Usuario = 2
    End Enum
    ''' <summary>
    ''' Enumerador de tipos de código de cuenta
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum fCodigosDeCuenta
        AcctCode = 0
        AcctName = 1
        FormatCode = 2
        SegmentedCode = 3
    End Enum
    ''' <summary>
    ''' Enumerador de Tipo de Servidor SQL
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum enumADOconnections
        SQL_Server_2000 = 12000
        SQL_Server_2005 = 12005
        Excel_2003 = 32003
        Excel_2007 = 32007
    End Enum
    ''' <summary>
    ''' Enumerador de Tipo de Mensaje
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum enumTipoMensaje
        MessageBox = 1
        StatusBarError = 2
        StatusBarWarning = 3
    End Enum

#End Region



    ''' <summary>
    ''' Instancia la Biblioteca de funciones de SofOS, C.A.
    ''' </summary>
    ''' <param name="objetfCompany">Un objeto SAPbobsCOM.Company instanciado para ser usado en las funciones</param>
    ''' <param name="mostrarErrores">Indica si las funciones deben mostrar mensajes al presentar errores</param>
    ''' <param name="mostrarExito">Indica si las funciones deben mostrar mensajes al tener éxito en los procesos</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal objetfCompany As SAPbobsCOM.Company, ByVal objectApplication As SAPbouiCOM.Application, Optional ByVal mostrarErrores As Boolean = False, Optional ByVal mostrarExito As Boolean = False)
        Try
            fCompany = objetfCompany
            fPappl = objectApplication
            mostrarMensajesError = mostrarErrores
            mostrarMensajesExito = mostrarExito

        Catch ex As Exception
        Finally
        End Try
    End Sub




    ' REGISTRO DE ADD-ONS

    ''' <summary>
    ''' Indica si la clase del add-on se encuentra instalada en la BD, en base a la tabla CMP_SETUP.
    ''' </summary>
    ''' <param name="addOnName">Nombre que identifica la clase</param>
    ''' <param name="addOnVersion">Versión de la clase</param>
    ''' <returns>Devuelve verdadero si el addon ya está instalado y falso si no se encuentra</returns>
    ''' <remarks></remarks>
    Public Function validarVersion(ByVal addOnName As String, ByVal addOnVersion As String) As Boolean
        Dim retorno As Boolean = False
        Try
            '1. Si LA TABLA no existe la creo
            If Not checkCampoBD("@CMP_SETUP", "U_CMP_VERS") Then
                creaTablaMD("CMP_SETUP", "Setup de AddOns de SofOS", BoUTBTableType.bott_NoObject)
                creaCampoMD("CMP_SETUP", "CMP_ADDN", "Nombre del AddOn", BoFieldTypes.db_Alpha, , 100)
                creaCampoMD("CMP_SETUP", "CMP_VERS", "Version del AddOn", BoFieldTypes.db_Alpha, , 100)
            Else
                '2. Valido que los datos de add-on y versión coincidan
                Dim VRS As SAPbobsCOM.Recordset = getRecordSet("SELECT * FROM [@CMP_SETUP] WHERE U_CMP_ADDN = '" & addOnName & "' ORDER BY U_CMP_VERS DESC")
                '3. Si coinciden retorno true, de lo contrario false
                If VRS.EoF Then
                    fPappl.MessageBox("Se creará la estructura de datos para el Add-On " & addOnName)
                Else
                    If VRS.Fields.Item("U_CMP_VERS").Value.ToString < addOnVersion Then
                        fPappl.MessageBox("Se actualizará la estructura de datos para el Add-On " & addOnName & " de versión " & VRS.Fields.Item("U_CMP_VERS").Value.ToString & " a " & addOnVersion)
                    ElseIf VRS.Fields.Item("U_CMP_VERS").Value.ToString > addOnVersion Then
                        fPappl.MessageBox("Se detectó una versión del Add-On " & addOnName & " más avanzada (" & VRS.Fields.Item("U_CMP_VERS").Value.ToString & ") instalada previamente. No se recomienda el uso de la versión que está intentando ejecutar (" & addOnVersion & ")")
                        retorno = True
                    ElseIf VRS.Fields.Item("U_CMP_VERS").Value.ToString = addOnVersion Then
                        retorno = True
                    End If
                End If
                Release(VRS)
            End If
        Catch ex As Exception
            manejaErrores(ex, "Validando Versión")
        End Try
        Return retorno
    End Function

    ''' <summary>
    ''' Ingresa la versión de la clase a la tabla CMP_SETUP.
    ''' </summary>
    ''' <param name="addOnName">Nombre de la clase</param>
    ''' <param name="addOnVersion">Versión de la clase</param>
    ''' <remarks></remarks>
    Public Sub confirmarVersion(ByVal addOnName As String, ByVal addOnVersion As String)
        Try
            ' Ejecuto insert a la tabla anexando data de add-on, versión y éxito al crear.
            ' De esta forma cuando se vuelva a ejecutar el add-on no creará los campos.
            Dim strSQL As String = ""
            strSQL += "INSERT INTO [@CMP_SETUP] "
            strSQL += "(Code, Name, [U_CMP_ADDN] ,[U_CMP_VERS]) VALUES "
            strSQL += "(" & getCorrelativo("Code", "[@CMP_SETUP]", , 1000) & ", '" & getCorrelativo("Code", "[@CMP_SETUP]", , 1000) & "', '" & addOnName & "','" & addOnVersion & "')"
            Dim RSV As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            RSV.DoQuery(strSQL)
            Release(RSV)
        Catch ex As Exception
            manejaErrores(ex, "Confirmando Versión")
        End Try
    End Sub

    ''' <summary>
    ''' Cierra cualquier instancia adicional del add-on, evitando que se ejecute simultáneamente en la misma máquina.
    ''' </summary>
    ''' <param name="procesoActual">Proceso actual (Process.GetCurrentProcess)</param>
    ''' <remarks>Para utilizar pasar en el parámetro Process.GetCurrentProcess</remarks>
    Public Sub cerrarInstancias(ByVal procesoActual As System.Diagnostics.Process, Optional ByVal autoKill As Boolean = False)
        Try
            Dim procesos() As System.Diagnostics.Process = Process.GetProcessesByName(procesoActual.ProcessName)
            For Z001 As Integer = 0 To procesos.Length - 1
                If procesos(Z001).Id <> procesoActual.Id AndAlso _
                 procesos(Z001).SessionId = Process.GetProcessById(procesoActual.Id).SessionId Then
                    procesos(Z001).Kill()
                End If
            Next
            If autoKill Then procesoActual.Kill()
        Catch ex As Exception
            manejaErrores(ex, "Cerrando Instancias")
        End Try
    End Sub

    ''' <summary>
    ''' Retorna el número de instancias del add-on que se encuentran ejecutandose al mismo tiempo.
    ''' </summary>
    ''' <param name="procesoActual">Proceso actual (Process.GetCurrentProcess)</param>
    ''' <returns>El número de instancias abiertas del proceso</returns>
    ''' <remarks>Para utilizar pasar en el parámetro Process.GetCurrentProcess</remarks>
    Public Function validarInstancias(ByVal procesoActual As System.Diagnostics.Process) As Integer
        Try
            Dim procesos() As System.Diagnostics.Process = Process.GetProcessesByName(procesoActual.ProcessName)
            Return procesos.Length
        Catch ex As Exception
            manejaErrores(ex, "Validando Instancias")
        End Try
    End Function



    ' OBJETOS

    ''' <summary>
    ''' Libera un objeto de la memoria. Se recomienda usar con objetos de meta-datos.
    ''' </summary>
    ''' <param name="myObject">Objeto a liberar</param>
    ''' <remarks></remarks>
    Public Sub Release(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            manejaErrores(ex, "Liberando Objeto")
        End Try
    End Sub



    ' XML

    ''' <summary>
    ''' Levanta un formulario desde un archivo XML ubicado en la carpeta de formularios del Add-On.
    ''' </summary>
    ''' <param name="FileName">Nombre del archivo (sin la extensión .srf) del formulario.</param>
    ''' <param name="cerrarSiExiste">Si el formulario se encuentra levantado, lo cierra.</param>
    ''' <remarks></remarks>
    Public Sub cargaFormXML(ByVal FileName As String, Optional ByVal cerrarSiExiste As Boolean = False)
        Try
            If Not FileName.Contains("<") Or Not FileName.Contains(">") Then

                Try
                    If cerrarSiExiste Then fPappl.Forms.Item(FileName.ToString).Close()
                Catch
                End Try
                Dim oXmlDoc As Xml.XmlDocument
                Dim sXmlFileName As String
                Try
                    oXmlDoc = New Xml.XmlDocument
                    sXmlFileName = System.Windows.Forms.Application.StartupPath & "\" & carpetaFormularios & "\" & FileName & ".srf"
                    oXmlDoc.Load(sXmlFileName)
                    fPappl.LoadBatchActions(CStr(oXmlDoc.InnerXml))
                Catch ex As Exception
                    Try
                        If mostrarMensajesError Then fPappl.StatusBar.SetText("CMP: Error cargando formulario: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        fPappl.Forms.Item(FileName).Close()
                    Catch exx As Exception
                    End Try
                End Try

            Else

                cargaForm(FileName)

            End If
        Catch ex As Exception
            manejaErrores(ex, "Cargando Formulario XML")

        End Try
    End Sub

    ''' <summary>
    ''' Importa a B1 un layout desde un archivo XML.
    ''' </summary>
    ''' <param name="Report">Nombre del archivo (sin la extensión .xml)</param>
    ''' <param name="igualarQuery">Indica si se debe sobreescribir el query del layout por el que se encuentra en el UserQuery del mismo nombre</param>
    ''' <param name="borrarSiExiste">Indica si el layout debe eliminarse si se encuentra creado previamente</param>
    ''' <remarks></remarks>
    Public Sub cargaReportXML(ByVal Report As String, ByVal igualarQuery As Boolean, ByVal borrarSiExiste As Boolean)
        Try

            If borrarSiExiste Then
                borraLayout(getIdLayout(Report))
            End If

            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService
            Dim oReportLayoutParam As SAPbobsCOM.ReportLayoutParams
            Dim oReportLayout As SAPbobsCOM.ReportLayout
            Dim sXmlFileName As String
            sXmlFileName = System.Windows.Forms.Application.StartupPath & "\" & carpetaReportes & "\" & Report & ".xml"

            oCmpSrv = fCompany.GetCompanyService
            oReportLayoutService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)

            Dim existe As String = getRSvalue("select DocName from rdoc where DocName = '" & Report & "'")

            If existe = "" Then
                oReportLayout = oReportLayoutService.GetDataInterfaceFromXMLFile(sXmlFileName)
                oReportLayoutParam = oReportLayoutService.AddReportLayout(oReportLayout)

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oReportLayoutParam)
                oReportLayoutParam = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oReportLayout)
                oReportLayout = Nothing
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrv)
            oCmpSrv = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oReportLayoutService)
            oReportLayoutService = Nothing
            GC.Collect()

            If igualarQuery Then
                Dim myQ As String = "UPDATE rdoc SET RDOC.QString = OUQR.QString " & _
                                    "FROM RDOC INNER JOIN OUQR ON RDOC.DocName = OUQR.QName " & _
                                    "INNER JOIN OQCN ON OQCN.CategoryId = OUQR.QCategory " & _
                                    "WHERE DocName = '" & Report & "'"
                getRecordSet(myQ)
            End If

        Catch ex As Exception
            manejaErrores(ex, "Creando Layout")

        End Try
    End Sub

    ''' <summary>
    ''' Exporta un layout de B1 a un archivo XML.
    ''' </summary>
    ''' <param name="Report">Nombre del archivo (sin la extensión .xml)</param>
    ''' <remarks></remarks>
    Public Sub exportReportXML(ByVal Report As String)
        Try
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService
            Dim oReportLayoutParam As SAPbobsCOM.ReportLayoutParams
            Dim oReportLayout As SAPbobsCOM.ReportLayout
            'get company service
            oCmpSrv = fCompany.GetCompanyService
            'get report layout service
            oReportLayoutService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
            'get Report Layout Param
            oReportLayoutParam = oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams)
            'set the report layout code
            oReportLayoutParam.LayoutCode = Report
            'get the report layout using layout code
            oReportLayout = oReportLayoutService.GetReportLayout(oReportLayoutParam)
            ' ,  , , , ,, 02, 03, 17
            Dim strSQLx As String = ""
            strSQLx = System.Windows.Forms.Application.StartupPath & "\" & carpetaReportes & "\" & getRSvalue("SELECT DocName FROM RDOC WHERE DocCode = '" & Report & "'") & ".xml"
            oReportLayout.ToXMLFile(strSQLx)

        Catch ex As Exception
            manejaErrores(ex, "Exportando Layout")

        End Try
    End Sub



    ' CREACIONES

    ''' <summary>
    ''' Crea una tabla de usuario (UDT) en B1.
    ''' </summary>
    ''' <param name="NbTabla">Código de la tabla (max 8 caracteres)</param>
    ''' <param name="DescTabla">Descripción de la tabla (30 caracteres)</param>
    ''' <param name="TablaTipo">Tipo de tabla</param>
    ''' <remarks></remarks>
    Public Sub creaTablaMD(ByVal NbTabla As String, ByVal DescTabla As String, ByVal TablaTipo As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try

            Dim iVer As Integer = 0
            oUserTablesMD = Nothing
            oUserTablesMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            If Not oUserTablesMD.GetByKey(NbTabla) Then

                Dim tablaACrear As SAPbobsCOM.UserTablesMD = fCompany.GetBusinessObject(BoObjectTypes.oUserTables)
                tablaACrear.TableName = Format(NbTabla)
                tablaACrear.TableDescription = Format(DescTabla)
                tablaACrear.TableType = TablaTipo

                Dim retX As Integer = 0
                Dim strSQLx As String = ""
                retX = tablaACrear.Add
                If Not retX = 0 Then
                    iVer = iVer + 1
                    fCompany.GetLastError(retX, strSQLx)
                    manejaErrores(strSQLx, "Creando Tabla " & NbTabla)
                Else
                    If mostrarMensajesExito Then fPappl.StatusBar.SetText("Tabla " & NbTabla & " creada con éxito", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                End If
                Release(tablaACrear)

            End If
            Release(oUserTablesMD)

        Catch ex As Exception
            manejaErrores(ex, "Creando Tabla " & NbTabla)

        End Try
    End Sub

    ''' <summary>
    ''' Crea un campo de usuario (UDF) en B1.
    ''' </summary>
    ''' <param name="NbTabla">Nombre de la tabla en la que se creará el campo (sin arroba)</param>
    ''' <param name="NbCampo">Código del campo a crear (8 caracteres)</param>
    ''' <param name="DescCampo">Descripción del campo (30 caracteres)</param>
    ''' <param name="TipoDato">Establece el tipo de dato que almacenará el campo</param>
    ''' <param name="subtipo">Sub-Tipo de campo</param>
    ''' <param name="Tamaño">Tamaño del campo</param>
    ''' <param name="Obligatorio">Establece si el campo admite o no valores nulos (requiere que se establezca un valor por defecto)</param>
    ''' <param name="validValues">Arreglo de valores string que contiene los valores válidos para el campo</param>
    ''' <param name="validDescription">Arreglo de valores string que contiene descripciones para los valores válidos para el campo</param>
    ''' <param name="valorPorDef">El valor que tomará el campo por defecto (debe ser un miembro de la lista de valores válidos)</param>
    ''' <param name="tablaVinculada">Nombre de la tabla de usuario de la cual se obtendrán los valores para el campo (sin arroba)</param>
    ''' <remarks></remarks>
    Public Sub creaCampoMD(ByVal NbTabla As String, ByVal NbCampo As String, ByVal DescCampo As String, ByVal TipoDato As SAPbobsCOM.BoFieldTypes, Optional ByVal subtipo As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal Tamaño As Integer = 10, Optional ByVal Obligatorio As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal validValues As String() = Nothing, Optional ByVal validDescription As String() = Nothing, Optional ByVal valorPorDef As String = "", Optional ByVal tablaVinculada As String = "")
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Try
            oUserFieldsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            oUserFieldsMD.TableName = NbTabla
            oUserFieldsMD.Name = NbCampo
            oUserFieldsMD.Description = DescCampo
            oUserFieldsMD.Type = TipoDato
            If TipoDato <> SAPbobsCOM.BoFieldTypes.db_Date Then oUserFieldsMD.EditSize = Tamaño
            If TipoDato = SAPbobsCOM.BoFieldTypes.db_Float Then oUserFieldsMD.SubType = subtipo

            If tablaVinculada <> "" Then
                oUserFieldsMD.LinkedTable = tablaVinculada
            Else
                If Not validValues Is Nothing Then
                    For i As Integer = 0 To validValues.Length - 1
                        If validDescription Is Nothing Then
                            oUserFieldsMD.ValidValues.Description = validValues(i)
                        Else
                            oUserFieldsMD.ValidValues.Description = validDescription(i)
                        End If
                        oUserFieldsMD.ValidValues.Value = validValues(i)
                        oUserFieldsMD.ValidValues.Add()
                    Next
                End If

                If valorPorDef <> "" Then
                    oUserFieldsMD.DefaultValue = valorPorDef
                    oUserFieldsMD.Mandatory = Obligatorio
                End If
            End If

            Dim retX As Integer = 0
            Dim strSQLx As String = ""
            retX = oUserFieldsMD.Add

            If retX <> 0 Then
                fCompany.GetLastError(retX, strSQLx)
                If Not strSQLx.Contains("exist") Then manejaErrores(strSQLx, "Creando Campo " & NbTabla & "." & NbCampo)
            Else
                If mostrarMensajesExito Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": Creado con éxito", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            oUserFieldsMD = Nothing
            GC.Collect()

            Exit Sub

        Catch ex As Exception
            manejaErrores(ex, "Creando Campo " & NbTabla & "." & NbCampo)

        End Try

    End Sub

    ''' <summary>
    ''' Crea un índice (UserKey) en una tabla específica. Un índice permite validar a nivel de metadatos la no-duplicidad de un dato o combinación de estos, además de acelerar los procesos de búsqueda.
    ''' </summary>
    ''' <param name="nombreDelIndice">Código de 8 caracteres máx que distingue al índice</param>
    ''' <param name="tablaSinArroba">Nombre de la tabla (sin arroba)</param>
    ''' <param name="camposSinU">Nombre de los campos a indexar (sin prefijo U_)</param>
    ''' <param name="esUnique">Establece si el índice permite o no valores duplicados</param>
    ''' <remarks></remarks>
    Public Sub creaIndice(ByVal nombreDelIndice As String, ByVal tablaSinArroba As String, ByVal camposSinU() As String, Optional ByVal esUnique As SAPbobsCOM.BoYesNoEnum = BoYesNoEnum.tYES)
        Try
            Dim oInd As SAPbobsCOM.UserKeysMD = fCompany.GetBusinessObject(BoObjectTypes.oUserKeys)
            Try
                Dim resI As Integer = 0
                Dim strErr As String = ""
                oInd.KeyName = nombreDelIndice
                oInd.TableName = tablaSinArroba
                oInd.Unique = esUnique
                For resI = 0 To camposSinU.Length - 1
                    oInd.Elements.ColumnAlias = camposSinU(resI)
                    If resI < camposSinU.Length - 1 Then oInd.Elements.Add()
                Next
                resI = oInd.Add()
                If resI <> 0 Then
                    strErr = fCompany.GetLastErrorDescription()
                End If
            Catch exxx As Exception
            Finally
                Release(oInd)
            End Try
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Verifica si un campo existe en la Base de datos
    ''' </summary>
    ''' <param name="Tabla">Nombre de la tabla (incluyendo arroba)</param>
    ''' <param name="Campo">Nombre del campo (incluyendo prefijo U_)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function checkCampoBD(ByVal Tabla As String, ByVal Campo As String) As Boolean
        Dim retorno As Boolean = False
        Try
            Dim strSQLBD As String
            Dim oLocalBD As SAPbobsCOM.Recordset
            oLocalBD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQLBD = "SELECT column_name "
            strSQLBD &= "FROM [" & fCompany.CompanyDB & "].INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = '" & Campo & "' AND Table_Name ='" & Tabla & "'"
            oLocalBD.DoQuery(strSQLBD)
            If oLocalBD.EoF = False Then
                retorno = True
            End If
            Release(oLocalBD)
        Catch ex As Exception
            manejaErrores(ex, "Revisando la existencia del campo " & Tabla & "." & Campo)

        End Try
        Return retorno
    End Function

    ''' <summary>
    ''' Crea un objeto definido por el usuario (UDO) en B1.
    ''' </summary>
    ''' <param name="Code">Código del UDO</param>
    ''' <param name="Name">Nombre del UDO</param>
    ''' <param name="TableName">Tabla principal del UDO (sin arroba)</param>
    ''' <param name="FindColumn1">Campo para búsqueda (sin prefijo U_)</param>
    ''' <param name="FindColumn2">Campo para búsqueda (sin prefijo U_)</param>
    ''' <param name="Cancel">Permitir Cancelar</param>
    ''' <param name="Close">Permitir Cerrar</param>
    ''' <param name="Deleted">Permitir Eliminar</param>
    ''' <param name="DefaultForm">Generar formulario por defecto</param>
    ''' <param name="Find">Permitir Buscar</param>
    ''' <param name="Log">Llevar Log</param>
    ''' <param name="objectType">Tipo de objeto correspondiente al UDO</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaUDO(ByVal Code As String, ByVal Name As String, ByVal TableName As String, ByVal FindColumn1 As String, Optional ByVal FindColumn2 As String = "", Optional ByVal Cancel As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Close As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, _
    Optional ByVal Deleted As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal DefaultForm As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Find As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Log As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal objectType As SAPbobsCOM.BoUDOObjType = BoUDOObjType.boud_MasterData) As Boolean
        Try
            '
            Dim oUserDataOMD As SAPbobsCOM.UserObjectsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            oUserDataOMD.Code = Code
            oUserDataOMD.Name = Name
            oUserDataOMD.ObjectType = objectType
            oUserDataOMD.TableName = TableName
            '
            oUserDataOMD.CanCancel = Cancel
            oUserDataOMD.CanClose = Close
            oUserDataOMD.CanDelete = Deleted
            oUserDataOMD.CanCreateDefaultForm = DefaultForm
            oUserDataOMD.CanFind = Find
            oUserDataOMD.CanLog = Log
            '
            If FindColumn1 <> "" Then
                oUserDataOMD.FindColumns.ColumnAlias = FindColumn1
            End If
            If FindColumn2 <> "" Then
                oUserDataOMD.FindColumns.ColumnAlias = FindColumn2
            End If
            If FindColumn2 <> "" Or FindColumn1 <> "" Then
                oUserDataOMD.FindColumns.Add()
            End If

            Dim ret As Integer = 0
            Dim strSQL As String = ""
            ret = oUserDataOMD.Add
            If ret <> 0 Then
                fCompany.GetLastError(ret, strSQL)
                If mantenerLogErrores Then System.IO.File.AppendAllText("C:\LogCreaUDO_" & Replace(Date.Today, "/", "-") & ".txt", Trim$(strSQL) & vbCrLf)
            End If

            Release(oUserDataOMD)

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    ''' <summary>
    ''' Crea un objeto definido por el usuario (UDO) en B1.
    ''' </summary>
    ''' <param name="Code">Código del UDO</param>
    ''' <param name="Name">Nombre del UDO</param>
    ''' <param name="TableName">Tabla principal del UDO (sin arroba)</param>
    ''' <param name="FindColumn">Arreglo de valores String que indica las columnas (sin U_) para búsqueda</param>
    ''' <param name="ChildTables">Arreglo de valores String que indica las tablas hijo (sin arroba)</param>
    ''' <param name="Cancel">Permitir Cancelar</param>
    ''' <param name="Close">Permitir Cerrar</param>
    ''' <param name="Deleted">Permitir Eliminar</param>
    ''' <param name="DefaultForm">Generar formulario por defecto</param>
    ''' <param name="Find">Permitir Buscar</param>
    ''' <param name="Log">Llevar Log</param>
    ''' <param name="objectType">Tipo de objeto correspondiente al UDO</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaUDO(ByVal Code As String, ByVal Name As String, ByVal TableName As String, Optional ByVal FindColumn As String() = Nothing, Optional ByVal ChildTables() As String = Nothing, Optional ByVal Cancel As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Close As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, _
    Optional ByVal Deleted As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal DefaultForm As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Find As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Log As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal objectType As SAPbobsCOM.BoUDOObjType = BoUDOObjType.boud_MasterData) As Boolean
        Try
            '
            Dim oUserDataOMD As SAPbobsCOM.UserObjectsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            oUserDataOMD.Code = Code
            oUserDataOMD.Name = Name
            oUserDataOMD.ObjectType = objectType
            oUserDataOMD.TableName = TableName
            '
            oUserDataOMD.CanCancel = Cancel
            oUserDataOMD.CanClose = Close
            oUserDataOMD.CanDelete = Deleted
            oUserDataOMD.CanCreateDefaultForm = DefaultForm
            oUserDataOMD.CanFind = Find
            oUserDataOMD.CanLog = Log
            '
            If Not FindColumn Is Nothing Then
                For FCi As Integer = 0 To FindColumn.Length - 1
                    oUserDataOMD.FindColumns.ColumnAlias = FindColumn(FCi)
                    oUserDataOMD.FindColumns.Add()
                Next
            End If

            If Not ChildTables Is Nothing Then
                For CTi As Integer = 0 To ChildTables.Length - 1
                    oUserDataOMD.ChildTables.TableName = ChildTables(CTi)
                    oUserDataOMD.FindColumns.Add()
                Next
            End If

            Dim ret As Integer = 0
            Dim strSQL As String = ""
            ret = oUserDataOMD.Add
            If ret <> 0 Then
                fCompany.GetLastError(ret, strSQL)
                If mantenerLogErrores Then System.IO.File.AppendAllText("C:\LogCreaUDO_" & Replace(Date.Today, "/", "-") & ".txt", Trim$(strSQL) & vbCrLf)
            End If

            Release(oUserDataOMD)

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Crea una Consulta (UserQuery) en B1.
    ''' </summary>
    ''' <param name="Nombre">Nombre con el que se identificará el Query</param>
    ''' <param name="Query">Sentencia SQL de la consulta</param>
    ''' <param name="QryCat">Nombre de la categoría en la que se registrará el query</param>
    ''' <param name="creaCat">Indica si se debe crear la categoría que contiene al query</param>
    ''' <param name="borrarSiExiste">Indica si el query debe eliminarse si se encuentra creado previamente</param>
    ''' <remarks></remarks>
    Public Sub creaQuery(ByVal Nombre As String, ByVal Query As String, ByVal QryCat As String, Optional ByVal creaCat As Boolean = False, Optional ByVal borrarSiExiste As Boolean = False)
        Try
            If borrarSiExiste Then
                borraQuery(Nombre, QryCat)
            End If
            If creaCat Then
                creaQueryCat(QryCat)
            End If
            Dim strSQLQ As String = ""
            Dim oUserQuery As SAPbobsCOM.UserQueries = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
            oUserQuery.Query = Query
            oUserQuery.QueryCategory = getIdQueryCat(QryCat)
            oUserQuery.QueryDescription = Nombre
            Dim ret As Integer = oUserQuery.Add
            If ret <> 0 Then
                fCompany.GetLastError(ret, strSQLQ)
            End If
            Release(oUserQuery)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Crea una categoría de Querys
    ''' </summary>
    ''' <param name="grupoQuery">Nombre de la categoría</param>
    ''' <param name="permisos">Permisos por grupo para la categoría de querys</param>
    ''' <remarks></remarks>
    Public Sub creaQueryCat(ByVal grupoQuery As String, Optional ByVal permisos As String = "YYYYYYYYYYYYYYYYYYYY")
        Try
            If getIdQueryCat(grupoQuery) = -1 Then
                ' la categoría no existe. la creo
                Dim gQ As SAPbobsCOM.QueryCategories = fCompany.GetBusinessObject(BoObjectTypes.oQueryCategories)
                gQ.Name = grupoQuery
                gQ.Permissions = permisos
                gQ.Add()
                Release(gQ)
            End If
        Catch ex01 As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Crea una Búsqueda Formateada
    ''' </summary>
    ''' <param name="queryName">Nombre del User Query</param>
    ''' <param name="query">Consulta SQL</param>
    ''' <param name="formID">Type del formulario</param>
    ''' <param name="itemUID">UID del item al que está vinculado la BF</param>
    ''' <param name="colUID">Columna a la que está vinculada la BF</param>
    ''' <param name="autoRefresh">Actualizar el valor automáticamente</param>
    ''' <param name="autoRefreshField">Campo que desencadena la BF</param>
    ''' <param name="borrarSiExiste">Borrar y volver a crear la BF cada vez que se inicie el Add-On</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaBusquedaF(ByVal queryName As String, ByVal query As String, ByVal formID As String, ByVal itemUID As String, Optional ByVal colUID As String = "-1", Optional ByVal autoRefresh As Boolean = False, Optional ByVal autoRefreshField As String = "", Optional ByVal borrarSiExiste As Boolean = True) As Boolean
        Dim fR As Boolean = False
        Try
            ' creo el query
            creaQuery(queryName, query, grupoQueryBusqF, True)

            ' eliminación de la BF
            Dim ret As Integer = 0
            Dim fUserBusFor2 As SAPbobsCOM.FormattedSearches
            fUserBusFor2 = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
            Dim existe As Boolean = fUserBusFor2.GetByKey(getIdBusquedaF(formID, itemUID, colUID))
            If existe And borrarSiExiste Then
                ret = fUserBusFor2.Remove()
                existe = False
            End If
            Release(fUserBusFor2)

            ' creación de la BF
            If Not existe Then
                Dim fUserBusFor As SAPbobsCOM.FormattedSearches
                fUserBusFor = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                fUserBusFor.FormID = formID
                fUserBusFor.ItemID = itemUID
                If colUID <> "-1" Then fUserBusFor.ColumnID = colUID
                fUserBusFor.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery

                fUserBusFor.QueryID = getIdQuery(queryName, getIdQueryCat(grupoQueryBusqF))
                If autoRefresh And autoRefreshField <> "" Then
                    fUserBusFor.Refresh = SAPbobsCOM.BoYesNoEnum.tYES
                    If colUID = "-1" Then
                        fUserBusFor.ByField = SAPbobsCOM.BoYesNoEnum.tYES
                    Else
                        fUserBusFor.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                    End If
                    fUserBusFor.FieldID = autoRefreshField
                    fUserBusFor.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    fUserBusFor.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                ret = fUserBusFor.Add
                If Not ret = 0 Then
                    fCompany.GetLastError(ret, query)
                End If
                Release(fUserBusFor)
            End If

            fR = True

        Catch ex As Exception
        End Try
        Return fR
    End Function

    ''' <summary>
    ''' Crea y Actualiza un Choose From List en un Item de un Formulario
    ''' </summary>
    ''' <param name="xForm">Formulario en el que se creará el CFL</param>
    ''' <param name="editItem">Item tipo EditText o Column que desplegará el CFL</param>
    ''' <param name="cflUID">Identificador único del CFL</param>
    ''' <param name="objectType">Tipo de objeto que desplegará el CFL</param>
    ''' <param name="cflAlias">Columna del CFL con el valor que se desea retornar</param>
    ''' <param name="multiSelection">Indica si se permitirá selección múltiple de registros</param>
    ''' <param name="cflConditions">Condiciones del CFL</param>
    ''' <remarks>Si el CFL existe, lo actualiza. Acepta objetos Edit y Column.</remarks>
    Public Sub creaCFL(ByVal xForm As SAPbouiCOM.Form, ByVal editItem As Object, ByVal cflUID As String, ByVal objectType As SAPbobsCOM.BoObjectTypes, ByVal cflAlias As String, Optional ByVal multiSelection As Boolean = False, Optional ByVal cflConditions As SAPbouiCOM.Conditions = Nothing)
        Try
            Dim xCFL As SAPbouiCOM.ChooseFromList
            Try ' creo el CFL
                Dim xCFLparams As SAPbouiCOM.ChooseFromListCreationParams = fPappl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                xCFLparams.MultiSelection = multiSelection
                xCFLparams.ObjectType = objectType
                xCFLparams.UniqueID = cflUID
                xCFL = xForm.ChooseFromLists.Add(xCFLparams)
            Catch ex0 As Exception ' si ya estaba creado, lo llamo
                xCFL = xForm.ChooseFromLists.Item(cflUID)
            End Try
            Try ' asigno condiciones
                If Not cflConditions Is Nothing Then
                    xCFL.SetConditions(cflConditions)
                End If
            Catch ex1 As Exception
            End Try
            ' asocio el CFL al item
            Try
                Dim xEdit As SAPbouiCOM.EditText = editItem
                xEdit.ChooseFromListUID = cflUID
                xEdit.ChooseFromListAlias = cflAlias
            Catch ex As Exception
                Dim xEdit As SAPbouiCOM.Column = editItem
                xEdit.ChooseFromListUID = cflUID
                xEdit.ChooseFromListAlias = cflAlias
            End Try
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Arma un objeto Conditions en base a un query
    ''' </summary>
    ''' <param name="elQuery">Query con los registros que se desean mostrar (el campo alias debe estar en la primera columna)</param>
    ''' <param name="campoAlias">Alias del campo por el que se realizará la comparación con el query</param>
    ''' <returns>Objeto Conditions en base a OR del camp alias contra los registros devuletos por el query</returns>
    ''' <remarks>Para su uso en CFLs</remarks>
    Public Function creaCFLCondiciones(ByVal elQuery As String, ByVal campoAlias As String) As SAPbouiCOM.Conditions
        Dim xCond As SAPbouiCOM.Conditions = New SAPbouiCOM.Conditions
        Try
            Dim xRS As SAPbobsCOM.Recordset = getRecordSet(elQuery)
            If Not xRS.EoF Then
                For ii As Integer = 0 To xRS.RecordCount - 1
                    Dim xCon As SAPbouiCOM.Condition = xCond.Add()
                    xCon.Alias = campoAlias
                    xCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    xCon.CondVal = xRS.Fields.Item(0).Value
                    xRS.MoveNext()
                    If Not xRS.EoF Then xCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                Next

            Else
                Dim xCon As SAPbouiCOM.Condition = xCond.Add()
                xCon.Alias = campoAlias
                xCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                xCon.CondVal = ""

            End If
            Release(xRS)
        Catch ex As Exception
        End Try
        Return xCond
    End Function



    ' GETS

    ''' <summary>
    ''' Devuelve un objeto UserFieldsMD lleno con los datos solicitados.
    ''' </summary>
    ''' <param name="tabla">Nombre de la tabla</param>
    ''' <param name="nombreCampo">Nombre del campo (sin prefijo U_)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCampo(ByVal tabla As String, ByVal nombreCampo As String) As SAPbobsCOM.UserFieldsMD
        Try
            Dim ufRs As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            ufRs.DoQuery("select FieldID from cufd where TableID = '" & tabla & "' and AliasID = '" & nombreCampo & "'")
            Dim k As Integer = 0
            k = ufRs.Fields.Item(0).Value.ToString
            Dim UFretorno As SAPbobsCOM.UserFieldsMD
            UFretorno = fCompany.GetBusinessObject(BoObjectTypes.oUserFields)
            UFretorno.GetByKey(tabla, k)
            If UFretorno.Name = nombreCampo Then
                Return UFretorno
                Exit Function
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuleve el ID interno de una categoría del Query Manager. Si no existe, devuelve -1.
    ''' </summary>
    ''' <param name="nombreCat">Nombre de la categoría</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdQueryCat(ByVal nombreCat As String) As Integer
        Try
            Dim queryId As Integer = -1
            Dim oLocalQ As SAPbobsCOM.Recordset
            oLocalQ = fCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oLocalQ.DoQuery("SELECT CategoryId as 'Id' FROM OQCN WHERE CatName = '" & nombreCat & "'")
            If oLocalQ.EoF = False Then queryId = oLocalQ.Fields.Item("Id").Value
            Return queryId
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un UserQuery. Si no existe, devuelve -1.
    ''' </summary>
    ''' <param name="nombreQuery">Nombre del query</param>
    ''' <param name="idCat">ID interno de la categoría del query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdQuery(ByVal nombreQuery As String, ByVal idCat As Integer) As Integer
        Try
            Dim queryId As Integer = -1
            Dim oLocalQ As SAPbobsCOM.Recordset
            oLocalQ = fCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oLocalQ.DoQuery("SELECT IntrnalKey as 'Id' FROM OUQR WHERE QName = '" & nombreQuery & "'  AND QCategory = " & idCat)
            If oLocalQ.EoF = False Then queryId = oLocalQ.Fields.Item("Id").Value
            Return queryId
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de una búsqueda formateada para ser usado en un getByKey. Si no existe, devuelve -1.
    ''' </summary>
    ''' <param name="FormID_o_TYPE">Type del formulario</param>
    ''' <param name="ItemID">UID del item al cual se encuentra ligado la BF</param>
    ''' <param name="ColID">Columna a la cual se encuentra asociada la BF (si el item es una matriz)</param>
    ''' <returns>Código interno de la búsqueda formateada</returns>
    ''' <remarks></remarks>
    Public Function getIdBusquedaF(ByVal FormID_o_TYPE As String, ByVal ItemID As String, Optional ByVal ColID As String = "-1") As Integer
        Try
            Dim oLocalBF As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strSQLBF As String = "SELECT IndexID FROM CSHS WHERE FormID='" & FormID_o_TYPE & "' and ItemID='" & ItemID & "' and ColID='" & ColID & "'"
            oLocalBF.DoQuery(strSQLBF)
            If oLocalBF.EoF = False Then
                Return oLocalBF.Fields.Item(0).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Devuelve el DocEntry de un Documento de Marketing. En caso de error, devuelve -1.
    ''' </summary>
    ''' <param name="DocNum"></param>
    ''' <param name="TipoDoc"></param>
    ''' <param name="SubType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDocEntry(ByVal DocNum As String, ByVal TipoDoc As T_Doc, Optional ByVal SubType As SAPbobsCOM.BoObjectTypes = SAPbobsCOM.BoObjectTypes.oCreditNotes) As Integer
        '
        Dim oDoc As Integer = -1
        Try
            Dim oLocalDoc As SAPbobsCOM.Recordset
            Dim strSQLDoc As String = ""
            oLocalDoc = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Select Case TipoDoc
                Case T_Doc.FacturaInmediata
                    strSQLDoc = "select docentry from oinv where docsubtype='--' and docnum='" & DocNum & "' and IsIns='N'"
                Case T_Doc.FacturaReserva
                    strSQLDoc = "select docentry from oinv where docsubtype='--' and docnum='" & DocNum & "' and IsIns='Y'"
                Case T_Doc.NotaDebito
                    strSQLDoc = "select docentry from oinv where docsubtype='DN' and docnum='" & DocNum & "'"
                Case T_Doc.Cotizacion
                    strSQLDoc = "select docentry from OQUT where docnum='" & DocNum & "'"
                Case T_Doc.NotaCredito
                    strSQLDoc = "select docentry from ORIN where docnum='" & DocNum & "'"
                Case T_Doc.DocumentoPreliminar
                    strSQLDoc = "select docentry from ODRF where docnum='" & DocNum & "' and ObjType='" & SubType & "'"
            End Select

            oLocalDoc.DoQuery(strSQLDoc)
            If oLocalDoc.EoF = False Then
                oDoc = oLocalDoc.Fields.Item("docentry").Value
            End If
        Catch ex As Exception
            oDoc = -1
            'Pappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return oDoc
        '
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un Banco. En caso de error, devuelve -1.
    ''' </summary>
    ''' <param name="BankCode">Código del Banco</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdBanco(ByVal BankCode As String) As Integer
        Dim oLocalODSC As SAPbobsCOM.Recordset
        Try
            Dim strSQL As String = "SELECT ABSENTRY FROM ODSC WHERE BANKCODE='" & BankCode & "'"
            oLocalODSC = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalODSC.DoQuery(strSQL)
            If oLocalODSC.EoF = False Then
                Return oLocalODSC.Fields.Item(0).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
        Release(oLocalODSC)
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un Campo de Usuario (UDF). En caso de error, devuelve -1.
    ''' </summary>
    ''' <param name="Tabla">Código de la tabla (con arroba)</param>
    ''' <param name="Campo">Nombre del campo (sin prefijo U_)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdUserField(ByVal Tabla As String, ByVal Campo As String) As Integer

        Dim oLocalUF As SAPbobsCOM.Recordset
        Try
            Dim strSQL As String = "SELECT FIELDID FROM CUFD WHERE TABLEID ='" & Tabla & "' AND ALIASID = '" & Campo & "'"
            oLocalUF = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalUF.DoQuery(strSQL)
            If oLocalUF.EoF = False Then
                Return oLocalUF.Fields.Item(0).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            manejaErrores(ex, "Obteniedo código del campo de usuario " & Tabla & "." & Campo)
            Return -1
        End Try
        Release(oLocalUF)

    End Function

    ''' <summary>
    ''' Devuelve el código interno de un Layout, a partir del nombre.
    ''' </summary>
    ''' <param name="layoutName">Nombre del Layout</param>
    ''' <param name="esUSR">Indica si el layout es de Usuario (código inicia por USR)</param>
    ''' <returns>Retorna el DocCode del Layout</returns>
    ''' <remarks>Si consigue varios layouts con el mismo nombre, devuelve el primer valor que encuentra. Si no consigue devuelve una cadena vacía.</remarks>
    Public Function getIdLayout(ByVal layoutName As String, Optional ByVal esUSR As Boolean = False) As String
        Dim rr As String = ""
        Try
            If Not esUSR Then
                rr = getRSvalue("SELECT DocCode FROM RDOC WHERE DocName = '" & layoutName & "'")
            Else
                rr = getRSvalue("SELECT DocCode FROM RDOC WHERE DocName = '" & layoutName & "' and DocCode >= 'USR' and DocCode <= 'USS'")
            End If
        Catch ex As Exception
            manejaErrores(ex, "Obteniendo código del layout")

        End Try
        Return rr
    End Function

    ''' <summary>
    ''' Devuelve el formulario del Item Event
    ''' </summary>
    ''' <param name="pVal">ItemEvent</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getForm(ByVal pVal As SAPbouiCOM.ItemEvent) As SAPbouiCOM.Form
        Try
            Return fPappl.Forms.Item(pVal.FormUID)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Devuelve el formulario del Data Event
    ''' </summary>
    ''' <param name="BusinessObjectInfo">Parámetro del data event</param>
    ''' <returns>Objeto SAPbouiCOM.Form del formulario activo</returns>
    ''' <remarks>En caso de error retorna Nothing</remarks>
    Public Function getForm(ByVal BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo) As SAPbouiCOM.Form
        Try
            Return fPappl.Forms.Item(BusinessObjectInfo.FormUID)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Devuelve el formulario actual
    ''' </summary>
    ''' <returns>Objeto SAPbouiCOM.Form del formulario activo</returns>
    ''' <remarks>En caso de error retorna Nothing</remarks>
    Public Function getForm() As SAPbouiCOM.Form
        Try
            Return fPappl.Forms.ActiveForm
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuleve el Item al que hace referencia el pVal
    ''' </summary>
    ''' <param name="pVal">Parámetro del ItemEvent</param>
    ''' <returns>Devuelve el objeto Item al que hace referencia el pVal. En caso de error retorna Nothing.</returns>
    ''' <remarks></remarks>
    Public Function getItem(ByVal pVal As SAPbouiCOM.ItemEvent) As SAPbouiCOM.Item
        Try
            Return getForm(pVal).Items.Item(pVal.ItemUID)
        Catch ex As Exception
            manejaErrores(ex, "Obteniedo Item del pVal")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuelve un Recordset a partir de un query
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRecordSet(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
            manejaErrores(ex, "Obteniendo Recordset")

        End Try
        Return fRS
    End Function

    ''' <summary>
    ''' Devuelve el valor de un campo de una consulta en formato String
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <param name="columnaRet">Columna de la consulta a retornar</param>
    ''' <param name="valorNulo">Valor a retornar en caso de error/nulo</param>
    ''' <returns>Devuleve un valor en específico de un query</returns>
    ''' <remarks>No requiere liberación de memoria</remarks>
    Public Function getRSvalue(ByVal query As String, ByVal columnaRet As String, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
        End Try
        Return ret
    End Function
    ''' <summary>
    ''' Devuelve el valor de un campo de una consulta en formato String
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <param name="columnaRet">Número de Columna de la consulta a retornar</param>
    ''' <param name="valorNulo">Valor a retornar en caso de error/nulo</param>
    ''' <returns>Devuleve un valor en específico de un query</returns>
    ''' <remarks>No requiere liberación de memoria</remarks>
    Public Function getRSvalue(ByVal query As String, Optional ByVal columnaRet As Integer = 0, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' Devuelve el valor seleccionado en un combo
    ''' </summary>
    ''' <param name="combo">ComboBox instanciado</param>
    ''' <param name="returnValue">Verdadero para retornar Value, Falso para description</param>
    ''' <param name="valorSiNulo">Valor a retornar si no hya nada seleccionado</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getComboSelected(ByVal combo As SAPbouiCOM.ComboBox, Optional ByVal returnValue As Boolean = True, Optional ByVal valorSiNulo As String = "") As String
        Dim r As String = valorSiNulo
        Try
            If returnValue Then
                r = combo.Selected.Value
            Else
                r = combo.Selected.Description
            End If
        Catch ex As Exception
        End Try
        Return r
    End Function

    ''' <summary>
    ''' Captura el valor seleccionado en un CFL y lo escribe en el EditText desde el que fue invocado.
    ''' </summary>
    ''' <param name="pVal">ItemEvent</param>
    ''' <param name="columna">Columna del CFL a devolver. Si se deja en blanco devuelve la primera</param>
    ''' <remarks></remarks>
    Public Sub getCFLvalue(ByVal pVal As SAPbouiCOM.ItemEvent, Optional ByVal columna As String = "")
        Try
            Dim fcflEvent As SAPbouiCOM.ChooseFromListEvent = pVal
            If columna = "" Then
                getForm(pVal).Items.Item(pVal.ItemUID).Specific.String = fcflEvent.SelectedObjects.GetValue(0, 0)
            Else
                getForm(pVal).Items.Item(pVal.ItemUID).Specific.String = fcflEvent.SelectedObjects.GetValue(columna, 0)
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Retorna datos de cuentas contables en cualquiera de 4 formatos (Código, Nombre, FormatCode y Segmentada)
    ''' </summary>
    ''' <param name="valor">Valor de la cuenta</param>
    ''' <param name="formatoOriginal">Formato del valor que se provee</param>
    ''' <param name="formatoDestino">Formato al que se desea convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAccount(ByVal valor As String, ByVal formatoOriginal As fCodigosDeCuenta, ByVal formatoDestino As fCodigosDeCuenta) As String
        Dim r As String = ""
        Try
            Dim fieldRetorno As String = ""
            Dim fieldWhere As String = ""

            If formatoOriginal = fCodigosDeCuenta.AcctCode Then fieldWhere = "AcctCode"
            If formatoOriginal = fCodigosDeCuenta.AcctName Then fieldWhere = "AcctName"
            If formatoOriginal = fCodigosDeCuenta.FormatCode Then fieldWhere = "FormatCode"
            If formatoOriginal = fCodigosDeCuenta.SegmentedCode Then
                fieldWhere = "FormatCode"
                valor = valor.Replace("-", "")
            End If

            If formatoDestino = fCodigosDeCuenta.AcctCode Then fieldRetorno = "AcctCode"
            If formatoDestino = fCodigosDeCuenta.AcctName Then fieldRetorno = "AcctName"
            If formatoDestino = fCodigosDeCuenta.FormatCode Then fieldRetorno = "FormatCode"
            If formatoDestino = fCodigosDeCuenta.SegmentedCode Then fieldRetorno = "Cuenta"

            Dim query As String = ""
            query += "SELECT AcctCode, AcctName, FormatCode, segment_0 + "
            query += "case when not segment_1 is null then '-' + segment_1 else '' end + "
            query += "case when not segment_2 is null then '-' + segment_2 else '' end + "
            query += "case when not segment_3 is null then '-' + segment_3 else '' end + "
            query += "case when not segment_4 is null then '-' + segment_4 else '' end + "
            query += "case when not segment_5 is null then '-' + segment_5 else '' end + "
            query += "case when not segment_6 is null then '-' + segment_6 else '' end + "
            query += "case when not segment_7 is null then '-' + segment_7 else '' end + "
            query += "case when not segment_8 is null then '-' + segment_8 else '' end + "
            query += "case when not segment_9 is null then '-' + segment_9 else '' end "
            query += "as Cuenta FROM OACT WHERE " & fieldWhere & " = '" & valor & "'"

            r = getRSvalue(query, fieldRetorno)

        Catch ex As Exception
            manejaErrores(ex, "Determinando Cuenta Contable")

        End Try
        Return r
    End Function

    ''' <summary>
    ''' Devuelve fecha y hora en formato Date
    ''' </summary>
    ''' <param name="Fecha">Fecha en formato String</param>
    ''' <param name="Hora">Hora y minutos como un número</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFechaHoraB1(ByVal Fecha As String, ByVal Hora As Long) As Date
        Dim FechaA As Date
        Try
            Dim Minutos As Long
            If Hora <> 0 Then
                'Explicacion
                'Hora = CInt(Hora.ToString.Remove(Hora.ToString.Length - 2, 2))
                'Minutos = CInt(Hora.ToString.Substring(Hora.ToString.Length - 2, 2))
                'Minutos = (Hora * 60) + Minutos

                Minutos = (CInt(Hora.ToString.Remove(Hora.ToString.Length - 2, 2)) * 60) + CInt(Hora.ToString.Substring(Hora.ToString.Length - 2, 2))

            Else
                Minutos = CInt(Hora.ToString.Substring(Hora.ToString.Length - 2, 2))
            End If

            FechaA = CDate(Fecha & " 00:00").AddMinutes(Minutos)

        Catch ex As Exception

        End Try

        Return FechaA

    End Function

    ''' <summary>
    ''' Devuelve Now en formato YYYYMMDD
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFechaActual() As String
        Dim mynum As String = ""
        Try
            mynum = Now.Year.ToString
            If Now.Month < 10 Then mynum += "0"
            mynum += Now.Month.ToString
            If Now.Day < 10 Then mynum += "0"
            mynum += Now.Day.ToString
        Catch ex As Exception
            manejaErrores(ex, "Obteniendo Fecha Actual")

        End Try
        Return mynum
    End Function

    ''' <summary>
    ''' Devuelve una variable date en el formato correcto en base a los valores del día, mes y año
    ''' </summary>
    ''' <param name="dia">Día de la Fecha</param>
    ''' <param name="mes">Mes de la Fecha</param>
    ''' <param name="year">Año de la Fecha</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDateVar(ByVal dia As Integer, ByVal mes As Integer, ByVal year As Integer) As Date
        Dim d As Date = Nothing
        Try
            Dim ok As Boolean = False
            Try
                d = CDate(IIf(mes < 10, "0", "") & mes & "/" & IIf(dia < 10, "0", "") & dia & "/" & year)
                If d.Day = dia And d.Month = mes And d.Year = year Then
                    ok = True
                Else
                    ok = False
                End If
            Catch ex1 As Exception
            End Try

            If Not ok Then
                d = CDate(IIf(dia < 10, "0", "") & dia & "/" & IIf(mes < 10, "0", "") & mes & "/" & year)
                If d.Day = dia And d.Month = mes And d.Year = year Then
                    ok = True
                Else
                    ok = False
                End If
            End If

            If Not ok Then
                d = Nothing
            End If

        Catch ex As Exception
        End Try
        Return d
    End Function



    ' ELIMINACIONES

    ''' <summary>
    ''' Elimina una consulta del Query Manager
    ''' </summary>
    ''' <param name="queryName">Nombre del Query</param>
    ''' <param name="grupoName">Nombre de la Categoría del Query</param>
    ''' <remarks></remarks>
    Public Sub borraQuery(ByVal queryName As String, ByVal grupoName As String)
        Try
            Dim oUQ As SAPbobsCOM.UserQueries = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
            Dim categID As Integer = getIdQueryCat(grupoName)
            Dim queryID As Integer = getIdQuery(queryName, categID)
            oUQ.GetByKey(queryID, categID)
            oUQ.Remove()
            Release(oUQ)
        Catch ex As Exception
            manejaErrores(ex, "Borrando Query " & queryName)

        End Try
    End Sub

    ''' <summary>
    ''' Elimina un Layout
    ''' </summary>
    ''' <param name="layoutCode">Código del Layout a eliminar</param>
    ''' <remarks></remarks>
    Public Sub borraLayout(ByVal layoutCode As String)
        Try
            Dim xCmpSrv As SAPbobsCOM.CompanyService
            Dim xReportLayoutService As SAPbobsCOM.ReportLayoutsService
            Dim xReportLayoutParam As SAPbobsCOM.ReportLayoutParams
            'get company service
            xCmpSrv = fCompany.GetCompanyService
            'get report layout service
            xReportLayoutService = xCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
            'get Report Layout Param
            xReportLayoutParam = xReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams)
            'set the report layout code
            xReportLayoutParam.LayoutCode = layoutCode
            xReportLayoutService.DeleteReportLayout(xReportLayoutParam)
            Release(xReportLayoutParam)
            Release(xReportLayoutService)
            Release(xCmpSrv)
        Catch ex As Exception
            manejaErrores(ex, "Borrando Layout " & layoutCode)

        End Try
    End Sub



    ' USER INTERFACE

    ''' <summary>
    ''' Dibuja un item en un formulario
    ''' </summary>
    ''' <param name="oFormItem">Objeto formulario en el que se dibujará el item</param>
    ''' <param name="TipoItem">Tipo de item a dibujar</param>
    ''' <param name="ItemID">UID del item</param>
    ''' <param name="ItemDesc">Descripción del item</param>
    ''' <param name="Left">Posición en píxeles desde la izquierda del formulario</param>
    ''' <param name="Top">Posición en píxeles desde el tope del formulario</param>
    ''' <param name="Width">Ancho del item</param>
    ''' <param name="Height">Altura del item</param>
    ''' <param name="Npanel">Panel del item</param>
    ''' <param name="DisplayDesc">Indica si la descripción debe mostrarse o no</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaControl(ByVal oFormItem As SAPbouiCOM.Form, ByVal TipoItem As SAPbouiCOM.BoFormItemTypes, ByVal ItemID As String, ByVal ItemDesc As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal Npanel As Integer = 0, Optional ByVal DisplayDesc As Boolean = False) As Boolean
        '
        Try
            Dim oItemDLL As SAPbouiCOM.Item
            oItemDLL = oFormItem.Items.Add(ItemID, TipoItem)
            With oItemDLL
                If Not Left = 0 Then .Left = Left
                If Not Width = 0 Then .Width = Width
                If Not Top = 0 Then .Top = Top
                If Not Height = 0 Then .Height = Height
                '
                ' Validación de Panels
                If Npanel <> 0 Then
                    .FromPane = Npanel
                    .ToPane = Npanel
                End If
                '
                'Se crean los rectángulos
                If Not ItemDesc = vbNullString Then
                    If TipoItem = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                        Dim oLabelDLL As SAPbouiCOM.StaticText
                        oLabelDLL = .Specific
                        oLabelDLL.Caption = ItemDesc
                    ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
                        Dim oButtonDLL As SAPbouiCOM.Button
                        oButtonDLL = .Specific
                        oButtonDLL.Caption = ItemDesc
                    ElseIf TipoItem = (SAPbouiCOM.BoFormItemTypes.it_EDIT Or SAPbouiCOM.BoFormItemTypes.it_EXTEDIT Or SAPbouiCOM.BoFormItemTypes.it_FOLDER) Then
                        Dim oEditDLL As SAPbouiCOM.EditText
                        oEditDLL = .Specific
                        oEditDLL.Caption = ItemDesc
                    Else
                        Return False
                    End If
                ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX Then
                    If DisplayDesc = True Then
                        oFormItem.Items.Item(ItemID).DisplayDesc = True
                    End If
                End If
                Return True
            End With
            '
        Catch ex As Exception
            'Pappl.MessageBox(ex.Message, 1, "Aceptar")
            Return False
            ' Pappl.MessageBox("Ya existe un objeto con ese nombre!!! ", 1, "Aceptar")
        End Try
        '
    End Function
    ''' <summary>
    ''' Dibuja un nuevo item en un formulario
    ''' </summary>
    ''' <param name="formulario">Formulario en el que se dibujará el item</param>
    ''' <param name="tipo">Tipo de item a dibujar</param>
    ''' <param name="itemUID">UID del nuevo item</param>
    ''' <param name="hOffSet">posición horizontal o desface con respecto a otro item</param>
    ''' <param name="vOffSet">posición vertical o desface con respecto a otro item</param>
    ''' <param name="respectoAItem">item en referencia al cual se dibujará</param>
    ''' <param name="ancho">ancho del nuevo item</param>
    ''' <param name="alto">alto del nuevo item</param>
    ''' <param name="copiarTamano">indica si se desea copiar el tamaño del item de referencia</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaControl(ByVal formulario As SAPbouiCOM.Form, ByVal tipo As SAPbouiCOM.BoFormItemTypes, ByVal itemUID As String, ByVal hOffSet As Integer, ByVal vOffSet As Integer, Optional ByVal respectoAItem As String = "", Optional ByVal ancho As Integer = -1, Optional ByVal alto As Integer = -1, Optional ByVal copiarTamano As Boolean = False, Optional ByVal visible As Boolean = True) As SAPbouiCOM.Item
        Try
            formulario.Freeze(True)
            formulario.Items.Add(itemUID, tipo)
            formulario.Items.Item(itemUID).Visible = visible
            If copiarTamano Then
                If respectoAItem <> "" Then
                    formulario.Items.Item(itemUID).Width = formulario.Items.Item(respectoAItem).Width
                    formulario.Items.Item(itemUID).Height = formulario.Items.Item(respectoAItem).Height
                End If
            Else
                If ancho >= 0 Then
                    formulario.Items.Item(itemUID).Width = ancho
                End If
                If alto >= 0 Then
                    formulario.Items.Item(itemUID).Height = alto
                End If
            End If
            If respectoAItem = "" Then
                formulario.Items.Item(itemUID).Top = vOffSet
                formulario.Items.Item(itemUID).Left = hOffSet
            Else

                If vOffSet > 0 Then
                    formulario.Items.Item(itemUID).Top = formulario.Items.Item(respectoAItem).Top + formulario.Items.Item(respectoAItem).Height + vOffSet
                ElseIf vOffSet = 0 Then
                    formulario.Items.Item(itemUID).Top = formulario.Items.Item(respectoAItem).Top
                Else
                    formulario.Items.Item(itemUID).Top = formulario.Items.Item(respectoAItem).Top - formulario.Items.Item(itemUID).Height + vOffSet
                End If

                If hOffSet > 0 Then
                    formulario.Items.Item(itemUID).Left = formulario.Items.Item(respectoAItem).Left + formulario.Items.Item(respectoAItem).Width + hOffSet
                ElseIf hOffSet = 0 Then
                    formulario.Items.Item(itemUID).Left = formulario.Items.Item(respectoAItem).Left
                Else
                    formulario.Items.Item(itemUID).Left = formulario.Items.Item(respectoAItem).Left - formulario.Items.Item(itemUID).Width + hOffSet
                End If
                formulario.Items.Item(itemUID).LinkTo = respectoAItem
            End If
            formulario.Refresh()
        Catch ex As Exception
        Finally
            formulario.Freeze(False)
        End Try
        Return formulario.Items.Item(itemUID)
    End Function

    ''' <summary>
    ''' Elimina un dato de un Grid
    ''' </summary>
    ''' <param name="Grid">UID del Grid</param>
    ''' <param name="col">Columna a evaluar</param>
    ''' <param name="cond">Valor a eliminar</param>
    ''' <param name="Formu">Formulario del grid</param>
    ''' <remarks></remarks>
    Public Sub borraGrid(ByVal Grid As String, ByVal col As Integer, ByVal cond As String, ByVal Formu As String)
        Dim kz As Integer = 0
        Try
            Dim fForm As SAPbouiCOM.Form = fPappl.Forms.Item(Formu)
            Dim oGridDLL As SAPbouiCOM.Grid = fForm.Items.Item(Grid).Specific

            For kz = oGridDLL.Rows.Count To 1 Step -1
                If Trim(oGridDLL.DataTable.GetValue(col, kz - 1)) = cond Then
                    oGridDLL.DataTable.Rows.Remove(kz - 1)
                End If
            Next

        Catch ex As Exception
            If mostrarMensajesError Then fPappl.StatusBar.SetText("CMP: Error borrando grid: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Cierra todas las instancias de un formulario.
    ''' </summary>
    ''' <param name="FormID">Identificador del formulario</param>
    ''' <param name="TipoForm">Tipo de formulario (Usuario o Estándar)</param>
    ''' <remarks>Puye patrocinado por Gabriel Mendes, muy feo pero se debe validar que siempre tengan la ventana y los campos correctos</remarks>
    Public Sub cierraFormularios(ByVal FormID As String, ByVal TipoForm As T_Form)
        Dim oFormM As SAPbouiCOM.Form
        Dim Count As Integer = fPappl.Forms.Count
        Dim Cantidad As New ArrayList
        '
        For il As Integer = 0 To Count - 1
            If TipoForm = T_Form.Standard Then
                If fPappl.Forms.Item(il).TypeEx = FormID Then
                    Cantidad.Add(fPappl.Forms.Item(il).TypeCount)
                End If
            Else
                If fPappl.Forms.Item(il).UniqueID = FormID Then
                    fPappl.Forms.Item(il).Close()
                End If
            End If
        Next
        If TipoForm = T_Form.Standard Then
            For il2 As Integer = 0 To Cantidad.Count - 1
                oFormM = fPappl.Forms.GetFormByTypeAndCount(FormID, Cantidad.Item(il2))
                If oFormM.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then oFormM.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                oFormM.Close()
            Next
        End If
        '
    End Sub

    ''' <summary>
    ''' Indica si un formulario se encuentra abierto.
    ''' </summary>
    ''' <param name="FormID">Identificador del formulario</param>
    ''' <param name="Tipo">Tipo de formulario (Estándar o Usuario)</param>
    ''' <param name="TypeCount">Número de instancias del formulario</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isFormLoaded(ByVal FormID As String, ByVal Tipo As T_Form, Optional ByVal TypeCount As Integer = 1) As Boolean
        '
        'Dim oFormA As SAPbouiCOM.Form
        Dim Count As Integer = fPappl.Forms.Count
        'Dim Cantidad As New ArrayList
        '
        For ia As Integer = 0 To Count - 1  'Suponiendo que pueden manejar 20 formularios iguales
            '
            If Tipo = T_Form.Standard Then
                If fPappl.Forms.Item(ia).TypeEx = FormID Then
                    If fPappl.Forms.Item(ia).TypeCount = TypeCount Then
                        Return True
                    End If
                End If
            Else
                If fPappl.Forms.Item(ia).UniqueID = FormID Then
                    Return True
                End If
            End If
        Next
        Return False
        '
    End Function

    ''' <summary>
    ''' Devuleve un número que indica la cantidad de veces que se encuentra abierto un formulario.
    ''' </summary>
    ''' <param name="FormId">Identificador del Formulario</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFormCount(ByVal FormId As String) As Integer
        Dim retorno As Integer = 1
        Try
            For retorno = 1 To 200
                fPappl.Forms.GetFormByTypeAndCount(FormId, retorno)
            Next
        Catch ex As Exception
        End Try
        Return retorno - 1
    End Function

    ''' <summary>
    ''' Funcion para renombrar los nombres de los campos en los formularios
    ''' </summary>
    ''' <param name="FormID">El FormTypeEx para los Estandar y el FormUID para Forms desarrollados</param>
    ''' <param name="ItemID">El nombre del Item dibujado en el formulario</param>
    ''' <param name="Descripcion">Nueva descripcion a asignar al Item</param>
    ''' <param name="ColumnID">ColumnID en caso de ser una Columna, Por default es -1</param>
    ''' <param name="IsBold">Si aplica Negritas</param>
    ''' <param name="IsItalic">Si aplica Cursiva</param>
    ''' <returns>Booleano, si es True se realizo el cambio, si es False hubo un error</returns>
    ''' <remarks></remarks>
    Public Function DynamicSystemStrings(ByVal FormID As String, ByVal ItemID As String, ByVal Descripcion As String, Optional ByVal ColumnID As String = "-1", _
    Optional ByVal IsBold As Boolean = False, Optional ByVal IsItalic As Boolean = False) As Boolean
        Dim DS As SAPbobsCOM.DynamicSystemStrings
        Dim ret As Integer = 0
        Try
            DS = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDynamicSystemStrings)
            If DS.GetByKey(FormID, ItemID, ColumnID) = False Then
                DS.FormID = FormID
                DS.ItemID = ItemID
                DS.ColumnID = ColumnID
                DS.ItemString = Descripcion
                DS.IsBold = IIf(IsBold = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                DS.IsItalics = IIf(IsItalic = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                ret = DS.Add
                If ret <> 0 Then : Return False
                Else : Return True
                End If
            Else
                DS.ItemString = Descripcion
                DS.IsBold = IIf(IsBold = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                DS.IsItalics = IIf(IsItalic = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                ret = DS.Update()
                If ret <> 0 Then : Return False
                Else : Return True
                End If
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Limpia cualquier mensaje del statusbar
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearStatusBar()
        Try
            fPappl.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Selecciona un valor de un ComboBox.
    ''' </summary>
    ''' <param name="Combo">Objeto Combobox a utilizar</param>
    ''' <param name="valor">Valor que se va a seleccionar</param>
    ''' <param name="searchKey">Tipo de valor</param>
    ''' <param name="exclusive">True para coincidencia exacta, False para coincidencia aproximada</param>
    ''' <returns>Retorna el error que se produce al seleccionar el valor. Si no hay error retorna una cadena vacía.</returns>
    ''' <remarks></remarks>
    Public Function setComboSelected(ByVal Combo As SAPbouiCOM.ComboBox, ByVal valor As String, Optional ByVal searchKey As SAPbouiCOM.BoSearchKey = BoSearchKey.psk_ByValue, Optional ByVal exclusive As Boolean = True) As String
        Dim r As String = ""
        Try
            If exclusive Then
                If searchKey = BoSearchKey.psk_Index Then
                    Combo.SelectExclusive(CInt(nzDouble(valor)), searchKey)
                Else
                    Combo.SelectExclusive(valor, searchKey)
                End If
            Else
                If searchKey = BoSearchKey.psk_Index Then
                    Combo.Select(CInt(nzDouble(valor)), searchKey)
                Else
                    Combo.Select(valor, searchKey)
                End If
            End If
        Catch ex As Exception
            r = ex.Message
        End Try
        Return r
    End Function


    ' MENU

    ''' <summary>
    ''' Añade un item tipo POP_UP al menú de usuario de B1.
    ''' </summary>
    ''' <param name="nombreCarpeta">Nombre con el que aparecerá identificada la carpeta</param>
    ''' <param name="idCarpeta">Código identificador interno del item de menú</param>
    ''' <param name="menu1">Carpeta de nivel 1 que contiene este item</param>
    ''' <param name="menu2">Carpeta de nivel 2 que contiene este item (contenida por la carpeta de nivel 1)</param>
    ''' <param name="menu3">Carpeta de nivel 3 que contiene este item (contenida por la carpeta de nivel 2)</param>
    ''' <param name="imagen">Imágen que se mostrará a la izquierda del item (solo nivel 1)</param>
    ''' <remarks></remarks>
    Public Sub addMenuFolder(ByVal nombreCarpeta As String, ByVal idCarpeta As String, Optional ByVal menu1 As String = "", Optional ByVal menu2 As String = "", Optional ByVal menu3 As String = "", Optional ByVal imagen As String = "", Optional ByVal posicion As Integer = -1)
        Try
            If imagen <> "" Then
                imagen = System.Windows.Forms.Application.StartupPath & "\" & carpetaImagenes & "\" & imagen
            End If
            If menu3 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
            ElseIf menu2 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
            ElseIf menu1 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item(menu1).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
            Else
                If fPappl.Menus.Item("43520").SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item("43520").SubMenus.Remove(fPappl.Menus.Item("43520").SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item("43520").SubMenus.Count
                fPappl.Menus.Item("43520").SubMenus.Add(idCarpeta, nombreCarpeta, BoMenuType.mt_POPUP, posicion)
                fPappl.Menus.Item("43520").SubMenus.Item(idCarpeta).Image = imagen
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Añade un item tipo STRING al menú de usuario de B1.
    ''' </summary>
    ''' <param name="idForm">Código identificador interno del item</param>
    ''' <param name="nombreForm">Nombre que se presentará al usuario</param>
    ''' <param name="menu1">Carpeta de nivel 1 que contiene este item</param>
    ''' <param name="menu2">Carpeta de nivel 2 que contiene este item (contenida por la carpeta de nivel 1)</param>
    ''' <param name="menu3">Carpeta de nivel 3 que contiene este item (contenida por la carpeta de nivel 2)</param>
    ''' <param name="menu4">Carpeta de nivel 4 que contiene este item (contenida por la carpeta de nivel 3)</param>
    ''' <remarks></remarks>
    Public Sub addMenuItem(ByVal idForm As String, ByVal nombreForm As String, ByVal menu1 As String, Optional ByVal menu2 As String = "", Optional ByVal menu3 As String = "", Optional ByVal menu4 As String = "", Optional ByVal posicion As Integer = -1)
        Try
            If menu4 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            ElseIf menu3 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            ElseIf menu2 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            Else
                If fPappl.Menus.Item(menu1).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Elimina un item de menú
    ''' </summary>
    ''' <param name="menuId">Identificador del Menú</param>
    ''' <remarks></remarks>
    Public Sub removeMenuItem(ByVal menuId As String)
        Try
            fPappl.Menus.Remove(fPappl.Menus.Item(menuId))
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Abre el formulario por defecto de una UDT
    ''' </summary>
    ''' <param name="tabla">Nombre (Código) de la Tabla de usuario a abrir</param>
    ''' <remarks></remarks>
    Public Sub abrirTablaUsuario(ByVal tabla As String)
        Try
            For i As Integer = 51201 To 51999
                If fPappl.Menus.Item(i.ToString).String.ToString.StartsWith(tabla) Then
                    fPappl.Menus.Item(i.ToString).Activate()
                    Exit Sub
                End If
            Next
            fPappl.MessageBox("No se pudo encontrar la tabla de usuario " & tabla)
        Catch ex As Exception

        End Try
    End Sub



    ' CARGAS AUTOMÁTICAS

    ''' <summary>
    ''' Carga los datos de un query en un combo.
    ''' </summary>
    ''' <param name="strFormUID">UID del formulario en el que se encuentra el combo</param>
    ''' <param name="strItemUID">UID del Combo o de la Matriz que lo contiene</param>
    ''' <param name="strQRY">String con la consulta de selección de los datos para el combo. Tomará la primera columna como VALUE y la segunda (opcional) como DESCRIPTION.</param>
    ''' <param name="incluirValorCero">Si se coloca en true, incluye un valor en blanco en el combo</param>
    ''' <remarks></remarks>
    Public Sub cargaCombo(ByVal strFormUID As String, ByVal strItemUID As String, ByVal strQRY As String, Optional ByVal incluirValorCero As Boolean = False)
        Try
            Dim Fila As Integer = 0
            Dim oComboX As SAPbouiCOM.ComboBox
            Dim strCMBdesc As String = vbNullString
            Dim xForm As SAPbouiCOM.Form = fPappl.Forms.Item(strFormUID)
            oComboX = xForm.Items.Item(strItemUID).Specific

            cargaCombo(oComboX, strQRY, incluirValorCero)

        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Carga los datos de un query en un combo que se encuentra en una Matriz.
    ''' </summary>
    ''' <param name="strFormUID">UID del formulario en el que se encuentra el combo</param>
    ''' <param name="strMatrixUID">UID de la Matriz que lo contiene</param>
    ''' <param name="colUID">UID de la columna tipo combo</param>
    ''' <param name="Fila">Fila de la matriz a actualizar</param>
    ''' <param name="strQRY">String con la consulta de selección de los datos para el combo. Tomará la primera columna como VALUE y la segunda (opcional) como DESCRIPTION.</param>
    ''' <param name="incluirValorCero">Si se coloca en true, incluye un valor en blanco</param>
    ''' <remarks></remarks>
    Public Sub cargaCombo(ByVal strFormUID As String, ByVal strMatrixUID As String, ByVal colUID As String, ByVal Fila As Integer, ByVal strQRY As String, Optional ByVal incluirValorCero As Boolean = False)
        Try
            Dim oComboX As SAPbouiCOM.ComboBox
            Dim strCMBdesc As String = vbNullString
            Dim xForm As SAPbouiCOM.Form = fPappl.Forms.Item(strFormUID)
            Dim xMatrix As SAPbouiCOM.Matrix
            Dim xColumn As SAPbouiCOM.Column
            xMatrix = xForm.Items.Item(strMatrixUID).Specific
            xColumn = xMatrix.Columns.Item(colUID)
            oComboX = xColumn.Cells.Item(Fila).Specific

            cargaCombo(oComboX, strQRY, incluirValorCero)

        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Carga los datos de un query en un combo.
    ''' </summary>
    ''' <param name="oComboX">ComboBox a llenar</param>
    ''' <param name="strQRY">String con la consulta de selección de los datos para el combo. Tomará la primera columna como VALUE y la segunda (opcional) como DESCRIPTION.</param>
    ''' <param name="incluirValorCero">Si se coloca en true, incluye un valor en blanco en el combo</param>
    ''' <remarks></remarks>
    Public Sub cargaCombo(ByVal oComboX As SAPbouiCOM.ComboBox, ByVal strQRY As String, Optional ByVal incluirValorCero As Boolean = False)
        Try
            Dim strCMBdesc As String = vbNullString

            If oComboX.ValidValues.Count > 0 Then
                For Fila As Integer = 0 To oComboX.ValidValues.Count - 1
                    oComboX.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Next Fila
            End If
            Dim xAuxLocal As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            xAuxLocal.DoQuery(strQRY)
            If incluirValorCero = True Then oComboX.ValidValues.Add("", "")
            If xAuxLocal.RecordCount > 0 Then
                xAuxLocal.MoveFirst()
                strCMBdesc = xAuxLocal.Fields.Item(0).Value
                While Not xAuxLocal.EoF
                    If xAuxLocal.Fields.Count = 1 Then
                        oComboX.ValidValues.Add(xAuxLocal.Fields.Item(0).Value, xAuxLocal.Fields.Item(0).Value)
                    Else
                        oComboX.ValidValues.Add(xAuxLocal.Fields.Item(0).Value, xAuxLocal.Fields.Item(1).Value)
                    End If
                    xAuxLocal.MoveNext()
                End While
            Else
                strCMBdesc = "Definir nuevo"
            End If
            Release(xAuxLocal)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Añade una fila en una matriz, validando que el campo de la primera columna de la última fila esté lleno.
    ''' </summary>
    ''' <param name="formulario">UID del formulario en el que se encuentra la matriz</param>
    ''' <param name="matriz">UID de la matriz</param>
    ''' <param name="celdaEsCombo">Indica si la celda a evaluar es un combo o un editText</param>
    ''' <remarks></remarks>
    Public Sub addRow(ByVal formulario As String, ByVal matriz As String, Optional ByVal celdaEsCombo As Boolean = False)
        Try
            Dim fMatrix As SAPbouiCOM.Matrix = fPappl.Forms.Item(formulario).Items.Item(matriz).Specific
            If celdaEsCombo = True Then
                Dim fCombo As SAPbouiCOM.ComboBox = fMatrix.Columns.Item(1).Cells.Item(fMatrix.RowCount).Specific
                If fMatrix.RowCount = 0 Or fCombo.Selected.Value > 0 Then
                    fMatrix.AddRow()
                    fPappl.Forms.Item(formulario).Update()
                    fPappl.Forms.Item(formulario).Refresh()
                    fCombo = fMatrix.Columns.Item(1).Cells.Item(fMatrix.RowCount).Specific
                    fCombo.Select("-1", BoSearchKey.psk_ByValue)
                End If
            Else
                Dim fEdit As SAPbouiCOM.EditText = fMatrix.Columns.Item(1).Cells.Item(fMatrix.RowCount).Specific
                If fMatrix.RowCount = 0 Or fEdit.Value > 0 Then
                    fMatrix.AddRow()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Devuelve el siguiente valor numérico para un campo. Si falla, devuelve cero.
    ''' </summary>
    ''' <param name="CampoMax">Campo del correlativo</param>
    ''' <param name="Tabla">Tabla del campo correlativo</param>
    ''' <param name="condicion">Cláusula WHERE... para la consulta</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCorrelativo(ByVal CampoMax As String, ByVal Tabla As String, Optional ByVal condicion As String = "", Optional ByVal primerCorrelativo As Integer = 1) As String
        Dim oMax As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Srt As String = primerCorrelativo.ToString
        Try
            Srt = "SELECT ISNULL(MAX(CAST(" & CampoMax & " AS numeric)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM " & Tabla
            If condicion <> "" Then
                Srt = "SELECT ISNULL(MAX(CAST(" & CampoMax & " AS numeric)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM (SELECT * FROM OWHS WHERE " & condicion & ") AS X WHERE " & condicion
            End If
            oMax.DoQuery(Srt)
            Srt = IIf(oMax.EoF = True, primerCorrelativo.ToString, oMax.Fields.Item("Numero").Value)

        Catch ex As Exception
            manejaErrores(ex, "Obteniendo Correlativo")
            Srt = "0"
        Finally
            Release(oMax)
        End Try
        Return Srt
    End Function



    ' MANEJO DE VALORES NULOS Y CONVERSIÓN

    ''' <summary>
    ''' Devuelve un string. Si el parámetro es nulo, devuelve una cadena vacía.
    ''' </summary>
    ''' <param name="unString">Valor a convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzString(ByVal unString As String, Optional ByVal formatoSQL As Boolean = False, Optional ByVal valorSiNulo As String = "") As String
        Try
            If Not IsDBNull(unString) Then
                If formatoSQL Then
                    unString = unString.Replace("'", "' + CHAR(39) + '")
                End If
                valorSiNulo = unString
            End If
        Catch ex As Exception
        End Try
        Return valorSiNulo
    End Function

    ''' <summary>
    ''' Devuelve un valor double, validando si es nulo o infinito, y si lo es devuelve un valor establecido.
    ''' </summary>
    ''' <param name="a">Valor a evaluar</param>
    ''' <param name="valorSiEsNulo">Valor que será devuelto si es nulo o infinito</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzDouble(ByVal a As Double, Optional ByVal valorSiEsNulo As Double = 0) As Double
        Try
            If Not IsDBNull(a) And Not Double.IsInfinity(a) Then
                Return a
            Else
                Return valorSiEsNulo
            End If
        Catch ex As Exception
            Return valorSiEsNulo
        End Try
    End Function

    ''' <summary>
    ''' Devuelve una fecha en formato de Business One
    ''' </summary>
    ''' <param name="fecha">Fecha a convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzDate(ByVal fecha As String, Optional ByVal valorSiNulo As String = "") As String
        Dim retorno As String = valorSiNulo
        Try
            If Not IsDBNull(fecha) Then

                Dim f As Date = CDate(fecha)
                retorno += f.Year.ToString
                If f.Month < 10 Then retorno += "0"
                retorno += f.Month.ToString
                If f.Day < 10 Then retorno += "0"
                retorno += f.Day.ToString

            End If
        Catch ex As Exception
            retorno = valorSiNulo
        End Try
        Return retorno
    End Function

    ''' <summary>
    ''' Redondea un número bajo los criterios especificados.
    ''' </summary>
    ''' <param name="valor">Valor a redondear</param>
    ''' <param name="posicionDecimal">Cantidad de decimales para redondeo. Ejemplo: 12,34... (-1) = 10 ... (0) = 12 ... (1) = 12,3</param>
    ''' <param name="siempreHaciaArriba">Indica si se desea que siempre redondee al siguiente número si existen decimales. Ejemplo: 12,34 ... (true) = 13 ... (false) = 12</param>
    ''' <param name="aCeroOCinco">Indica si se redondea en base 5 en lugar de base 10. Ejemplo: 12,34 ... (true) = 12,5 ... (false) = 12,0</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzRedondear(ByVal valor As Double, Optional ByVal posicionDecimal As Integer = 0, Optional ByVal siempreHaciaArriba As Boolean = False, Optional ByVal aCeroOCinco As Boolean = False) As Double
        Dim retorno As Double = valor
        Dim sensibilidad As Double = 0.01
        Dim Rnumero As Double = 0
        Dim valorASumar As Double = 0
        Try
            If aCeroOCinco = True Then posicionDecimal += 1

            valor = valor / Math.Pow(10, posicionDecimal)
            If siempreHaciaArriba = True And aCeroOCinco = False Then valor += 0.5 - sensibilidad

            If aCeroOCinco = True And Not valor = Math.Round(valor) Then
                Dim Rvst As String = valor.ToString
                If Rvst.Contains(",") Then
                    Rvst = Rvst.Substring(Rvst.IndexOf(",") + 1, Rvst.Length - Rvst.IndexOf(",") - 1)
                    Rvst = CDbl(Rvst) / Math.Pow(10, Rvst.Length - 1)
                ElseIf Rvst.Contains(".") Then
                    Rvst = Rvst.Substring(Rvst.IndexOf(".") + 1, Rvst.Length - Rvst.IndexOf(".") - 1)
                    Rvst = CDbl(Rvst) / Math.Pow(10, Rvst.Length - 1)
                Else
                    Rvst = Rvst.Substring(Rvst.Length - 1, 1)
                End If
                Rnumero = CDbl(Rvst)

                If Rnumero = 0 Then
                    valorASumar = -10
                ElseIf Rnumero > 0 And Rnumero < 5 Then
                    valorASumar = 5
                ElseIf Rnumero = 5 Then
                    valorASumar = 5
                ElseIf Rnumero > 5 Then
                    valorASumar = 0
                End If

            End If

            retorno = Math.Round(valor, 0)
            If aCeroOCinco = True Then
                retorno = (retorno * 10) + valorASumar
                posicionDecimal -= 1
            End If
            retorno = retorno * Math.Pow(10, posicionDecimal)

        Catch ex As Exception
        End Try
        Return retorno

    End Function

    ''' <summary>
    ''' Retorna una fecha en formato para consultas SQL
    ''' </summary>
    ''' <param name="fecha">Variable tipo fecha a convertir</param>
    ''' <returns>Devuelve un string que se puede concatenar en un query</returns>
    ''' <remarks>Evalua el dateformat en el servidor SQL y convierte la fecha acorde a este formato</remarks>
    Public Function getDateSQL(ByVal fecha As String) As Date
        Dim ret As Date = Nothing
        Try

            Dim formato As String = Me.getRSvalue("select dateformat from master.dbo.syslanguages where name=(SELECT @@language)")
            Select Case formato
                Case "dmy"
                    ret = getDateVar(fecha.Substring(0, 2), fecha.Substring(3, 2), fecha.Substring(6, 4))
                Case "mdy"
                    ret = getDateVar(fecha.Substring(3, 2), fecha.Substring(0, 2), fecha.Substring(6, 4))
                Case "ymd"
                    ret = getDateVar(fecha.Substring(8, 2), fecha.Substring(5, 2), fecha.Substring(0, 4))
            End Select

        Catch ex As Exception
        End Try
        Return ret
    End Function
    ''' <summary>
    ''' Retorna una fecha en formato para consultas SQL
    ''' </summary>
    ''' <param name="fecha">Variable tipo fecha a convertir</param>
    ''' <returns>Devuelve un string que se puede concatenar en un query</returns>
    ''' <remarks>Evalua el dateformat en el servidor SQL y convierte la fecha acorde a este formato</remarks>
    Public Function getDateSQL(ByVal fecha As Date) As String
        Dim ret As String = ""
        Try

            Dim formato As String = Me.getRSvalue("select dateformat from master.dbo.syslanguages where name=(SELECT @@language)")
            Select Case formato
                Case "dmy"
                    ret = Format(CDate(fecha), "dd/MM/yyyy")
                Case "mdy"
                    ret = Format(CDate(fecha), "MM/dd/yyyy")
                Case "ymd"
                    ret = Format(CDate(fecha), "yyyy/dd/MM")
            End Select

        Catch ex As Exception
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' Retorna la parte numérica de una cadena string
    ''' </summary>
    ''' <param name="cadena">Cadena string a evaluar</param>
    ''' <param name="separadorDecimal">Caracter a tomar en cuenta como separador decimal</param>
    ''' <returns>Retorna cadena string con la parte numérica y el separador de decimales asignado</returns>
    ''' <remarks></remarks>
    Public Function getParteNumerica(ByVal cadena As String, Optional ByVal separadorDecimal As String = ",") As String
        Dim Codevar As String = ""
        Try
            For ii As Integer = 1 To Len(cadena)
                If IsNumeric(Mid(cadena, ii, 1)) Or Mid(cadena, ii, 1) = separadorDecimal Then
                    Codevar &= Mid(cadena, ii, 1)
                End If
            Next
        Catch ex As Exception
            manejaErrores(ex, "Obteniendo parte numérica")
        End Try
        Return Codevar
    End Function


    ' OTRAS

    ''' <summary>
    ''' Si el idioma de Business One es cualquier variante de Español, devuelve true. En caso contrario, false.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isInSpanish() As Boolean
        Dim miBool As Boolean = False
        Try
            If fPappl.Language = BoLanguages.ln_Spanish Or fPappl.Language = BoLanguages.ln_Spanish_Ar Or fPappl.Language = BoLanguages.ln_Spanish_La Or fPappl.Language = BoLanguages.ln_Spanish_Pa Then miBool = True
        Catch ex As Exception
            manejaErrores(ex, "Determinando idioma de SAP")
        End Try
        Return miBool
    End Function

    ''' <summary>
    ''' Agrega una línea al archivo txt del log.
    ''' </summary>
    ''' <param name="Contenido">Contenido de la línea de texto</param>
    ''' <param name="FileName">Nombre del archivo an el que se registra el log (sin extensión .txt)</param>
    ''' <param name="Ruta">Ruta en la que se guardará el archivo (Ejemplo: C:\Logs)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function addLogTxt(ByVal Contenido As String, Optional ByVal FileName As String = "", Optional ByVal Ruta As String = "") As Boolean
        Try
            If Ruta = "" Then
                Ruta = System.IO.Directory.GetCurrentDirectory & "\Logs"
            End If
            If FileName = "" Then
                FileName = logErroresArchivo
            End If
            If System.IO.Directory.Exists(Ruta) = False Then
                System.IO.Directory.CreateDirectory(Ruta)
            End If
            If Not Ruta.EndsWith("\") Or Not Ruta.EndsWith("/") Then Ruta &= "\"
            System.IO.File.AppendAllText(Ruta & FileName & ".txt", Date.Now & "; " & Contenido & Chr(13))
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Muestra un MessageBox con las opciones Sí / No (en el idioma correcto)
    ''' </summary>
    ''' <param name="mensaje">Mensaje a mostrar</param>
    ''' <param name="defaultButton">Botón por defecto</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function msgboxYN(ByVal mensaje As String, Optional ByVal defaultButton As Integer = 1) As Integer
        Try
            If isInSpanish() Then
                Return fPappl.MessageBox(mensaje, defaultButton, "Sí", "No")
            Else
                Return fPappl.MessageBox(mensaje, defaultButton, "Yes", "No")
            End If
        Catch ex As Exception
            manejaErrores(ex, "Cargando Ventana Sí/No")
            Return Nothing
        End Try
    End Function


    ' RECURSOS

    ''' <summary>
    ''' Copia un recurso a una carpeta determinada
    ''' </summary>
    ''' <param name="recurso">Recurso del proyecto (My.Resources...)</param>
    ''' <param name="nombreArchivo">Nombre con el que se copiará el archivo (nombre y extensión)</param>
    ''' <param name="ruta">Ruta física en la que se copiará el archivo</param>
    ''' <remarks>Se implementa importando recursos al proyecto y cargandolos así: cargaRecurso(my.Resources.QUERY, "QUERY.sql")</remarks>
    Public Sub cargaRecurso(ByVal recurso As Object, ByVal nombreArchivo As String, Optional ByVal ruta As String = "", Optional ByVal borrarSiExiste As Boolean = False)
        Try

            If ruta = "" Then
                ruta = carpetaFormularios
            End If

            If Not IO.Directory.Exists(ruta) Then IO.Directory.CreateDirectory(ruta)

            If Not ruta.EndsWith("\") Then ruta &= "\"
            ruta &= nombreArchivo

            If IO.File.Exists(ruta) Then
                If borrarSiExiste Then
                    IO.File.Delete(ruta)
                Else
                    Exit Sub
                End If
            End If

            Try
                ' archivos XML
                Dim xmlDoc As New Xml.XmlDocument
                xmlDoc.LoadXml(recurso)
                xmlDoc.Save(ruta)
            Catch
                Try
                    ' imagenes
                    recurso.toBitmap.Save(ruta)
                Catch

                    Try
                        ' otros
                        recurso.Save(ruta)

                    Catch

                        ' archivo de texto
                        IO.File.WriteAllText(ruta, recurso.ToString)

                    End Try
                End Try
            End Try

        Catch ex As Exception
            manejaErrores(ex, "Cargando Recurso " & nombreArchivo)

        End Try
    End Sub

    ''' <summary>
    ''' Levanta un formulario desde un Recurso del Proyecto.
    ''' </summary>
    ''' <param name="Recurso">Recurso del formulario (My.Resources...)</param>
    ''' <param name="medianteLBA">Indica si el método por el que se levantará el formulario es LoadBatchActions (true) o creationPackage (false)</param>
    ''' <returns>Retorna una cadena vacía si tuvo éxito y un mensaje de error en caso de fallo</returns>
    ''' <remarks>Permite cargar formularios desde recursos, así: cargaForm(My.Resources.XMLdelFORM)</remarks>
    Public Function cargaForm(ByVal Recurso As Object, Optional ByVal medianteLBA As Boolean = False) As String
        Dim rr As String = ""
        Try
            If medianteLBA Then
                fPappl.LoadBatchActions(Recurso.ToString)
            Else
                Dim creationPackage As SAPbouiCOM.FormCreationParams
                creationPackage = fPappl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                creationPackage.XmlData = Recurso.ToString
                fPappl.Forms.AddEx(creationPackage)
            End If
        Catch ex As Exception
            manejaErrores(ex, "Cargando Formulario")
            rr = ex.Message
        End Try
        Return rr
    End Function



    'EXCEL

    ''' <summary>
    ''' Retorna una conexión ADO a un archivo Excel.
    ''' </summary>
    ''' <param name="NombreArch">Nombre del archivo incluyendo extensión (Ej: "Prueba.xls")</param>
    ''' <param name="Ruta">Ubicación del archivo (Ej: "C:\Excel\"). Si no se especifica toma la carpeta de archivos excel de las parametrizaciones generales.</param>
    ''' <returns>Devuelve una conexión abierta a Excel</returns>
    ''' <remarks>La conexión debe cerrarse manualmente</remarks>
    Public Function getExcelConnection(ByVal NombreArch As String, Optional ByVal Ruta As String = "") As ADODB.Connection
        Try
            Dim ConexionExcel As New ADODB.Connection
            ConexionExcel.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            ConexionExcel.CommandTimeout = 1200

            Dim strSQL As String = ""
            If Ruta = "" Then Ruta = fCompany.ExcelDocsPath
            If Ruta = vbNullString Then
                Ruta = "C:\"
            End If

            Dim v2007 As Boolean = NombreArch.ToUpper.EndsWith(".XLSX")

            If v2007 Then
                strSQL = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Ruta & NombreArch & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""

            Else
                strSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Ruta & NombreArch & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""

            End If

            ConexionExcel.ConnectionString = strSQL
            ConexionExcel.Open()

            Return ConexionExcel

        Catch ex As Exception
            manejaErrores(ex, "Conectando a Archivo Excel")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuelve un Recordset sobre un archivo Excel.
    ''' </summary>
    ''' <param name="Query">Consulta a ejecutar sobre el archivo Excel</param>
    ''' <param name="NombreArch">Nombre del archivo incluyendo extensión (Ej: "Prueba.xls")</param>
    ''' <param name="Ruta">Ubicación del archivo (Ej: "C:\Excel\"). Si no se especifica toma la carpeta de archivos excel de las parametrizaciones generales.</param>
    ''' <returns>Objeto ADODB.Recordset con el resultado del query especificado</returns>
    ''' <remarks>La conexión debe cerrarse manualmente</remarks>
    Public Function getExcelRecordset(ByVal Query As String, ByVal NombreArch As String, Optional ByVal Ruta As String = "") As ADODB.Recordset
        Try
            Dim ConexionExcel As ADODB.Connection = getExcelConnection(NombreArch, Ruta)
            Dim rsLocal As New ADODB.Recordset
            rsLocal.Open(Query, ConexionExcel, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
            Return rsLocal
        Catch ex As Exception
            manejaErrores(ex, "Obteniedo Datos de Excel")
            Return Nothing
        End Try
    End Function


    ' ADO / SQL ( ESTA SECCIÓN NO HA SIDO PROBABA )

    ''' <summary>
    ''' Devuelve una cadena de conexión ADO
    ''' </summary>
    ''' <param name="tipoDeConexion">Seleccione el tipo de conexión que desea realizar</param>
    ''' <param name="servidorOruta">Nombre del servidor o Ruta del Archivo</param>
    ''' <param name="baseDeDatosOarchivo">Nombre de la Base de Daots o Archivo (con extensión)</param>
    ''' <param name="usuario">Nombre de usuario del origen de datos</param>
    ''' <param name="password">Contraseña del origen de datos</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getADOConnectionString(ByVal tipoDeConexion As enumADOconnections, ByVal servidorOruta As String, ByVal baseDeDatosOarchivo As String, ByVal usuario As String, ByVal password As String) As String
        Dim strCadena As String = ""
        Try
            If tipoDeConexion = enumADOconnections.SQL_Server_2000 Or tipoDeConexion = enumADOconnections.SQL_Server_2005 Then
                strCadena = "Provider=SQLOLEDB.1;"
                strCadena = strCadena & "Password=" & password & ";"
                strCadena = strCadena & "Persist Security Info=True;"
                strCadena = strCadena & "User ID=" & usuario & ";"
                strCadena = strCadena & "Initial Catalog=" & baseDeDatosOarchivo & ";"
                strCadena = strCadena & "Data Source=" & servidorOruta

            ElseIf tipoDeConexion = enumADOconnections.Excel_2007 Then
                strCadena = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & servidorOruta & baseDeDatosOarchivo & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""

            ElseIf tipoDeConexion = enumADOconnections.Excel_2003 Then
                strCadena = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & servidorOruta & baseDeDatosOarchivo & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""

            End If

        Catch ex As Exception
            manejaErrores(ex, "armando cadena de conexión ADO")
        End Try
        Return strCadena
    End Function

    ''' <summary>
    ''' Devuelve una cadena de conexión ADO en base a los parámetros de conexión de B1
    ''' </summary>
    ''' <param name="passwordSQL">Contraseña del origen de datos</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getB1ConnectionString(Optional ByVal usuarioSQL As String = Nothing, Optional ByVal passwordSQL As String = Nothing) As String
        Dim strCadena As String = ""
        Try
            Dim tipoBD As enumADOconnections = enumADOconnections.SQL_Server_2005
            If fCompany.DbServerType = BoDataServerTypes.dst_MSSQL Then tipoBD = enumADOconnections.SQL_Server_2000
            If passwordSQL Is Nothing Then passwordSQL = fCompany.DbPassword
            If passwordSQL.StartsWith("*") And passwordSQL.EndsWith("*") Then passwordSQL = "B1Admin"
            If usuarioSQL Is Nothing Then usuarioSQL = fCompany.DbUserName
            strCadena = getADOConnectionString(tipoBD, fCompany.Server, fCompany.CompanyDB, usuarioSQL, passwordSQL)
        Catch ex As Exception
            manejaErrores(ex, "Armando cadena de conexión ADO B1")
        End Try
        Return strCadena
    End Function

    ''' <summary>
    ''' Establece una conexión ADO (SQL, etc)
    ''' </summary>
    ''' <param name="connectionString">Cadena de conexión ADO a ejecutar</param>
    ''' <param name="commandTimeout">Tiempo máximo de timeout para la conexión</param>
    ''' <returns>La conexión se retorna abierta</returns>
    ''' <remarks></remarks>
    Public Function getADOconnection(ByVal connectionString As String, Optional ByVal commandTimeout As Integer = 1200) As ADODB.Connection
        Try
            Dim dbConexionADO As ADODB.Connection = New ADODB.Connection
            dbConexionADO.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            dbConexionADO.CommandTimeout = commandTimeout
            dbConexionADO.ConnectionString = connectionString
            dbConexionADO.Open()
            Return dbConexionADO
        Catch ex As Exception
            manejaErrores(ex, "abriendo conexión ADO")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Ejecuta un comando ADO
    ''' </summary>
    ''' <param name="conexion">Conexión ADO abierta</param>
    ''' <param name="query">Sentencia SQL a ejecutar</param>
    ''' <param name="cerrarConexion">Indica si se debe cerrar la conexión una vez ejecutado el comando</param>
    ''' <returns>Retorna un valor entero con la cantidad de registros afectados</returns>
    ''' <remarks>Si falla retorna cero</remarks>
    Public Function runADOcommand(ByVal conexion As ADODB.Connection, ByVal query As String, Optional ByVal cerrarConexion As Boolean = False) As Integer
        Dim registrosAfectados As Integer = 0
        Try
            conexion.Execute(query, registrosAfectados)
        Catch ex As Exception
            manejaErrores(ex, "Ejecutando comando ADO")
        Finally
            Try
                If cerrarConexion Then conexion.Close()
            Catch
            End Try
        End Try
        Return registrosAfectados
    End Function

    ''' <summary>
    ''' Crea una Vista por medio de un comando ADO
    ''' </summary>
    ''' <param name="query">Sentencia SQL a ejecutar</param>
    ''' <param name="passwordSQL">Clave del usuario SQL. Si no se especifica, tomará la clave por defecto.</param>
    ''' <returns>Retorna un valor entero con la cantidad de registros afectados</returns>
    ''' <remarks>Si falla retorna cero</remarks>
    Public Function runADOcommandB1(ByVal query As String, Optional ByVal passwordSQL As String = Nothing) As Integer
        Dim registrosAfectados As Integer = 0
        Try
            Dim cnx As ADODB.Connection = getADOconnection(getB1ConnectionString(, passwordSQL))
            registrosAfectados = runADOcommand(cnx, query)
            cnx.Close()
        Catch ex As Exception
            manejaErrores(ex, "Ejecutando comando ADO B1")
        Finally
        End Try
        Return registrosAfectados
    End Function

    ''' <summary>
    ''' Ejecuta un query y retorna un recordset ADO
    ''' </summary>
    ''' <param name="conexion">Conexión ADO abierta</param>
    ''' <param name="query">Sentencia SQL a ejecutar</param>
    ''' <returns>Retorna un recordset ADO</returns>
    ''' <remarks>Si falla retorna Nothing</remarks>
    Public Function getADOrecordset(ByRef conexion As ADODB.Connection, ByVal query As String, Optional ByVal cerrarConexion As Boolean = False) As ADODB.Recordset
        Dim rs As ADODB.Recordset = Nothing
        Try
            rs = conexion.Execute(query)
        Catch ex As Exception
            manejaErrores(ex, "Obteniendo RecordSet ADO")
        Finally
            Try
                If cerrarConexion Then conexion.Close()
            Catch ex As Exception
            End Try
        End Try
        Return rs
    End Function



    ' INTERNAS

    ''' <summary>
    ''' Maneja los errores de acuerdo a lo configurado en la librería
    ''' </summary>
    ''' <param name="exx">Excepcion</param>
    ''' <param name="proceso">Nombre del proceso que desató el error</param>
    ''' <remarks></remarks>
    Private Sub manejaErrores(ByVal exx As Exception, Optional ByVal proceso As String = "")
        Try
            manejaErrores(exx.Message, proceso)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub manejaErrores(ByVal exx As String, Optional ByVal proceso As String = "")
        Try
            If proceso <> "" Then proceso = "Error " & proceso & ": "
            If mostrarMensajesError Then
                If mostrarMensajesTipo = enumTipoMensaje.MessageBox Then fPappl.MessageBox("CMP: " & proceso & exx)
                If mostrarMensajesTipo = enumTipoMensaje.StatusBarError Then fPappl.StatusBar.SetText("CMP: " & proceso & exx, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                If mostrarMensajesTipo = enumTipoMensaje.StatusBarWarning Then fPappl.StatusBar.SetText("CMP: " & proceso & exx, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If
            If mantenerLogErrores Then addLogTxt(proceso & exx)
        Catch ex As Exception
        End Try
    End Sub



End Class
