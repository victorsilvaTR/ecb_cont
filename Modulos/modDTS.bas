Attribute VB_Name = "modDTS"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: c:\Importacion XLS.bas
'Package Name: Importacion XLS
'Package Description: Descripción del paquete DTS
'Generated Date: 26/06/2008
'Generated Time: 06:21:25 p.m.
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2

Dim sSERVIDOR As String
Dim sBaseDatos As String
Dim sUSUARIO As String
Dim sPW As String


Public Function CrearTablas(sArchivoXLS As String, ByRef oBarra As ProgressBar, ByRef oMensaje As Label, ByRef cParametros As String) As Boolean
    '========================================
    'sArchivoXLS = "C:\ESTRUC.xls"
    sSERVIDOR = gsServidor
    sBaseDatos = gsBD
    sUSUARIO = gsBDUS
    sPW = gsBDPW
    '========================================
    Dim cEntidad As String
    Dim cTC As String
    Dim cApertura As String
    Dim cCompras As String
    Dim cVentas As String
    Dim cCajaIng As String
    Dim cCajaEgr As String
    Dim cPlanilla As String

    oBarra.Max = 28

    cEntidad = Mid(cParametros, 1, 1)
    cTC = Mid(cParametros, 2, 1)
    cApertura = Mid(cParametros, 3, 1)
    cCompras = Mid(cParametros, 4, 1)
    cVentas = Mid(cParametros, 5, 1)
    cCajaIng = Mid(cParametros, 6, 1)
    cCajaEgr = Mid(cParametros, 7, 1)
    cPlanilla = Mid(cParametros, 8, 1)

    Call IniciaObjetoDTS(sArchivoXLS, cParametros)

    If cApertura = "1" Then
        '------------- call Task_Sub1 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea)
        Call MsgCreaTabla("ZIMP_APE_CAB", oMensaje)
        Call Task_Sub1(goPackage)

        oBarra.Value = 1
        '------------- call Task_Sub2 for task Copy Data from APE_CAB$ to [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea (Copy Data from APE_CAB$ to [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea)
        Call MsgCopiaTabla("ZIMP_APE_CAB", oMensaje)
        Call Task_Sub2(goPackage)
        oBarra.Value = 2

        '------------- call Task_Sub3 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea)
        Call MsgCreaTabla("ZIMP_APE_DET", oMensaje)
        Call Task_Sub3(goPackage)

        oBarra.Value = 3
        '------------- call Task_Sub4 for task Copy Data from APE_DET$ to [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea (Copy Data from APE_DET$ to [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea)
        Call MsgCopiaTabla("ZIMP_APE_DET", oMensaje)
        Call Task_Sub4(goPackage)

        oBarra.Value = 4
    End If

    If cCajaEgr = "1" Then
        '------------- call Task_Sub5 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea)
        Call MsgCreaTabla("ZIMP_CAJAEGR_CAB", oMensaje)
        Call Task_Sub5(goPackage)

        oBarra.Value = 5
        '------------- call Task_Sub6 for task Copy Data from CAJAEGR_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea (Copy Data from CAJAEGR_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea)
        Call MsgCopiaTabla("ZIMP_CAJAEGR_CAB", oMensaje)
        Call Task_Sub6(goPackage)

        oBarra.Value = 6
        '------------- call Task_Sub7 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea)
        Call MsgCreaTabla("ZIMP_CAJAEGR_DET", oMensaje)
        Call Task_Sub7(goPackage)

        oBarra.Value = 7
        '------------- call Task_Sub8 for task Copy Data from CAJAEGR_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea (Copy Data from CAJAEGR_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea)
        Call MsgCopiaTabla("ZIMP_CAJAEGR_DET", oMensaje)
        Call Task_Sub8(goPackage)

        oBarra.Value = 8
    End If

    If cCajaIng = "1" Then
        '------------- call Task_Sub9 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea)
        Call MsgCreaTabla("ZIMP_CAJAING_CAB", oMensaje)
        Call Task_Sub9(goPackage)

        oBarra.Value = 9
        '------------- call Task_Sub10 for task Copy Data from CAJAING_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea (Copy Data from CAJAING_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea)
        Call MsgCopiaTabla("ZIMP_CAJAING_CAB", oMensaje)
        Call Task_Sub10(goPackage)

        oBarra.Value = 10
        '------------- call Task_Sub11 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea)
        Call MsgCreaTabla("ZIMP_CAJAING_DET", oMensaje)
        Call Task_Sub11(goPackage)

        oBarra.Value = 11
        '------------- call Task_Sub12 for task Copy Data from CAJAING_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea (Copy Data from CAJAING_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea)
        Call MsgCopiaTabla("ZIMP_CAJAING_DET", oMensaje)
        Call Task_Sub12(goPackage)

        oBarra.Value = 12
    End If

    If cCompras = "1" Then
        '------------- call Task_Sub13 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea)
        Call MsgCreaTabla("ZIMP_COMPRAS_CAB", oMensaje)
        Call Task_Sub13(goPackage)

        oBarra.Value = 13
        '------------- call Task_Sub14 for task Copy Data from COMPRAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea (Copy Data from COMPRAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea)
        Call MsgCopiaTabla("ZIMP_COMPRAS_CAB", oMensaje)
        Call Task_Sub14(goPackage)

        oBarra.Value = 14
        '------------- call Task_Sub15 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea)
        Call MsgCreaTabla("ZIMP_COMPRAS_DET", oMensaje)
        Call Task_Sub15(goPackage)

        oBarra.Value = 15
        '------------- call Task_Sub16 for task Copy Data from COMPRAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea (Copy Data from COMPRAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea)
        Call MsgCopiaTabla("ZIMP_COMPRAS_DET", oMensaje)
        Call Task_Sub16(goPackage)

        oBarra.Value = 16

    End If

    If cEntidad = "1" Then
        '------------- call Task_Sub17 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea)
        Call MsgCreaTabla("ZIMP_ENTIDAD", oMensaje)
        Call Task_Sub17(goPackage)

        oBarra.Value = 17
        '------------- call Task_Sub18 for task Copy Data from ENTIDAD$ to [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea (Copy Data from ENTIDAD$ to [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea)
        Call MsgCopiaTabla("ZIMP_ENTIDAD", oMensaje)
        Call Task_Sub18(goPackage)

        oBarra.Value = 18
    End If

    If cTC = "1" Then
        '------------- call Task_Sub19 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea)
        Call MsgCreaTabla("ZIMP_TIPOCAMBIO", oMensaje)
        Call Task_Sub19(goPackage)

        oBarra.Value = 19
        '------------- call Task_Sub20 for task Copy Data from TIPOCAMBIO$ to [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea (Copy Data from TIPOCAMBIO$ to [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea)
        Call MsgCopiaTabla("ZIMP_TIPOCAMBIO", oMensaje)
        Call Task_Sub20(goPackage)

        oBarra.Value = 20
    End If

    If cVentas = "1" Then
        '------------- call Task_Sub21 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea)
        Call MsgCreaTabla("ZIMP_VENTAS_CAB", oMensaje)
        Call Task_Sub21(goPackage)

        oBarra.Value = 21
        '------------- call Task_Sub22 for task Copy Data from VENTAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea (Copy Data from VENTAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea)
        Call MsgCopiaTabla("ZIMP_VENTAS_CAB", oMensaje)
        Call Task_Sub22(goPackage)

        oBarra.Value = 22
        '------------- call Task_Sub23 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea)
        Call MsgCreaTabla("ZIMP_VENTAS_DET", oMensaje)
        Call Task_Sub23(goPackage)

        oBarra.Value = 23
        '------------- call Task_Sub24 for task Copy Data from VENTAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea (Copy Data from VENTAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea)
        Call MsgCopiaTabla("ZIMP_VENTAS_DET", oMensaje)
        Call Task_Sub24(goPackage)

    End If

    If cPlanilla = "1" Then
        '------------- call Task_Sub25 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_PLAN_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_PLAN_CAB] Tarea)
        Call MsgCreaTabla("ZIMP_PLAN_CAB", oMensaje)
        Call Task_Sub25(goPackage)

        oBarra.Value = 24
        '------------- call Task_Sub26 for task Copy Data from PLAN_CAB$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_CAB] Tarea (Copy Data from PLAN_CAB$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_CAB] Tarea)
        Call MsgCopiaTabla("ZIMP_PLAN_CAB", oMensaje)
        Call Task_Sub26(goPackage)

        oBarra.Value = 25
        '------------- call Task_Sub27 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea)
        Call MsgCreaTabla("ZIMP_PLAN_DET", oMensaje)
        Call Task_Sub27(goPackage)

        oBarra.Value = 26
        '------------- call Task_Sub28 for task Copy Data from PLAN_DET$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea (Copy Data from PLAN_DET$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea)
        Call MsgCopiaTabla("ZIMP_PLAN_DET", oMensaje)
        Call Task_Sub28(goPackage)

    End If

    oBarra.Value = 27

    '---------------------------------------------------------------------------
    ' Save or execute package
    '---------------------------------------------------------------------------
    On Error GoTo sErrorPack
    goPackage.Execute
sErrorPack:

    If tracePackageError(goPackage) = True Then
        'SE ENCONTRARON ERRORES AL CREAR LAS TABLAS
        CrearTablas = False
    Else
        'NO SE ENCONTRARON ERRORES AL CREAR LAS TABLAS
        CrearTablas = True
    End If
    '------------------------------------------
    goPackage.UnInitialize
    'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
    Set goPackage = Nothing

    Set goPackageOld = Nothing

    Debug.Print ("Proceso terminado")
    oMensaje.Caption = ""
    oBarra.Value = 0



End Function

Private Sub MsgCreaTabla(cTabla As String, ByRef oMensaje As Label)
    oMensaje.Caption = "CREANDO : " & cTabla
    oMensaje.Refresh
    DoEvents
End Sub

Private Sub MsgCopiaTabla(cTabla As String, ByRef oMensaje As Label)
    oMensaje.Caption = "COPIANDO : " & cTabla
    oMensaje.Refresh
    DoEvents
End Sub

'-----------------------------------------------------------------------------
' error reporting using step.GetExecutionErrorInfo after execution
'-----------------------------------------------------------------------------
Public Function tracePackageError(oPackage As DTS.Package) As Boolean
Dim ErrorCode As Long
Dim ErrorSource As String
Dim ErrorDescription As String
Dim ErrorHelpFile As String
Dim ErrorHelpContext As Long
Dim ErrorIDofInterfaceWithError As String
Dim i As Integer
        tracePackageError = False

        Dim cMensaje As String

        If ErrorCode = -2147467259 Then
            cMensaje = "Intente nuevamente ..."
        Else
            cMensaje = ""
        End If

        For i = 1 To oPackage.Steps.Count
                If oPackage.Steps(i).ExecutionResult = DTSStepExecResult_Failure Then
                        oPackage.Steps(i).GetExecutionErrorInfo ErrorCode, ErrorSource, ErrorDescription, _
                                        ErrorHelpFile, ErrorHelpContext, ErrorIDofInterfaceWithError
                        Mensajes oPackage.Steps(i).Name & ", Fallo: " & vbCrLf & ErrorSource & vbCrLf & ErrorDescription & Salto(2) & UCase(cMensaje)

                        tracePackageError = True
                End If
        Next i

End Function

'------------- define Task_Sub1 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
        oCustomTask1.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"

        oCustomTask1.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "drop table [dbo].[ZIMP_APE_CAB]" & vbCrLf

        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from APE_CAB$ to [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea (Copy Data from APE_CAB$ to [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
        oCustomTask2.Description = "Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Ase_dFecha`,`Ase_cGlosa`,`Ase_cTipoMoneda` from `APE_CAB$`"
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB]"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0

Call oCustomTask2_Trans_Sub1(oCustomTask2)


goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub3 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea)
Public Sub Task_Sub3(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask3 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
Set oCustomTask3 = oTask.CustomTask

        oCustomTask3.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
        oCustomTask3.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"

        oCustomTask3.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "drop table [dbo].[ZIMP_APE_DET]" & vbCrLf

        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] (" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nItem] INT NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_dFecDoc] datetime null , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_dFecVen] datetime NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cTipoMoneda] char (3) NULL" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & ")"
        oCustomTask3.ConnectionID = 4
        oCustomTask3.CommandTimeout = 0
        oCustomTask3.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask3 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub4 for task Copy Data from APE_DET$ to [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea (Copy Data from APE_DET$ to [SAFC_ECB].[dbo].[ZIMP_APE_DET] Tarea)
Public Sub Task_Sub4(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask4 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
Set oCustomTask4 = oTask.CustomTask

        oCustomTask4.Name = "Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
        oCustomTask4.Description = "Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
        oCustomTask4.SourceConnectionID = 3
        oCustomTask4.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Pla_cCuentaContable`,`Asd_nItem`,`Asd_cGlosa`,`Asd_nDebeSoles`,`Asd_nHaberSoles`,`Asd_nTipoCambio`,`Asd_nDebeMonExt`,`Asd_nHaberMonExt`,`Cos_cCodigo`,`Ten_cTipoEntidad`,`Ent_"
        oCustomTask4.SourceSQLStatement = oCustomTask4.SourceSQLStatement & "cCodEntidad`,`Asd_cTipoDoc`,`Asd_dFecDoc`,`Asd_cSerieDoc`,`Asd_cNumDoc`,`Asd_dFecVen`,`Asd_cProvCanc`,`Asd_cOperaTC`,`Asd_cTipoMoneda` from `APE_DET$`"
        oCustomTask4.DestinationConnectionID = 4
        oCustomTask4.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_APE_DET]"
        oCustomTask4.ProgressRowCount = 1000
        oCustomTask4.MaximumErrorCount = 0
        oCustomTask4.FetchBufferSize = 1
        oCustomTask4.UseFastLoad = True
        oCustomTask4.InsertCommitSize = 0
        oCustomTask4.ExceptionFileColumnDelimiter = "|"
        oCustomTask4.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask4.AllowIdentityInserts = False
        oCustomTask4.FirstRow = 0
        oCustomTask4.LastRow = 0
        oCustomTask4.FastLoadOptions = 2
        oCustomTask4.ExceptionFileOptions = 1
        oCustomTask4.DataPumpOptions = 0

Call oCustomTask4_Trans_Sub1(oCustomTask4)


goPackage.Tasks.Add oTask
Set oCustomTask4 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask4_Trans_Sub1(ByVal oCustomTask4 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask4.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cProvCanc", 22)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cOperaTC", 23)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoMoneda", 24)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cProvCanc", 22)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cOperaTC", 23)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoMoneda", 24)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask4.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub5 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea)
Public Sub Task_Sub5(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask5 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
Set oCustomTask5 = oTask.CustomTask

        oCustomTask5.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
        oCustomTask5.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"

        oCustomTask5.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "drop table [dbo].[ZIMP_CAJAEGR_CAB]" & vbCrLf

        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] (" & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & ")"
        oCustomTask5.ConnectionID = 2
        oCustomTask5.CommandTimeout = 0
        oCustomTask5.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask5 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub6 for task Copy Data from CAJAEGR_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea (Copy Data from CAJAEGR_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_CAB] Tarea)
Public Sub Task_Sub6(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask6 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
Set oCustomTask6 = oTask.CustomTask

        oCustomTask6.Name = "Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
        oCustomTask6.Description = "Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
        oCustomTask6.SourceConnectionID = 1
        oCustomTask6.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Ase_dFecha`,`Ase_cGlosa`,`Ase_cTipoMoneda` from `CAJAEGR_CAB$`"
        oCustomTask6.DestinationConnectionID = 2
        oCustomTask6.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB]"
        oCustomTask6.ProgressRowCount = 1000
        oCustomTask6.MaximumErrorCount = 0
        oCustomTask6.FetchBufferSize = 1
        oCustomTask6.UseFastLoad = True
        oCustomTask6.InsertCommitSize = 0
        oCustomTask6.ExceptionFileColumnDelimiter = "|"
        oCustomTask6.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask6.AllowIdentityInserts = False
        oCustomTask6.FirstRow = 0
        oCustomTask6.LastRow = 0
        oCustomTask6.FastLoadOptions = 2
        oCustomTask6.ExceptionFileOptions = 1
        oCustomTask6.DataPumpOptions = 0

Call oCustomTask6_Trans_Sub1(oCustomTask6)


goPackage.Tasks.Add oTask
Set oCustomTask6 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask6_Trans_Sub1(ByVal oCustomTask6 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask6.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask6.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub7 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea)
Public Sub Task_Sub7(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask7 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
Set oCustomTask7 = oTask.CustomTask

        oCustomTask7.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
        oCustomTask7.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"

        oCustomTask7.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "drop table [dbo].[ZIMP_CAJAEGR_DET]" & vbCrLf


        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] (" & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_nItem] INT NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_dFecDoc] datetime NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_dFecVen] datetime NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_dFecDocRef] datetime NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cNumDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cRetencion] nvarchar (255) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cTipoMoneda] char (3) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Tra_cCodigo] nvarchar (255) NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[Asd_cFormaPago] nvarchar (255) NULL" & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & ")"
        oCustomTask7.ConnectionID = 4
        oCustomTask7.CommandTimeout = 0
        oCustomTask7.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask7 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub8 for task Copy Data from CAJAEGR_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea (Copy Data from CAJAEGR_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAEGR_DET] Tarea)
Public Sub Task_Sub8(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask8 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
Set oCustomTask8 = oTask.CustomTask

        oCustomTask8.Name = "Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
        oCustomTask8.Description = "Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
        oCustomTask8.SourceConnectionID = 3
        oCustomTask8.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Pla_cCuentaContable`,`Asd_nItem`,`Asd_cGlosa`,`Asd_nDebeSoles`,`Asd_nHaberSoles`,`Asd_nTipoCambio`,`Asd_nDebeMonExt`,`Asd_nHaberMonExt`,`Cos_cCodigo`,`Ten_cTipoEntidad`,`Ent_"
        oCustomTask8.SourceSQLStatement = oCustomTask8.SourceSQLStatement & "cCodEntidad`,`Asd_cTipoDoc`,`Asd_dFecDoc`,`Asd_cSerieDoc`,`Asd_cNumDoc`,`Asd_dFecVen`,`Asd_cTipoDocRef`,`Asd_dFecDocRef`,`Asd_cSerieDocRef`,`Asd_cNumDocRef`,`Asd_cRetencion`,`Asd_cProvCanc`,`Asd_cOperaTC`,`Asd_cTipoMoneda`,`Tra_cCodigo`,`Asd_cFormaPago` f"
        oCustomTask8.SourceSQLStatement = oCustomTask8.SourceSQLStatement & "rom `CAJAEGR_DET$`"
        oCustomTask8.DestinationConnectionID = 4
        oCustomTask8.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET]"
        oCustomTask8.ProgressRowCount = 1000
        oCustomTask8.MaximumErrorCount = 0
        oCustomTask8.FetchBufferSize = 1
        oCustomTask8.UseFastLoad = True
        oCustomTask8.InsertCommitSize = 0
        oCustomTask8.ExceptionFileColumnDelimiter = "|"
        oCustomTask8.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask8.AllowIdentityInserts = False
        oCustomTask8.FirstRow = 0
        oCustomTask8.LastRow = 0
        oCustomTask8.FastLoadOptions = 2
        oCustomTask8.ExceptionFileOptions = 1
        oCustomTask8.DataPumpOptions = 0

Call oCustomTask8_Trans_Sub1(oCustomTask8)


goPackage.Tasks.Add oTask
Set oCustomTask8 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask8_Trans_Sub1(ByVal oCustomTask8 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask8.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDocRef", 22)
                        oColumn.Name = "Asd_cTipoDocRef"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDocRef", 23)
                        oColumn.Name = "Asd_dFecDocRef"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDocRef", 24)
                        oColumn.Name = "Asd_cSerieDocRef"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDocRef", 25)
                        oColumn.Name = "Asd_cNumDocRef"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cRetencion", 26)
                        oColumn.Name = "Asd_cRetencion"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cProvCanc", 27)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 27
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cOperaTC", 28)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 28
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoMoneda", 29)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 29
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tra_cCodigo", 30)
                        oColumn.Name = "Tra_cCodigo"
                        oColumn.Ordinal = 30
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cFormaPago", 31)
                        oColumn.Name = "Asd_cFormaPago"
                        oColumn.Ordinal = 31
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDocRef", 22)
                        oColumn.Name = "Asd_cTipoDocRef"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDocRef", 23)
                        oColumn.Name = "Asd_dFecDocRef"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDocRef", 24)
                        oColumn.Name = "Asd_cSerieDocRef"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDocRef", 25)
                        oColumn.Name = "Asd_cNumDocRef"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cRetencion", 26)
                        oColumn.Name = "Asd_cRetencion"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cProvCanc", 27)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 27
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cOperaTC", 28)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 28
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoMoneda", 29)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 29
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tra_cCodigo", 30)
                        oColumn.Name = "Tra_cCodigo"
                        oColumn.Ordinal = 30
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cFormaPago", 31)
                        oColumn.Name = "Asd_cFormaPago"
                        oColumn.Ordinal = 31
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask8.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub9 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea)
Public Sub Task_Sub9(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask9 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
Set oCustomTask9 = oTask.CustomTask

        oCustomTask9.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
        oCustomTask9.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"

        oCustomTask9.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "drop table [dbo].[ZIMP_CAJAING_CAB]" & vbCrLf


        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] (" & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf
        oCustomTask9.SQLStatement = oCustomTask9.SQLStatement & ")"
        oCustomTask9.ConnectionID = 2
        oCustomTask9.CommandTimeout = 0
        oCustomTask9.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask9 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub10 for task Copy Data from CAJAING_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea (Copy Data from CAJAING_CAB$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_CAB] Tarea)
Public Sub Task_Sub10(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask10 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
Set oCustomTask10 = oTask.CustomTask

        oCustomTask10.Name = "Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
        oCustomTask10.Description = "Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
        oCustomTask10.SourceConnectionID = 1
        oCustomTask10.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Ase_dFecha`,`Ase_cGlosa`,`Ase_cTipoMoneda` from `CAJAING_CAB$`"
        oCustomTask10.DestinationConnectionID = 2
        oCustomTask10.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB]"
        oCustomTask10.ProgressRowCount = 1000
        oCustomTask10.MaximumErrorCount = 0
        oCustomTask10.FetchBufferSize = 1
        oCustomTask10.UseFastLoad = True
        oCustomTask10.InsertCommitSize = 0
        oCustomTask10.ExceptionFileColumnDelimiter = "|"
        oCustomTask10.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask10.AllowIdentityInserts = False
        oCustomTask10.FirstRow = 0
        oCustomTask10.LastRow = 0
        oCustomTask10.FastLoadOptions = 2
        oCustomTask10.ExceptionFileOptions = 1
        oCustomTask10.DataPumpOptions = 0

Call oCustomTask10_Trans_Sub1(oCustomTask10)


goPackage.Tasks.Add oTask
Set oCustomTask10 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask10_Trans_Sub1(ByVal oCustomTask10 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask10.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask10.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub11 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea)
Public Sub Task_Sub11(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask11 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
Set oCustomTask11 = oTask.CustomTask

        oCustomTask11.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
        oCustomTask11.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"

        oCustomTask11.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "drop table [dbo].[ZIMP_CAJAING_DET]" & vbCrLf


        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] (" & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_nItem] INT NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_dFecDoc] datetime NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_dFecVen] datetime NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_dFecDocRef] datetime NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cNumDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cRetencion] nvarchar (255) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cTipoMoneda] char (3) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Tra_cCodigo] nvarchar (255) NULL, " & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & "[Asd_cFormaPago] nvarchar (255) NULL" & vbCrLf
        oCustomTask11.SQLStatement = oCustomTask11.SQLStatement & ")"
        oCustomTask11.ConnectionID = 4
        oCustomTask11.CommandTimeout = 0
        oCustomTask11.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask11 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub12 for task Copy Data from CAJAING_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea (Copy Data from CAJAING_DET$ to [SAFC_ECB].[dbo].[ZIMP_CAJAING_DET] Tarea)
Public Sub Task_Sub12(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask12 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
Set oCustomTask12 = oTask.CustomTask

        oCustomTask12.Name = "Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
        oCustomTask12.Description = "Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
        oCustomTask12.SourceConnectionID = 3
        oCustomTask12.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Pla_cCuentaContable`,`Asd_nItem`,`Asd_cGlosa`,`Asd_nDebeSoles`,`Asd_nHaberSoles`,`Asd_nTipoCambio`,`Asd_nDebeMonExt`,`Asd_nHaberMonExt`,`Cos_cCodigo`,`Ten_cTipoEntidad`,`Ent_"
        oCustomTask12.SourceSQLStatement = oCustomTask12.SourceSQLStatement & "cCodEntidad`,`Asd_cTipoDoc`,`Asd_dFecDoc`,`Asd_cSerieDoc`,`Asd_cNumDoc`,`Asd_dFecVen`,`Asd_cTipoDocRef`,`Asd_dFecDocRef`,`Asd_cSerieDocRef`,`Asd_cNumDocRef`,`Asd_cRetencion`,`Asd_cProvCanc`,`Asd_cOperaTC`,`Asd_cTipoMoneda`,`Tra_cCodigo`,`Asd_cFormaPago` f"
        oCustomTask12.SourceSQLStatement = oCustomTask12.SourceSQLStatement & "rom `CAJAING_DET$`"
        oCustomTask12.DestinationConnectionID = 4
        oCustomTask12.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET]"
        oCustomTask12.ProgressRowCount = 1000
        oCustomTask12.MaximumErrorCount = 0
        oCustomTask12.FetchBufferSize = 1
        oCustomTask12.UseFastLoad = True
        oCustomTask12.InsertCommitSize = 0
        oCustomTask12.ExceptionFileColumnDelimiter = "|"
        oCustomTask12.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask12.AllowIdentityInserts = False
        oCustomTask12.FirstRow = 0
        oCustomTask12.LastRow = 0
        oCustomTask12.FastLoadOptions = 2
        oCustomTask12.ExceptionFileOptions = 1
        oCustomTask12.DataPumpOptions = 0

Call oCustomTask12_Trans_Sub1(oCustomTask12)


goPackage.Tasks.Add oTask
Set oCustomTask12 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask12_Trans_Sub1(ByVal oCustomTask12 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask12.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDocRef", 22)
                        oColumn.Name = "Asd_cTipoDocRef"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDocRef", 23)
                        oColumn.Name = "Asd_dFecDocRef"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDocRef", 24)
                        oColumn.Name = "Asd_cSerieDocRef"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDocRef", 25)
                        oColumn.Name = "Asd_cNumDocRef"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cRetencion", 26)
                        oColumn.Name = "Asd_cRetencion"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cProvCanc", 27)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 27
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cOperaTC", 28)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 28
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoMoneda", 29)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 29
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tra_cCodigo", 30)
                        oColumn.Name = "Tra_cCodigo"
                        oColumn.Ordinal = 30
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cFormaPago", 31)
                        oColumn.Name = "Asd_cFormaPago"
                        oColumn.Ordinal = 31
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDocRef", 22)
                        oColumn.Name = "Asd_cTipoDocRef"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDocRef", 23)
                        oColumn.Name = "Asd_dFecDocRef"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDocRef", 24)
                        oColumn.Name = "Asd_cSerieDocRef"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDocRef", 25)
                        oColumn.Name = "Asd_cNumDocRef"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cRetencion", 26)
                        oColumn.Name = "Asd_cRetencion"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cProvCanc", 27)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 27
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cOperaTC", 28)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 28
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoMoneda", 29)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 29
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tra_cCodigo", 30)
                        oColumn.Name = "Tra_cCodigo"
                        oColumn.Ordinal = 30
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cFormaPago", 31)
                        oColumn.Name = "Asd_cFormaPago"
                        oColumn.Ordinal = 31
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask12.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub13 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea)
Public Sub Task_Sub13(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask13 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
Set oCustomTask13 = oTask.CustomTask

        oCustomTask13.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
        oCustomTask13.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"

        oCustomTask13.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "drop table [dbo].[ZIMP_COMPRAS_CAB]" & vbCrLf

        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] (" & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Ase_dFecha] DATETIME NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf
        oCustomTask13.SQLStatement = oCustomTask13.SQLStatement & ")"
        oCustomTask13.ConnectionID = 2
        oCustomTask13.CommandTimeout = 0
        oCustomTask13.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask13 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub14 for task Copy Data from COMPRAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea (Copy Data from COMPRAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_CAB] Tarea)
Public Sub Task_Sub14(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask14 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
Set oCustomTask14 = oTask.CustomTask

        oCustomTask14.Name = "Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
        oCustomTask14.Description = "Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
        oCustomTask14.SourceConnectionID = 1
        oCustomTask14.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Ase_dFecha`,`Ase_cGlosa`,`Ase_cTipoMoneda` from `COMPRAS_CAB$`"
        oCustomTask14.DestinationConnectionID = 2
        oCustomTask14.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB]"
        oCustomTask14.ProgressRowCount = 1000
        oCustomTask14.MaximumErrorCount = 0
        oCustomTask14.FetchBufferSize = 1
        oCustomTask14.UseFastLoad = True
        oCustomTask14.InsertCommitSize = 0
        oCustomTask14.ExceptionFileColumnDelimiter = "|"
        oCustomTask14.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask14.AllowIdentityInserts = False
        oCustomTask14.FirstRow = 0
        oCustomTask14.LastRow = 0
        oCustomTask14.FastLoadOptions = 2
        oCustomTask14.ExceptionFileOptions = 1
        oCustomTask14.DataPumpOptions = 0

Call oCustomTask14_Trans_Sub1(oCustomTask14)


goPackage.Tasks.Add oTask
Set oCustomTask14 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask14_Trans_Sub1(ByVal oCustomTask14 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask14.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask14.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub15 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea)
Public Sub Task_Sub15(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask15 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
Set oCustomTask15 = oTask.CustomTask

        oCustomTask15.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
        oCustomTask15.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"

        oCustomTask15.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "drop table [dbo].[ZIMP_COMPRAS_DET]" & vbCrLf

        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] (" & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nItem] INT NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_dFecDoc] datetime NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_dFecVen] datetime NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_dFecDocRef] datetime NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cNumDocRef] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cBaseImp] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cRetencion] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_dFechaSpot] datetime NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cNumSpot] nvarchar (255) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cTipoMoneda] char (3) NULL, " & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & "[Asd_cComprobante] nvarchar (255) NULL" & vbCrLf
        oCustomTask15.SQLStatement = oCustomTask15.SQLStatement & ")"
        oCustomTask15.ConnectionID = 4
        oCustomTask15.CommandTimeout = 0
        oCustomTask15.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask15 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub16 for task Copy Data from COMPRAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea (Copy Data from COMPRAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_COMPRAS_DET] Tarea)
Public Sub Task_Sub16(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask16 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
Set oCustomTask16 = oTask.CustomTask

        oCustomTask16.Name = "Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
        oCustomTask16.Description = "Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
        oCustomTask16.SourceConnectionID = 3
        oCustomTask16.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Pla_cCuentaContable`,`Asd_nItem`,`Asd_cGlosa`,`Asd_nDebeSoles`,`Asd_nHaberSoles`,`Asd_nTipoCambio`,`Asd_nDebeMonExt`,`Asd_nHaberMonExt`,`Cos_cCodigo`,`Ten_cTipoEntidad`,`Ent_"
        oCustomTask16.SourceSQLStatement = oCustomTask16.SourceSQLStatement & "cCodEntidad`,`Asd_cTipoDoc`,`Asd_dFecDoc`,`Asd_cSerieDoc`,`Asd_cNumDoc`,`Asd_dFecVen`,`Asd_cTipoDocRef`,`Asd_dFecDocRef`,`Asd_cSerieDocRef`,`Asd_cNumDocRef`,`Asd_nMontoInafecto`,`Asd_cBaseImp`,`Asd_cRetencion`,`Asd_dFechaSpot`,`Asd_cNumSpot`,`Asd_cProvCan"
        oCustomTask16.SourceSQLStatement = oCustomTask16.SourceSQLStatement & "c`,`Asd_cOperaTC`,`Asd_cTipoMoneda`,`Asd_cComprobante` from `COMPRAS_DET$`"
        oCustomTask16.DestinationConnectionID = 4
        oCustomTask16.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET]"
        oCustomTask16.ProgressRowCount = 1000
        oCustomTask16.MaximumErrorCount = 0
        oCustomTask16.FetchBufferSize = 1
        oCustomTask16.UseFastLoad = True
        oCustomTask16.InsertCommitSize = 0
        oCustomTask16.ExceptionFileColumnDelimiter = "|"
        oCustomTask16.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask16.AllowIdentityInserts = False
        oCustomTask16.FirstRow = 0
        oCustomTask16.LastRow = 0
        oCustomTask16.FastLoadOptions = 2
        oCustomTask16.ExceptionFileOptions = 1
        oCustomTask16.DataPumpOptions = 0

Call oCustomTask16_Trans_Sub1(oCustomTask16)


goPackage.Tasks.Add oTask
Set oCustomTask16 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask16_Trans_Sub1(ByVal oCustomTask16 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask16.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDocRef", 22)
                        oColumn.Name = "Asd_cTipoDocRef"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDocRef", 23)
                        oColumn.Name = "Asd_dFecDocRef"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDocRef", 24)
                        oColumn.Name = "Asd_cSerieDocRef"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDocRef", 25)
                        oColumn.Name = "Asd_cNumDocRef"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nMontoInafecto", 26)
                        oColumn.Name = "Asd_nMontoInafecto"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cBaseImp", 27)
                        oColumn.Name = "Asd_cBaseImp"
                        oColumn.Ordinal = 27
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cRetencion", 28)
                        oColumn.Name = "Asd_cRetencion"
                        oColumn.Ordinal = 28
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFechaSpot", 29)
                        oColumn.Name = "Asd_dFechaSpot"
                        oColumn.Ordinal = 29
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumSpot", 30)
                        oColumn.Name = "Asd_cNumSpot"
                        oColumn.Ordinal = 30
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cProvCanc", 31)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 31
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cOperaTC", 32)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 32
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoMoneda", 33)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 33
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cComprobante", 34)
                        oColumn.Name = "Asd_cComprobante"
                        oColumn.Ordinal = 34
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDocRef", 22)
                        oColumn.Name = "Asd_cTipoDocRef"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDocRef", 23)
                        oColumn.Name = "Asd_dFecDocRef"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDocRef", 24)
                        oColumn.Name = "Asd_cSerieDocRef"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDocRef", 25)
                        oColumn.Name = "Asd_cNumDocRef"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nMontoInafecto", 26)
                        oColumn.Name = "Asd_nMontoInafecto"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cBaseImp", 27)
                        oColumn.Name = "Asd_cBaseImp"
                        oColumn.Ordinal = 27
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cRetencion", 28)
                        oColumn.Name = "Asd_cRetencion"
                        oColumn.Ordinal = 28
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFechaSpot", 29)
                        oColumn.Name = "Asd_dFechaSpot"
                        oColumn.Ordinal = 29
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumSpot", 30)
                        oColumn.Name = "Asd_cNumSpot"
                        oColumn.Ordinal = 30
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cProvCanc", 31)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 31
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cOperaTC", 32)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 32
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoMoneda", 33)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 33
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cComprobante", 34)
                        oColumn.Name = "Asd_cComprobante"
                        oColumn.Ordinal = 34
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask16.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub17 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea)
Public Sub Task_Sub17(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask17 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
Set oCustomTask17 = oTask.CustomTask

        oCustomTask17.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
        oCustomTask17.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"

        oCustomTask17.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_ENTIDAD]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "drop table [dbo].[ZIMP_ENTIDAD]" & vbCrLf


        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] (" & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ent_cPersona] nvarchar (255) NULL, " & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ent_cDireccion] nvarchar (255) NULL, " & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ent_nRuc] nvarchar (255) NULL, " & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ent_cTipoDoc] nvarchar (255) NULL, " & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & "[Ent_cFlagPersona] nvarchar (255) NULL" & vbCrLf
        oCustomTask17.SQLStatement = oCustomTask17.SQLStatement & ")"
        oCustomTask17.ConnectionID = 2
        oCustomTask17.CommandTimeout = 0
        oCustomTask17.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask17 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub18 for task Copy Data from ENTIDAD$ to [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea (Copy Data from ENTIDAD$ to [SAFC_ECB].[dbo].[ZIMP_ENTIDAD] Tarea)
Public Sub Task_Sub18(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask18 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
Set oCustomTask18 = oTask.CustomTask

        oCustomTask18.Name = "Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
        oCustomTask18.Description = "Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
        oCustomTask18.SourceConnectionID = 1
        oCustomTask18.SourceSQLStatement = "select `Ent_cCodEntidad`,`Ten_cTipoEntidad`,`Ent_cPersona`,`Ent_cDireccion`,`Ent_nRuc`,`Ent_cTipoDoc`,`Ent_cFlagPersona` from `ENTIDAD$`"
        oCustomTask18.DestinationConnectionID = 2
        oCustomTask18.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD]"
        oCustomTask18.ProgressRowCount = 1000
        oCustomTask18.MaximumErrorCount = 0
        oCustomTask18.FetchBufferSize = 1
        oCustomTask18.UseFastLoad = True
        oCustomTask18.InsertCommitSize = 0
        oCustomTask18.ExceptionFileColumnDelimiter = "|"
        oCustomTask18.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask18.AllowIdentityInserts = False
        oCustomTask18.FirstRow = 0
        oCustomTask18.LastRow = 0
        oCustomTask18.FastLoadOptions = 2
        oCustomTask18.ExceptionFileOptions = 1
        oCustomTask18.DataPumpOptions = 0

Call oCustomTask18_Trans_Sub1(oCustomTask18)


goPackage.Tasks.Add oTask
Set oCustomTask18 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask18_Trans_Sub1(ByVal oCustomTask18 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask18.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 1)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 2)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cPersona", 3)
                        oColumn.Name = "Ent_cPersona"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cDireccion", 4)
                        oColumn.Name = "Ent_cDireccion"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_nRuc", 5)
                        oColumn.Name = "Ent_nRuc"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cTipoDoc", 6)
                        oColumn.Name = "Ent_cTipoDoc"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cFlagPersona", 7)
                        oColumn.Name = "Ent_cFlagPersona"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 1)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 2)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cPersona", 3)
                        oColumn.Name = "Ent_cPersona"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cDireccion", 4)
                        oColumn.Name = "Ent_cDireccion"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_nRuc", 5)
                        oColumn.Name = "Ent_nRuc"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cTipoDoc", 6)
                        oColumn.Name = "Ent_cTipoDoc"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cFlagPersona", 7)
                        oColumn.Name = "Ent_cFlagPersona"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 104
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask18.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub19 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea)
Public Sub Task_Sub19(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask19 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
Set oCustomTask19 = oTask.CustomTask

        oCustomTask19.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
        oCustomTask19.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"

        oCustomTask19.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_TIPOCAMBIO]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "drop table [dbo].[ZIMP_TIPOCAMBIO]" & vbCrLf

        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] (" & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "[Tca_dFecha] DATETIME NULL, " & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "[Tca_cCodigoOrigen] nvarchar (255) NULL, " & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "[Tca_cCodigoDestino] nvarchar (255) NULL, " & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "[Tca_nCompra] nvarchar (255) NULL, " & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "[Tca_nVenta] nvarchar (255) NULL, " & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & "[Tca_nVentaP] nvarchar (255) NULL" & vbCrLf
        oCustomTask19.SQLStatement = oCustomTask19.SQLStatement & ")"
        oCustomTask19.ConnectionID = 4
        oCustomTask19.CommandTimeout = 0
        oCustomTask19.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask19 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub20 for task Copy Data from TIPOCAMBIO$ to [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea (Copy Data from TIPOCAMBIO$ to [SAFC_ECB].[dbo].[ZIMP_TIPOCAMBIO] Tarea)
Public Sub Task_Sub20(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask20 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
Set oCustomTask20 = oTask.CustomTask

        oCustomTask20.Name = "Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
        oCustomTask20.Description = "Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
        oCustomTask20.SourceConnectionID = 3
        oCustomTask20.SourceSQLStatement = "select `Tca_dFecha`,`Tca_cCodigoOrigen`,`Tca_cCodigoDestino`,`Tca_nCompra`,`Tca_nVenta`,`Tca_nVentaP` from `TIPOCAMBIO$`"
        oCustomTask20.DestinationConnectionID = 4
        oCustomTask20.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO]"
        oCustomTask20.ProgressRowCount = 1000
        oCustomTask20.MaximumErrorCount = 0
        oCustomTask20.FetchBufferSize = 1
        oCustomTask20.UseFastLoad = True
        oCustomTask20.InsertCommitSize = 0
        oCustomTask20.ExceptionFileColumnDelimiter = "|"
        oCustomTask20.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask20.AllowIdentityInserts = False
        oCustomTask20.FirstRow = 0
        oCustomTask20.LastRow = 0
        oCustomTask20.FastLoadOptions = 2
        oCustomTask20.ExceptionFileOptions = 1
        oCustomTask20.DataPumpOptions = 0

Call oCustomTask20_Trans_Sub1(oCustomTask20)


goPackage.Tasks.Add oTask
Set oCustomTask20 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask20_Trans_Sub1(ByVal oCustomTask20 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask20.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Tca_dFecha", 1)
                        oColumn.Name = "Tca_dFecha"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tca_cCodigoOrigen", 2)
                        oColumn.Name = "Tca_cCodigoOrigen"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tca_cCodigoDestino", 3)
                        oColumn.Name = "Tca_cCodigoDestino"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tca_nCompra", 4)
                        oColumn.Name = "Tca_nCompra"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tca_nVenta", 5)
                        oColumn.Name = "Tca_nVenta"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Tca_nVentaP", 6)
                        oColumn.Name = "Tca_nVentaP"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tca_dFecha", 1)
                        oColumn.Name = "Tca_dFecha"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tca_cCodigoOrigen", 2)
                        oColumn.Name = "Tca_cCodigoOrigen"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tca_cCodigoDestino", 3)
                        oColumn.Name = "Tca_cCodigoDestino"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tca_nCompra", 4)
                        oColumn.Name = "Tca_nCompra"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tca_nVenta", 5)
                        oColumn.Name = "Tca_nVenta"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Tca_nVentaP", 6)
                        oColumn.Name = "Tca_nVentaP"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask20.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub21 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea)
Public Sub Task_Sub21(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask21 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
Set oCustomTask21 = oTask.CustomTask

        oCustomTask21.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
        oCustomTask21.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"

        oCustomTask21.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "drop table [dbo].[ZIMP_VENTAS_CAB]" & vbCrLf


        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] (" & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf
        oCustomTask21.SQLStatement = oCustomTask21.SQLStatement & ")"
        oCustomTask21.ConnectionID = 2
        oCustomTask21.CommandTimeout = 0
        oCustomTask21.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask21 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub22 for task Copy Data from VENTAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea (Copy Data from VENTAS_CAB$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_CAB] Tarea)
Public Sub Task_Sub22(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask22 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
Set oCustomTask22 = oTask.CustomTask

        oCustomTask22.Name = "Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
        oCustomTask22.Description = "Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
        oCustomTask22.SourceConnectionID = 1
        oCustomTask22.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Ase_dFecha`,`Ase_cGlosa`,`Ase_cTipoMoneda` from `VENTAS_CAB$`"
        oCustomTask22.DestinationConnectionID = 2
        oCustomTask22.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB]"
        oCustomTask22.ProgressRowCount = 1000
        oCustomTask22.MaximumErrorCount = 0
        oCustomTask22.FetchBufferSize = 1
        oCustomTask22.UseFastLoad = True
        oCustomTask22.InsertCommitSize = 0
        oCustomTask22.ExceptionFileColumnDelimiter = "|"
        oCustomTask22.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask22.AllowIdentityInserts = False
        oCustomTask22.FirstRow = 0
        oCustomTask22.LastRow = 0
        oCustomTask22.FastLoadOptions = 2
        oCustomTask22.ExceptionFileOptions = 1
        oCustomTask22.DataPumpOptions = 0

Call oCustomTask22_Trans_Sub1(oCustomTask22)


goPackage.Tasks.Add oTask
Set oCustomTask22 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask22_Trans_Sub1(ByVal oCustomTask22 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask22.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask22.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub23 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea)
Public Sub Task_Sub23(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask23 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
Set oCustomTask23 = oTask.CustomTask

        oCustomTask23.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
        oCustomTask23.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"


        oCustomTask23.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "drop table [dbo].[ZIMP_VENTAS_DET]" & vbCrLf


        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] (" & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nItem] INT NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_dFecDoc] datetime NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_dFecVen] datetime NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cBaseImp] nvarchar (255) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & "[Asd_cTipoMoneda] char (3) NULL" & vbCrLf
        oCustomTask23.SQLStatement = oCustomTask23.SQLStatement & ")"
        oCustomTask23.ConnectionID = 4
        oCustomTask23.CommandTimeout = 0
        oCustomTask23.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask23 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub24 for task Copy Data from VENTAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea (Copy Data from VENTAS_DET$ to [SAFC_ECB].[dbo].[ZIMP_VENTAS_DET] Tarea)
Public Sub Task_Sub24(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask24 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
Set oCustomTask24 = oTask.CustomTask

        oCustomTask24.Name = "Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
        oCustomTask24.Description = "Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
        oCustomTask24.SourceConnectionID = 3
        oCustomTask24.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Pla_cCuentaContable`,`Asd_nItem`,`Asd_cGlosa`,`Asd_nDebeSoles`,`Asd_nHaberSoles`,`Asd_nTipoCambio`,`Asd_nDebeMonExt`,`Asd_nHaberMonExt`,`Cos_cCodigo`,`Ten_cTipoEntidad`,`Ent_"
        oCustomTask24.SourceSQLStatement = oCustomTask24.SourceSQLStatement & "cCodEntidad`,`Asd_cTipoDoc`,`Asd_dFecDoc`,`Asd_cSerieDoc`,`Asd_cNumDoc`,`Asd_dFecVen`,`Asd_nMontoInafecto`,`Asd_cBaseImp`,`Asd_cProvCanc`,`Asd_cOperaTC`,`Asd_cTipoMoneda` from `VENTAS_DET$`"
        oCustomTask24.DestinationConnectionID = 4
        oCustomTask24.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET]"
        oCustomTask24.ProgressRowCount = 1000
        oCustomTask24.MaximumErrorCount = 0
        oCustomTask24.FetchBufferSize = 1
        oCustomTask24.UseFastLoad = True
        oCustomTask24.InsertCommitSize = 0
        oCustomTask24.ExceptionFileColumnDelimiter = "|"
        oCustomTask24.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask24.AllowIdentityInserts = False
        oCustomTask24.FirstRow = 0
        oCustomTask24.LastRow = 0
        oCustomTask24.FastLoadOptions = 2
        oCustomTask24.ExceptionFileOptions = 1
        oCustomTask24.DataPumpOptions = 0

Call oCustomTask24_Trans_Sub1(oCustomTask24)


goPackage.Tasks.Add oTask
Set oCustomTask24 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask24_Trans_Sub1(ByVal oCustomTask24 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask24.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nMontoInafecto", 22)
                        oColumn.Name = "Asd_nMontoInafecto"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cBaseImp", 23)
                        oColumn.Name = "Asd_cBaseImp"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cProvCanc", 24)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cOperaTC", 25)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoMoneda", 26)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nMontoInafecto", 22)
                        oColumn.Name = "Asd_nMontoInafecto"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cBaseImp", 23)
                        oColumn.Name = "Asd_cBaseImp"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cProvCanc", 24)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cOperaTC", 25)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 25
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoMoneda", 26)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 26
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask24.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub


Private Sub IniciaObjetoDTS(sArchivoXLS As String, ByRef cParametros As String)

    Dim cEntidad As String
    Dim cTC As String
    Dim cApertura As String
    Dim cCompras As String
    Dim cVentas As String
    Dim cCajaIng As String
    Dim cCajaEgr As String
    Dim cPlanilla As String

    cEntidad = Mid(cParametros, 1, 1)
    cTC = Mid(cParametros, 2, 1)
    cApertura = Mid(cParametros, 3, 1)
    cCompras = Mid(cParametros, 4, 1)
    cVentas = Mid(cParametros, 5, 1)
    cCajaIng = Mid(cParametros, 6, 1)
    cCajaEgr = Mid(cParametros, 7, 1)
    cPlanilla = Mid(cParametros, 8, 1)
    '---------------------------------------------

    Set goPackage = goPackageOld

    goPackage.Name = "Importacion XLS"
    goPackage.Description = "Descripción del paquete DTS"
    goPackage.WriteCompletionStatusToNTEventLog = False
    goPackage.FailOnError = False
    goPackage.PackagePriorityClass = 2
    goPackage.MaxConcurrentSteps = 4
    goPackage.LineageOptions = 0
    goPackage.UseTransaction = True
    goPackage.TransactionIsolationLevel = 4096
    goPackage.AutoCommitTransaction = True
    goPackage.RepositoryMetadataOptions = 0
    goPackage.UseOLEDBServiceComponents = True
    goPackage.LogToSQLServer = False
    goPackage.LogServerFlags = 0
    goPackage.FailPackageOnLogFailure = False
    goPackage.ExplicitGlobalVariables = False
    goPackage.PackageType = 0


    Dim oConnProperty As DTS.OleDBProperty

    '---------------------------------------------------------------------------
    ' create package connection information
    '---------------------------------------------------------------------------

    Dim oConnection As DTS.Connection2

    '------------- a new connection defined below.
    'For security purposes, the password is never scripted

    'Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
    Set oConnection = goPackage.Connections.New("MicrosoftACE.OLEDB.12.0")

            oConnection.ConnectionProperties("Data Source") = sArchivoXLS
            oConnection.ConnectionProperties("Extended Properties") = "Excel 8.0;HDR=YES;"

            oConnection.Name = "Conexión1"
            oConnection.Id = 1
            oConnection.Reusable = True
            oConnection.ConnectImmediate = False
            oConnection.DataSource = sArchivoXLS
            oConnection.ConnectionTimeout = 60
            oConnection.UseTrustedConnection = False
            oConnection.UseDSL = False

'cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'           "Data Source=c:\somepath\ExcelFile.xls;" & _
'           "Extended Properties=""Excel 8.0;HDR=Yes;"";"

            'If you have a password for this connection, please uncomment and add your password below.
            'oConnection.Password = "<put the password here>"

    goPackage.Connections.Add oConnection
    Set oConnection = Nothing

    '------------- a new connection defined below.
    'For security purposes, the password is never scripted

    Set oConnection = goPackage.Connections.New("SQLOLEDB")

            oConnection.ConnectionProperties("Persist Security Info") = True
            oConnection.ConnectionProperties("User ID") = sUSUARIO
            oConnection.ConnectionProperties("Password") = sPW
            oConnection.ConnectionProperties("Initial Catalog") = sBaseDatos
            oConnection.ConnectionProperties("Data Source") = sSERVIDOR
            oConnection.ConnectionProperties("Application Name") = "Asistente para importación/exportación con DTS"

            oConnection.Name = "Conexión2"
            oConnection.Id = 2
            oConnection.Reusable = True
            oConnection.ConnectImmediate = False
            oConnection.DataSource = sSERVIDOR
            oConnection.UserId = sUSUARIO
            oConnection.password = sPW
            oConnection.ConnectionTimeout = 60
            oConnection.Catalog = sBaseDatos
            oConnection.UseTrustedConnection = False
            oConnection.UseDSL = False

            'If you have a password for this connection, please uncomment and add your password below.
            'oConnection.Password = "<put the password here>"

    goPackage.Connections.Add oConnection
    Set oConnection = Nothing

    '------------- a new connection defined below.
    'For security purposes, the password is never scripted

    Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")

            oConnection.ConnectionProperties("Data Source") = sArchivoXLS
            oConnection.ConnectionProperties("Extended Properties") = "Excel 8.0;HDR=YES;"

            oConnection.Name = "Conexión3"
            oConnection.Id = 3
            oConnection.Reusable = True
            oConnection.ConnectImmediate = False
            oConnection.DataSource = sArchivoXLS
            oConnection.ConnectionTimeout = 60
            oConnection.UseTrustedConnection = False
            oConnection.UseDSL = False

            'If you have a password for this connection, please uncomment and add your password below.
            'oConnection.Password = "<put the password here>"

    goPackage.Connections.Add oConnection
    Set oConnection = Nothing

    '------------- a new connection defined below.
    'For security purposes, the password is never scripted

    Set oConnection = goPackage.Connections.New("SQLOLEDB")

            oConnection.ConnectionProperties("Persist Security Info") = True
            oConnection.ConnectionProperties("User ID") = sUSUARIO
            oConnection.ConnectionProperties("Password") = sPW
            oConnection.ConnectionProperties("Initial Catalog") = sBaseDatos
            oConnection.ConnectionProperties("Data Source") = sSERVIDOR
            oConnection.ConnectionProperties("Application Name") = "Asistente para importación/exportación con DTS"

            oConnection.Name = "Conexión4"
            oConnection.Id = 4
            oConnection.Reusable = True
            oConnection.ConnectImmediate = False
            oConnection.DataSource = sSERVIDOR
            oConnection.UserId = sUSUARIO
            oConnection.password = sPW
            oConnection.ConnectionTimeout = 60
            oConnection.Catalog = sBaseDatos
            oConnection.UseTrustedConnection = False
            oConnection.UseDSL = False

            'If you have a password for this connection, please uncomment and add your password below.
            'oConnection.Password = "<put the password here>"

    goPackage.Connections.Add oConnection
    Set oConnection = Nothing

    '---------------------------------------------------------------------------
    ' create package steps information
    '---------------------------------------------------------------------------

    Dim oStep As DTS.Step2
    Dim oPrecConstraint As DTS.PrecedenceConstraint

If cApertura = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso"
            oStep.Description = "Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso"
            oStep.Description = "Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing
End If


If cCajaEgr = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso"
            oStep.Description = "Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso"
            oStep.Description = "Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

End If

If cCajaIng = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso"
            oStep.Description = "Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso"
            oStep.Description = "Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

End If

If cCompras = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso"
            oStep.Description = "Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso"
            oStep.Description = "Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

End If

If cPlanilla = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso"
            oStep.Description = "Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso"
            oStep.Description = "Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

End If
If cEntidad = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso"
            oStep.Description = "Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing
End If

If cTC = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso"
            oStep.Description = "Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

End If

If cVentas = "1" Then
    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso"
            oStep.Description = "Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso"
            oStep.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = False
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

    '------------- a new step defined below

    Set oStep = goPackage.Steps.New

            oStep.Name = "Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso"
            oStep.Description = "Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso"
            oStep.ExecutionStatus = 1
            oStep.TaskName = "Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Tarea"
            oStep.CommitSuccess = False
            oStep.RollbackFailure = False
            oStep.ScriptLanguage = "VBScript"
            oStep.AddGlobalVariables = True
            oStep.RelativePriority = 3
            oStep.CloseConnection = False
            oStep.ExecuteInMainThread = True
            oStep.IsPackageDSORowset = False
            oStep.JoinTransactionIfPresent = False
            oStep.DisableStep = False
            oStep.FailPackageOnError = False

    goPackage.Steps.Add oStep
    Set oStep = Nothing

End If


If cApertura = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from APE_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from APE_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

End If

If cCajaEgr = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from CAJAEGR_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from CAJAEGR_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing
End If

If cCajaIng = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from CAJAING_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from CAJAING_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing
End If

If cCompras = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from COMPRAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from COMPRAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

End If

If cPlanilla = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing
End If

If cEntidad = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from ENTIDAD$ to [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

End If

If cTC = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from TIPOCAMBIO$ to [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

End If

If cVentas = "1" Then
    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from VENTAS_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing

    '------------- a precedence constraint for steps defined below

    Set oStep = goPackage.Steps("Copy Data from VENTAS_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso")
    Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso")
            oPrecConstraint.StepName = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] Paso"
            oPrecConstraint.PrecedenceBasis = 0
            oPrecConstraint.Value = 4

    oStep.PrecedenceConstraints.Add oPrecConstraint
    Set oPrecConstraint = Nothing
End If


End Sub


'------------- define Task_Sub1 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_APE_CAB] Tarea)
Public Sub Task_Sub25(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
        oCustomTask1.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"

        oCustomTask1.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "drop table [dbo].[ZIMP_PLAN_CAB]" & vbCrLf

        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from PLAN_CAB$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_CAB] Tarea (Copy Data from PLAN_CAB$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_CAB] Tarea)
Public Sub Task_Sub26(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
        oCustomTask2.Description = "Copy Data from PLAN_CAB$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] Tarea"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Ase_dFecha`,`Ase_cGlosa`,`Ase_cTipoMoneda` from `PLAN_CAB$`"
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB]"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0

Call oCustomTask2_Trans_Sub25(oCustomTask2)


goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask2_Trans_Sub25(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_dFecha", 6)
                        oColumn.Name = "Ase_dFecha"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cGlosa", 7)
                        oColumn.Name = "Ase_cGlosa"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cTipoMoneda", 8)
                        oColumn.Name = "Ase_cTipoMoneda"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub3 for task Crear tabla [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea (Crear tabla [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea)
Public Sub Task_Sub27(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask3 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
Set oCustomTask3 = oTask.CustomTask

        oCustomTask3.Name = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
        oCustomTask3.Description = "Crear tabla [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"

        oCustomTask3.SQLStatement = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "drop table [dbo].[ZIMP_PLAN_DET]" & vbCrLf

        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] (" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ase_cNummov] char (10) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Pan_cAnio] char (4) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Per_cPeriodo] char (2) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ase_nVoucher] char (10) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nItem] INT NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_dFecDoc] datetime null , " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_dFecVen] datetime NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[Asd_cTipoMoneda] char (3) NULL" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & ")"
        oCustomTask3.ConnectionID = 4
        oCustomTask3.CommandTimeout = 0
        oCustomTask3.OutputAsRecordset = False

goPackage.Tasks.Add oTask
Set oCustomTask3 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub4 for task Copy Data from PLAN_DET$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea (Copy Data from PLAN_DET$ to [SAFC_ECB].[dbo].[ZIMP_PLAN_DET] Tarea)
Public Sub Task_Sub28(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask4 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
Set oCustomTask4 = oTask.CustomTask

        oCustomTask4.Name = "Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
        oCustomTask4.Description = "Copy Data from PLAN_DET$ to [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] Tarea"
        oCustomTask4.SourceConnectionID = 3
        oCustomTask4.SourceSQLStatement = "select `Ase_cNummov`,`Pan_cAnio`,`Per_cPeriodo`,`Lib_cTipoLibro`,`Ase_nVoucher`,`Pla_cCuentaContable`,`Asd_nItem`,`Asd_cGlosa`,`Asd_nDebeSoles`,`Asd_nHaberSoles`,`Asd_nTipoCambio`,`Asd_nDebeMonExt`,`Asd_nHaberMonExt`,`Cos_cCodigo`,`Ten_cTipoEntidad`,`Ent_"
        oCustomTask4.SourceSQLStatement = oCustomTask4.SourceSQLStatement & "cCodEntidad`,`Asd_cTipoDoc`,`Asd_dFecDoc`,`Asd_cSerieDoc`,`Asd_cNumDoc`,`Asd_dFecVen`,`Asd_cProvCanc`,`Asd_cOperaTC`,`Asd_cTipoMoneda` from `PLAN_DET$`"
        oCustomTask4.DestinationConnectionID = 4
        oCustomTask4.DestinationObjectName = "[" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET]"
        oCustomTask4.ProgressRowCount = 1000
        oCustomTask4.MaximumErrorCount = 0
        oCustomTask4.FetchBufferSize = 1
        oCustomTask4.UseFastLoad = True
        oCustomTask4.InsertCommitSize = 0
        oCustomTask4.ExceptionFileColumnDelimiter = "|"
        oCustomTask4.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask4.AllowIdentityInserts = False
        oCustomTask4.FirstRow = 0
        oCustomTask4.LastRow = 0
        oCustomTask4.FastLoadOptions = 2
        oCustomTask4.ExceptionFileOptions = 1
        oCustomTask4.DataPumpOptions = 0

Call oCustomTask4_Trans_Sub28(oCustomTask4)


goPackage.Tasks.Add oTask
Set oCustomTask4 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask4_Trans_Sub28(ByVal oCustomTask4 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask4.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4

                Set oColumn = oTransformation.SourceColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 9
                        oColumn.DataType = adNumeric
                        oColumn.Precision = 14
                        oColumn.NumericScale = 3
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cProvCanc", 22)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cOperaTC", 23)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Asd_cTipoMoneda", 24)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_cNummov", 1)
                        oColumn.Name = "Ase_cNummov"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pan_cAnio", 2)
                        oColumn.Name = "Pan_cAnio"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Per_cPeriodo", 3)
                        oColumn.Name = "Per_cPeriodo"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Lib_cTipoLibro", 4)
                        oColumn.Name = "Lib_cTipoLibro"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ase_nVoucher", 5)
                        oColumn.Name = "Ase_nVoucher"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Pla_cCuentaContable", 6)
                        oColumn.Name = "Pla_cCuentaContable"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nItem", 7)
                        oColumn.Name = "Asd_nItem"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cGlosa", 8)
                        oColumn.Name = "Asd_cGlosa"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeSoles", 9)
                        oColumn.Name = "Asd_nDebeSoles"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberSoles", 10)
                        oColumn.Name = "Asd_nHaberSoles"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nTipoCambio", 11)
                        oColumn.Name = "Asd_nTipoCambio"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nDebeMonExt", 12)
                        oColumn.Name = "Asd_nDebeMonExt"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_nHaberMonExt", 13)
                        oColumn.Name = "Asd_nHaberMonExt"
                        oColumn.Ordinal = 13
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Cos_cCodigo", 14)
                        oColumn.Name = "Cos_cCodigo"
                        oColumn.Ordinal = 14
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ten_cTipoEntidad", 15)
                        oColumn.Name = "Ten_cTipoEntidad"
                        oColumn.Ordinal = 15
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Ent_cCodEntidad", 16)
                        oColumn.Name = "Ent_cCodEntidad"
                        oColumn.Ordinal = 16
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoDoc", 17)
                        oColumn.Name = "Asd_cTipoDoc"
                        oColumn.Ordinal = 17
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecDoc", 18)
                        oColumn.Name = "Asd_dFecDoc"
                        oColumn.Ordinal = 18
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cSerieDoc", 19)
                        oColumn.Name = "Asd_cSerieDoc"
                        oColumn.Ordinal = 19
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cNumDoc", 20)
                        oColumn.Name = "Asd_cNumDoc"
                        oColumn.Ordinal = 20
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_dFecVen", 21)
                        oColumn.Name = "Asd_dFecVen"
                        oColumn.Ordinal = 21
                        oColumn.Flags = 102
                        oColumn.Size = 8
                        oColumn.DataType = adDate
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cProvCanc", 22)
                        oColumn.Name = "Asd_cProvCanc"
                        oColumn.Ordinal = 22
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cOperaTC", 23)
                        oColumn.Name = "Asd_cOperaTC"
                        oColumn.Ordinal = 23
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Asd_cTipoMoneda", 24)
                        oColumn.Name = "Asd_cTipoMoneda"
                        oColumn.Ordinal = 24
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True

                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties


        Set oTransProps = Nothing

        oCustomTask4.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

