Attribute VB_Name = "modBusqueda"

Public Function BuscarEntidad(ByRef oFormCall As Form, Optional ByVal tipoEnt As String = "T", Optional ByVal sCadBusq As String = "") As String
Dim gtxtSQL As String, sCaption As String
On Error Resume Next
    BuscarEntidad = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then
        gtxtSQL = "SELECT E.Ent_cCodEntidad, E.Ten_cTipoEntidad, E.Ent_cPersona, E.Ent_nRuc, TD.Tab_cDescripCampo As TipoDoc" _
                & " FROM CNM_ENTIDAD E WITH(NOLOCK) " _
                & " LEFT JOIN TABLA TD WITH(NOLOCK) ON E.Emp_cCodigo = TD.Emp_cCodigo AND TD.Tab_cTabla = '003' AND E.Ent_cTipoDoc = TD.Tab_cCodigo" _
                & " WHERE E.Emp_cCodigo = '" & gsEmpresa & "'" _
                & " AND E.Ten_cTipoEntidad = '" & Left(tipoEnt, 1) & "'"
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            gtxtSQL = "SELECT Ten_cNombreEntidad As Valor FROM CNT_ENTIDAD WITH(NOLOCK) " _
                    & " WHERE Emp_cCodigo = '" & gsEmpresa & "'" & " AND Ten_cTipoEntidad = '" & Left(tipoEnt, 1) & "'"
            sCaption = fRetornaValor(gtxtSQL)
            oFormCall.Enabled = False
            frmBusqueda.Caption = "BUSQUEDA DE " & IIf(sCaption = "", "ENTIDADES", sCaption)
            frmBusqueda.tdbgListado.Columns(0).Caption = "Código"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Ent_cCodEntidad"
            frmBusqueda.tdbgListado.Columns(0).Width = 750
            frmBusqueda.tdbgListado.Columns(1).Caption = "Entidad"
            frmBusqueda.tdbgListado.Columns(1).DataField = "Ent_cPersona"
            frmBusqueda.tdbgListado.Columns(1).Width = 3600
            frmBusqueda.tdbgListado.Columns(2).Caption = "Tipo Doc."
            frmBusqueda.tdbgListado.Columns(2).DataField = "TipoDoc"
            frmBusqueda.tdbgListado.Columns(2).Width = 1650
            frmBusqueda.tdbgListado.Columns(3).Caption = "Doc.Id."
            frmBusqueda.tdbgListado.Columns(3).DataField = "Ent_nRuc"
            frmBusqueda.tdbgListado.Columns(3).Width = 1500
            frmBusqueda.tdbgListado.Columns(3).Visible = True
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen Entidades."
        End If
        BuscarEntidad = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function BuscaNombreCC(ByVal cCodigo As String, ByRef bExiste As Boolean)
Dim gtxtSQL As String
Dim rsCta As New ADODB.Recordset
    BuscaNombreCC = ""
    cCodigo = CE(cCodigo)
    If cCodigo <> "" Then
        gtxtSQL = "spCn_GrabaCentroCosto 'BUSCAR_REGNOTIT', '" & gsEmpresa & "', '" & gsAnio & "', '" & cCodigo & "'"
        Set rsCta = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(rsCta) > 0 Then
            BuscaNombreCC = CE(rsCta!Cos_cDescripcion)
            bExiste = True
        Else
            'Mensajes "Cuenta Contable no existe"
            bExiste = False
        End If
        Call CerrarRecordSet(rsCta)
    End If
End Function

Public Function BuscarConcepto(ByRef oFormCall As Form, Optional cTipo As String = "", Optional ByVal sCadBusq As String = "") As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarConcepto = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then

        gtxtSQL = "spCn_GrabaCuentaRubro 'BUSCARTODOS', '" & gsEmpresa & "','" & cTipo & "','','','','','','" & gsAnio & "'"
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            frmBusqueda.Caption = "Busqueda de Conceptos"
            
            frmBusqueda.tdbgListado.Columns(0).Caption = "Codigo"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Rub_cCodigo"
            frmBusqueda.tdbgListado.Columns(0).Width = 1000
            frmBusqueda.tdbgListado.Columns(1).Caption = "Descripción de Concepto"
            frmBusqueda.tdbgListado.Columns(1).DataField = "Rub_cDescripcion"
            frmBusqueda.tdbgListado.Columns(1).Width = 5000
            frmBusqueda.tdbgListado.Columns(2).Caption = "Cuenta Contable"
            frmBusqueda.tdbgListado.Columns(2).DataField = "Rub_cCuentaContable"
            frmBusqueda.tdbgListado.Columns(2).Width = 0
            frmBusqueda.tdbgListado.Columns(2).Visible = False
            frmBusqueda.tdbgListado.Columns(3).Caption = "Tipo Entidad"
            frmBusqueda.tdbgListado.Columns(3).DataField = "Ten_cTipoEntidad"
            frmBusqueda.tdbgListado.Columns(3).Width = 0
            frmBusqueda.tdbgListado.Columns(3).Visible = False
            
            frmBusqueda.tdbgListado.Columns(4).Caption = "Documento"
            frmBusqueda.tdbgListado.Columns(4).DataField = "Pla_cDocumento"
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            
            frmBusqueda.tdbgListado.Columns(5).Caption = "Centro Costo"
            frmBusqueda.tdbgListado.Columns(5).DataField = "Pla_cCosCodigo"
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen Conceptos."
        End If
        BuscarConcepto = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function BuscaNombreConcepto(ByVal cTipo As String, ByVal cCodigo As String, ByRef bExiste As Boolean)
Dim gtxtSQL As String
Dim rsCta As New ADODB.Recordset
    BuscaNombreConcepto = ""
    cCodigo = CE(cCodigo)
    If cCodigo <> "" Then
        
        gtxtSQL = "spCn_GrabaCuentaRubro 'BUSCARREGISTRO', '" & gsEmpresa & "','" & cTipo & "','" & cCodigo & "','','','','','" & gsAnio & "'"
        Set rsCta = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(rsCta) > 0 Then
            BuscaNombreConcepto = CE(rsCta!Rub_cDescripcion)
            bExiste = True
        Else
            bExiste = False
        End If
        Call CerrarRecordSet(rsCta)
    End If
End Function

Public Function BuscarAsientoTipo(ByRef sTipoLibro As String, ByRef oFormCall As Form, Optional ByVal sCadBusq As String = "") As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarAsientoTipo = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then
        gtxtSQL = "spCn_GrabaTipoAsiento 'SEL_ALL_LIB', '" & gsEmpresa & "', '" & gsAnio & "', '" & sTipoLibro & "', '', 0, '', '', '', 0, '', '' "
        
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            frmBusqueda.Caption = "Busqueda de Tipo de Provisiones"
            frmBusqueda.tdbgListado.Columns(0).Caption = "Codigo"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Asl_cOperacion"
            frmBusqueda.tdbgListado.Columns(0).Width = 1000
            
            frmBusqueda.tdbgListado.Columns(1).Caption = "Descripción "
            frmBusqueda.tdbgListado.Columns(1).DataField = "Asl_cDescripcion"
            frmBusqueda.tdbgListado.Columns(1).Width = 5000
            
            frmBusqueda.tdbgListado.Columns(2).Caption = ""
            frmBusqueda.tdbgListado.Columns(2).DataField = "Lib_cTipoLibro"
            frmBusqueda.tdbgListado.Columns(2).Width = 0
            frmBusqueda.tdbgListado.Columns(2).Visible = False
            
            frmBusqueda.tdbgListado.Columns(3).Caption = ""
            frmBusqueda.tdbgListado.Columns(3).DataField = "Pla_cCentroCosto"
            frmBusqueda.tdbgListado.Columns(3).Width = 0
            frmBusqueda.tdbgListado.Columns(3).Visible = False
            
            frmBusqueda.tdbgListado.Columns(4).Caption = ""
            frmBusqueda.tdbgListado.Columns(4).DataField = "Ten_cTipoEntidad"
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            
            frmBusqueda.tdbgListado.Columns(5).Caption = ""
            frmBusqueda.tdbgListado.Columns(5).DataField = "Ten_cNombreEntidad"
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            
            
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen tipos de provisiones."
        End If
        BuscarAsientoTipo = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function BuscaNombreAsientoTipo(ByRef sTipoLibro As String, ByVal Codigo As String) As String
    Dim gtxtSQL As String
    Dim rsEnt As New ADODB.Recordset

    BuscaNombreAsientoTipo = ""
    gtxtSQL = "spCn_GrabaTipoAsiento 'SEL_REG', '" & gsEmpresa & "', '" & gsAnio & "', '" & sTipoLibro & "', '" & Codigo & "', 0, '', '', '', 0, '', '' "
    Set rsEnt = fRetornaRS(gtxtSQL)
        
    If GetRsRecordCount(rsEnt) > 0 Then
        BuscaNombreAsientoTipo = CE(rsEnt!Asl_cDescripcion)
    Else
        Mensajes "Código de tipo de provision no existe"
    End If
    Call CerrarRecordSet(rsEnt)
End Function

Public Function BuscarCentroCosto(ByRef oFormCall As Form, Optional ByVal sCadBusq As String = "") As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarCentroCosto = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then
        '***** Centros de Costo *****
        Dim sNivel As String
        sNivel = CE(fRetornaValor("spCNT_CONFIG_LIBROS 'BUSCARNIVEL','" & gsEmpresa & "','','','','','','','','',0,'','','','','','','','','','','','','','','','" & gsAnio & "'"))
        '----------------------------
        gtxtSQL = "spCNT_CENTRO_COSTO 'BUSAR_CC_F1GROUP','" & gsEmpresa & "','" & gsAnio & "'"
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            frmBusqueda.Caption = "Busqueda de Centros de Costo"
            
            frmBusqueda.tdbgListado.Columns(0).Caption = "Codigo"
            frmBusqueda.tdbgListado.Columns(0).DataField = "CodigoCCN"
            frmBusqueda.tdbgListado.Columns(0).Width = 1000
            
            frmBusqueda.tdbgListado.Columns(1).Caption = "Descripción"
            frmBusqueda.tdbgListado.Columns(1).DataField = "DescripCCN"
            frmBusqueda.tdbgListado.Columns(1).Width = 4000
            
            frmBusqueda.tdbgListado.Columns(2).Caption = "3er Nivel"
            frmBusqueda.tdbgListado.Columns(2).DataField = "DescripCC3"
            frmBusqueda.tdbgListado.Columns(2).Width = 2000
            frmBusqueda.tdbgListado.Columns(2).Merge = True
            
            frmBusqueda.tdbgListado.Columns(3).Caption = "2do Nivel"
            frmBusqueda.tdbgListado.Columns(3).DataField = "DescripCC2"
            frmBusqueda.tdbgListado.Columns(3).Width = 1800
            frmBusqueda.tdbgListado.Columns(3).Merge = True
            frmBusqueda.tdbgListado.Columns(3).Visible = True
            
            frmBusqueda.tdbgListado.Columns(4).Caption = "1er Nivel"
            frmBusqueda.tdbgListado.Columns(4).DataField = "DescripCC1"
            frmBusqueda.tdbgListado.Columns(4).Width = 1800
            frmBusqueda.tdbgListado.Columns(4).Merge = True
            frmBusqueda.tdbgListado.Columns(4).Visible = True
            
            frmBusqueda.tdbgListado.Columns(5).Caption = ""
            frmBusqueda.tdbgListado.Columns(5).DataField = ""
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False

            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            
            If sNivel = "T" Or sNivel = "C" Then
                frmBusqueda.Width = 11445
                frmBusqueda.tdbgListado.Width = 11445 - 50
            End If
            
            
            If sNivel = "P" Then
                frmBusqueda.tdbgListado.Columns(2).Visible = False
                frmBusqueda.tdbgListado.Columns(3).Visible = False
                frmBusqueda.tdbgListado.Columns(4).Visible = False
                frmBusqueda.tdbgListado.Columns(1).Caption = "1er Nivel"
            End If
            
            If sNivel = "S" Then
                frmBusqueda.tdbgListado.Columns(2).Visible = False
                frmBusqueda.tdbgListado.Columns(3).Visible = False
                frmBusqueda.tdbgListado.Columns(1).Caption = "2do Nivel"
            End If
            
            If sNivel = "T" Then
                frmBusqueda.tdbgListado.Columns(2).Visible = False
                frmBusqueda.tdbgListado.Columns(1).Caption = "3er Nivel"
            End If
            
            If sNivel = "C" Then
                frmBusqueda.tdbgListado.Columns(1).Caption = "4to Nivel"
            End If
            
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen Centros de Costo."
        End If
        BuscarCentroCosto = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function BuscaNombreEntidad(ByVal tipoEnt As String, ByVal Codigo As String) As String
    Dim gtxtSQL As String
    Dim rsEnt As New ADODB.Recordset

    BuscaNombreEntidad = ""
    gtxtSQL = "spCn_GrabaEntidad 'SEL_REG', '" & gsEmpresa & "', '" & Codigo & "', '" & tipoEnt & "'"
    Set rsEnt = fRetornaRS(gtxtSQL)
        
    If GetRsRecordCount(rsEnt) > 0 Then
        BuscaNombreEntidad = CE(rsEnt!Ent_cPersona)
        gsCampo3 = CE(rsEnt!Ent_cTipoDoc)
        gsCampo4 = CE(rsEnt!Ent_nRuc)
        gsCampo5 = CE(rsEnt!Ent_cDireccion)
    Else
        Mensajes "Código de entidad no existe"
    End If
    Call CerrarRecordSet(rsEnt)
End Function


Public Function BuscarTipoDoc(ByRef oFormCall As Form, Optional ByVal sCadBusq As String = "") As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarTipoDoc = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then
        '***** Centros de Costo *****
        gtxtSQL = "spCn_GrabaTipoDocumento 'SEL_ALL', '" & gsEmpresa & "'"
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            frmBusqueda.Caption = "Busqueda de Tipos de documentos"
            frmBusqueda.tdbgListado.Columns(0).Caption = "Codigo"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Tdo_cCodigo"
            frmBusqueda.tdbgListado.Columns(0).Width = 1000
            frmBusqueda.tdbgListado.Columns(1).Caption = "Descripción del Tipo de Documento"
            frmBusqueda.tdbgListado.Columns(1).DataField = "Tdo_cNombreLargo"
            frmBusqueda.tdbgListado.Columns(1).Width = 5000
            frmBusqueda.tdbgListado.Columns(2).Caption = ""
            frmBusqueda.tdbgListado.Columns(2).DataField = ""
            frmBusqueda.tdbgListado.Columns(2).Width = 0
            frmBusqueda.tdbgListado.Columns(2).Visible = False
            frmBusqueda.tdbgListado.Columns(3).Caption = ""
            frmBusqueda.tdbgListado.Columns(3).DataField = ""
            frmBusqueda.tdbgListado.Columns(3).Width = 0
            frmBusqueda.tdbgListado.Columns(3).Visible = False
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen Tipos de Documentos."
        End If
        BuscarTipoDoc = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function BuscaNombreTipDoc(ByVal cCodigo As String, ByRef bExiste As Boolean)
Dim gtxtSQL As String
Dim rsReg As New ADODB.Recordset
    BuscaNombreTipDoc = ""
    cCodigo = CE(cCodigo)
    If cCodigo <> "" Then
        gtxtSQL = "spCn_GrabaTipoDocumento 'SEL_REG', '" & gsEmpresa & "', '" & cCodigo & "'"
        Set rsReg = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(rsReg) > 0 Then
            BuscaNombreTipDoc = CE(rsReg!Tdo_cNombreLargo)
            bExiste = True
        Else
            bExiste = False
        End If
        Call CerrarRecordSet(rsReg)
    End If
End Function


Public Function BuscarVoucher(ByRef oFormCall As Form, ByVal sPeriodoIni, ByVal sPeriodoFin, ByVal sLibro) As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarVoucher = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then
        '***** Centros de Costo *****
        gtxtSQL = "select ase_nvoucher , dbo.fnombremes(per_cperiodo) as nombremes from cnc_asiento_voucher where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
                  "per_cperiodo>='" & sPeriodoIni & "' and per_cperiodo<='" & sPeriodoFin & "' and lib_ctipolibro='" & sLibro & "' and ase_cdeleted<>'*' Order by ase_nvoucher"
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            frmBusqueda.Caption = "Busqueda de Vouchers"
            frmBusqueda.tdbgListado.Columns(0).Caption = "Codigo de voucher"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Ase_nVoucher"
            frmBusqueda.tdbgListado.Columns(0).Width = 3500
            frmBusqueda.tdbgListado.Splits(0).Columns(0).Alignment = dbgCenter
            frmBusqueda.tdbgListado.Columns(1).Caption = "Nombre del periodo"
            frmBusqueda.tdbgListado.Columns(1).DataField = "nombremes"
            frmBusqueda.tdbgListado.Columns(1).Width = 1500
            frmBusqueda.tdbgListado.Splits(0).Columns(1).Merge = True

            frmBusqueda.tdbgListado.Columns(2).Caption = ""
            frmBusqueda.tdbgListado.Columns(2).DataField = ""
            frmBusqueda.tdbgListado.Columns(2).Width = 0
            frmBusqueda.tdbgListado.Columns(2).Visible = False
            frmBusqueda.tdbgListado.Columns(3).Caption = ""
            frmBusqueda.tdbgListado.Columns(3).DataField = ""
            frmBusqueda.tdbgListado.Columns(3).Width = 0
            frmBusqueda.tdbgListado.Columns(3).Visible = False
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen vouchers para este periodo y libro seleccionado."
        End If
        BuscarVoucher = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function BuscaConceptosValores(ByVal cLibro As String, ByRef cCodigo As String, ByRef cDescripcion As String) As Boolean
Dim gtxtSQL As String
Dim rsReg As New ADODB.Recordset
    
    If CE(cLibro) <> "" Then
        gtxtSQL = "select Asl_cCodigo , Asl_cDescripcion  from CNT_CONCEPTO_LIBRO where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
                  "lib_ctipolibro='" & cLibro & "' and isnull(Asl_cDefecto,'')='1' "
        
        Set rsReg = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(rsReg) > 0 Then
            cCodigo = CE(rsReg!Asl_cCodigo)
            cDescripcion = CE(rsReg!Asl_cDescripcion)
            
            bExiste = True
        Else
            
            cCodigo = ""
            cDescripcion = ""
            
            bExiste = False
        End If
        Call CerrarRecordSet(rsReg)
    End If
    
    BuscaConceptosValores = bExiste
End Function


Public Function BuscarConceptos(ByRef oFormCall As Form, ByVal sLibro) As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarConceptos = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then

        gtxtSQL = "select Asl_cCodigo , Asl_cDescripcion  from CNT_CONCEPTO_LIBRO where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
                  "lib_ctipolibro='" & sLibro & "' "
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            frmBusqueda.Caption = "Busqueda de Operaciones para el Diario Simplificado"
            frmBusqueda.tdbgListado.Columns(0).Caption = "Codigo"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Asl_cCodigo"
            frmBusqueda.tdbgListado.Columns(0).Width = 1000
            frmBusqueda.tdbgListado.Splits(0).Columns(0).Alignment = dbgCenter
            frmBusqueda.tdbgListado.Columns(1).Caption = "Descripcion de Operaciones"
            frmBusqueda.tdbgListado.Columns(1).DataField = "Asl_cDescripcion"
            frmBusqueda.tdbgListado.Columns(1).Width = 4000

            frmBusqueda.tdbgListado.Columns(2).Caption = ""
            frmBusqueda.tdbgListado.Columns(2).DataField = ""
            frmBusqueda.tdbgListado.Columns(2).Width = 0
            frmBusqueda.tdbgListado.Columns(2).Visible = False
            frmBusqueda.tdbgListado.Columns(3).Caption = ""
            frmBusqueda.tdbgListado.Columns(3).DataField = ""
            frmBusqueda.tdbgListado.Columns(3).Width = 0
            frmBusqueda.tdbgListado.Columns(3).Visible = False
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen operaciones asociadas al libro seleccionado."
        End If
        BuscarConceptos = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function


