VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECActivacionValidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Sub Main()

    Dim todayDate As Date
    Dim exitoLoad As Boolean
    Dim hasConnection As Boolean
    Dim openLicencias As Boolean
    Dim openUpdates As Boolean
    Dim ActualizacionDisponible As Boolean
    Dim ActivationDt As Variant
    Set objActivacion = New ECActivacion
    todayDate = Now
     
    If (objActivacion.ExistDataBase = True) Then 'Base datos Local existe
        
        'Trae informacion de la Base datos Local
        Set ActivationDt = objActivacion.GetActivacion()
  
        If Not ActivationDt Is Nothing Then
                Dim ActivationDtAux As ECCOMActivacion.ActivationData
                
                'Validar si Webservice tiene conexion
                hasConnection = objActivacion.WSHasConnection()
                 
                If hasConnection = True Then
                    'Traer infomacion actualizada de la activacion con el WebService
                    Set ActivationDtAux = objActivacion.GetWSActivation(ActivationDt.Tax_Number, ActivationDt.Host)
                    
                    'actualiza en Base datos Local lo traido por el WebService
                     objActivacion.LoadActivacion ActivationDtAux
                     
                    'Trae informacion de la Base datos Local
                     Set ActivationDt = objActivacion.GetActivacion()
                End If
                
                If Not ActivationDt Is Nothing Then
                       
                    'El producto esta Desactivado, entonces inicia aplicativo Licenciamiento
                    If (ActivationDt.Estado_Id = 6 Or ActivationDt.Estado_Id = 7) Then
                        MsgBox "No existe producto Activado o Vigente. Se abrira aplicativo Licencias para continuar con la Activación. Despues de su activación debera volver a abrir el aplicativo ECB-CONT para que los cambios tengan efecto!", vbCritical
                        openLicencias = objActivacion.OpenAppLicencias()
                        If openLicencias = False Then
                            MsgBox "No se pudo abrir el aplicativo de Licenciamiento, si el problema persiste comuniquese con mesa de ayuda!", vbCritical
                        End If
                        
                    'Vencio la fecha de vigencia del producto, entonces inicia aplicativo Licenciamiento
                    ElseIf (DateValue(Format(todayDate, "dd/MM/yyyy")) > DateValue(Format(ActivationDt.FECHA_FIN, "dd/MM/yyyy"))) Then
                        MsgBox "Su producto llego a su fecha limite de vigencia, para continuar su uso debera Activar con su ultima compra aquirirda. Se abrira aplicativo Licencias. Despues de su activación debera volver a abrir el aplicativo ECB-CONT para que los cambios tengan efecto!", vbCritical
                        openLicencias = objActivacion.OpenAppLicencias()
                        If openLicencias = False Then
                            MsgBox "No se pudo abrir el aplicativo de Licenciamiento, si el problema persiste comuniquese con mesa de ayuda!", vbCritical
                        End If
                        
                    Else
                        ''Logica de Actualizaciones automaticas
                        'ActualizacionDisponible = objActivacion.ActualizacionDisponible(ActivationDt.Tax_Number, gsVersion, gsNumSubida)
                        
                        'If ActualizacionDisponible = True Then
                         '   openUpdates = objActivacion.OpenAppUpdates()
                        'Else
                            frmPrcIngresoSistema.Show
                        'End If
                    
                    End If
                    
                Else
                    'El producto no existe o No hay Activaciones activas por lo traido en el WebService posiblemente poque hubo una [Desactivacion Automatica] o el registro esta [Suspendido], entonces entonces inicia aplicativo Licenciamiento
                    MsgBox "No existe producto Activado o Vigente. Se abrira aplicativo Licencias para continuar con la Activación. Despues de su activación debera volver a abrir el aplicativo ECB-CONT para que los cambios tengan efecto!", vbCritical
                    openLicencias = objActivacion.OpenAppLicencias()
                    If openLicencias = False Then
                        MsgBox "No se pudo abrir el aplicativo de Licenciamiento, si el problema persiste comuniquese con mesa de ayuda!", vbCritical
                    End If
                End If
                

         Else
           'El producto no existe, entonces entonces inicia aplicativo Licenciamiento
            MsgBox "No existe producto Activado o Vigente. Se abrira aplicativo Licencias para continuar con la Activación. Despues de su activación debera volver a abrir el aplicativo ECB-CONT para que los cambios tengan efecto!", vbCritical
            openLicencias = objActivacion.OpenAppLicencias()
            If openLicencias = False Then
                MsgBox "No se pudo abrir el aplicativo de Licenciamiento, si el problema persiste comuniquese con mesa de ayuda!", vbCritical
            End If
         End If
            
    Else
      'El producto no existe, entonces inicia aplicativo Licenciamiento
       MsgBox "No existe producto Activado o Vigente. Se abrira aplicativo Licencias para continuar con la Activación. Despues de su activación debera volver a abrir el aplicativo ECB-CONT para que los cambios tengan efecto!", vbCritical
       openLicencias = objActivacion.OpenAppLicencias()
       If openLicencias = False Then
            MsgBox "No se pudo abrir el aplicativo de Licenciamiento, si el problema persiste comuniquese con mesa de ayuda!", vbCritical
       End If
    End If

End Sub

    
