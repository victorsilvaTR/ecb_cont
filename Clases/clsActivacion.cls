VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActivacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public Tax_Number As String
Public Opportunity_ID As String
Public Product_Code As String
Public Product_Name As String
Public FechaInicio As String
Public FechaFin As String
Public Fecha_Activacion As String
Public Fecha_Desactivacion As String

Public Codigo1 As String
Public Codigo2 As String
Public Codigo_Desact As String
Public Host As String
Public Sist_op_desc As String
Public Sist_op_version As String
Public Suscribe As Boolean

Public Estado_Id As Long
