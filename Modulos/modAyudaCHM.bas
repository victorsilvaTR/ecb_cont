Attribute VB_Name = "modAyudaCHM"
'******************************************************************************
'----- Modul - definition for HTMLHelp - (c) Ulrich Kulle
'----- 2002-08-26 Version 1.0 first release
'----- 2005-07-17 Version 1.1 updated for Pop-Up help
'******************************************************************************
'----- Portions of this code courtesy of David Liske.
'----- Thanks to David Liske, Don Lammers, Matthew Brown and Thomas Schulz
'------------------------------------------------------------------------------
Dim Result As Long
Public VarGsIndDS As Boolean

Type HH_IDPAIR
  dwControlId As Long
  dwTopicId As Long
End Type

'This array should contain the number of controls that have
'context-sensitive help, plus one more for a zero-terminating
'pair.

Public ids(2) As HH_IDPAIR

Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
                
Declare Function HTMLHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As String) As Long
         
Private Declare Function HtmlHelpSearch Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, dwData As HH_FTS_QUERY) As Long
         

Public Const HH_DISPLAY_TOPIC = &H0         ' select last opened tab, [display a specified topic]
Public Const HH_DISPLAY_TOC = &H1           ' select contents tab, [display a specified topic]
Public Const HH_DISPLAY_INDEX = &H2         ' select index tab and searches for a keyword
Public Const HH_DISPLAY_SEARCH = &H3        ' select search tab and perform a search
      
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or
  
Public Const HH_HELP_CONTEXT = &HF          ' display mapped numeric value in dwData
     
Private Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Private Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.


Public Type HH_FTS_QUERY                ' UDT for accessing the Search tab
  cbStruct          As Long             ' Sizeof structure in bytes.
  fUniCodeStrings   As Long             ' True if all strings are unicode.
  pszSearchQuery    As String           ' String containing the search query.
  iProximity        As Long             ' Word proximity.
  fStemmedSearch    As Long             ' True for StemmedSearch only.
  fTitleOnly        As Long             ' True for Title search only.
  fExecute          As Long             ' True to initiate the search.
  pszWindow         As String           ' Window to display in
End Type

Public Function HFile(ByVal i_HFile As Integer) As String
'----- Set the string variable to include the application path of helpfile
  Select Case i_HFile
  Case 1
    HFile = App.Path & "\Ayuda\EcbCont - Sistema Contable.chm"
  Case 2
'----- Place other Help file paths in successive case statements
    HFile = App.Path & "\Ayuda\Libros_Electronicos.chm"
  End Select
End Function

Private Sub MensajeAyuda(Param As Long)
    If Param = 0 Then Mensajes "No se encuentra el archivo de ayuda"
End Sub

Public Sub ShowContents(ByVal intHelpFile As Integer)
   Result = HtmlHelp(hwnd, HFile(intHelpFile), HH_DISPLAY_TOC, 0)
   MensajeAyuda Result
End Sub

Public Sub ShowIndex(ByVal intHelpFile As Integer)
   Result = HtmlHelp(hwnd, HFile(intHelpFile), HH_DISPLAY_INDEX, 0)
   MensajeAyuda Result
End Sub

Public Sub ShowTopic(ByVal intHelpFile As Integer, strTopic As String)
   Result = HTMLHelpTopic(hwnd, HFile(intHelpFile), HH_DISPLAY_TOPIC, strTopic)
   MensajeAyuda Result
End Sub

Public Sub ShowTopicID(ByVal intHelpFile As Integer, IdTopic As Long)
   Result = HtmlHelp(hwnd, HFile(intHelpFile), HH_HELP_CONTEXT, IdTopic)
   MensajeAyuda Result
End Sub
'------------------------------------------------------------------------------
'----- display the search tab
'----- bug: start searching with a string dosn't work
'------------------------------------------------------------------------------
Public Sub ShowSearch(ByVal intHelpFile As Integer)
Dim searchIt As HH_FTS_QUERY
  With searchIt
    .cbStruct = Len(searchIt)
    .fUniCodeStrings = 1&
    .pszSearchQuery = "foobar"
    .iProximity = 0&
    .fStemmedSearch = 0&
    .fTitleOnly = 1&
    .fExecute = 1&
    .pszWindow = ""
  End With
  Result = HtmlHelpSearch(0&, HFile(intHelpFile), HH_DISPLAY_SEARCH, searchIt)
  MensajeAyuda Result
End Sub

