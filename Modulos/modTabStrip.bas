Attribute VB_Name = "modTabStrip"
Option Explicit
' **************************************************************************
' User variables
' **************************************************************************

Public Const Version = "1.1"
Public Const VDate = "13-05-2005"

Private m_Tabstrip_hwnd As Long                            'Tabstrip handle

Dim m_Tabs_TextColor As Long                               ' Color of text of tabstrip tabs
Dim m_Tabstrip_BackColor As Long                           ' BackColor of Tabstrip
Dim m_Tabstrip_MainColor As Long                           ' Backcolor of main tabstrip

' **************************************************************************
' API to manage Windows
' **************************************************************************

'La fonction GetWindowLong recupere les attributs d'une fenêtre spécifique.
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nindex As Long) As Long

'La fonction SetWindowLong modifie les attributs d'une fenêtre spécifique.
'Elle insere une variable (32 bits) à une deplacement spécifique de la mémoire
'de Windows. Elle donne en retour la valeur qui y était stockée auparavent.
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nindex As Long, ByVal dwnewlong As Long) As Long

'CallWindowProc passe un message d'information à la proc d'une fenetre donnée.
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' **************************************************************************
' To subclass tabstrip control and its form parent
' **************************************************************************
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)

' **************************************************************************
' Graphical APIs
' **************************************************************************
Declare Function GetSysColorBrush Lib "user32" (ByVal nindex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Public Const ILD_TRANSPARENT = 1&

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
        
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Public Const TRANSPARENT = 1

Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = -4

Public Const TCIF_TEXT = &H1
Public Const TCIF_IMAGE = &H2

Public Const TCM_FIRST = &H1300                            '// Tab control messages
Public Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)
'Public const TabCtrl_GetItemCount(hwnd) \
'    (int)SNDMSG((hwnd), TCM_GETITEMCOUNT, 0, 0L)

Public Const TCM_GETITEMA = (TCM_FIRST + 5)
Public Const TCM_GETITEMW = (TCM_FIRST + 60)
Public Const TCM_GETIMAGELIST = (TCM_FIRST + 2)
'private const TabCtrl_GetImageList(hwnd) \
'    (HIMAGELIST)SNDMSG((hwnd), TCM_GETIMAGELIST, 0, 0L)

Public Const TCM_SETIMAGELIST = (TCM_FIRST + 3)
'private const TabCtrl_SetImageList(hwnd, himl) \
'    (HIMAGELIST)SNDMSG((hwnd), TCM_SETIMAGELIST, 0, (LPARAM)(UINT)(HIMAGELIST)(himl))

 Type TCITEM
    mask As Long
    ' #if (_WIN32_IE >= =&H0300)
    dwState As Long
    dwStateMask As Long
    ' #Else
    '    UINT lpReserved1;
    '    UINT lpReserved2;
    ' #End If
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

 Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type Size
    cx          As Long
    cy          As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemAction  As Long
    itemState   As Long
    hWndItem    As Long
    hDC         As Long
    rcItem      As RECT
    itemData    As Long
End Type

Public Const COLOR_WINDOW = 5
Public Const COLOR_HIGHLIGHT = 13

Public Const ODS_SELECTED = &H1
Public Const ODT_TAB = 101

'Window Message to be intercepted
Public Const WM_DRAWITEM = &H2B
Public Const WM_PRINTCLIENT = &H318

'Tab Style to allow user drawning
Public Const TCS_OWNERDRAWFIXED = &H2000                   'must be set

 Function GetTabImage(ByVal vKey As Variant) As Long

    Dim lIndex As Long
    Dim tTI As TCITEM
    Dim lR As Long

    'Find Icon to display
    'Works only with COMCTL32
    lIndex = vKey - 1
    If (lIndex > -1) Then
        tTI.mask = TCIF_IMAGE
        lR = SendMessage(CtlHwnd, TCM_GETITEMA, lIndex, tTI)
        If (lR <> 0) Then
            GetTabImage = tTI.iImage
        Else
            Debug.Print "Failed to get image for tab " & vKey
        End If
    End If
End Function
 Public Function GetTabText(ByVal vKey As Variant) As String

    Dim lIndex As Long
    Dim tTI As TCITEM
    Dim lR As Long
    Dim sText As String

    'Recupération du texte à afficher
    lIndex = vKey - 1

    tTI.cchTextMax = 255
    tTI.pszText = String$(255, 0)
    tTI.mask = TCIF_TEXT
    lR = SendMessage(CtlHwnd, TCM_GETITEMA, lIndex, tTI)
    If (lR <> 0) Then
        sText = tTI.pszText
        lR = InStr(sText, Chr$(0))
        If (lR <> 0) Then
            GetTabText = Left$(sText, lR - 1)
        Else
            GetTabText = sText
        End If
    Else
        Debug.Print "TabIndex " & vKey & " does not exist"
    End If
End Function

Public Sub Hook(ByRef aTabstrip As Object)
    Dim origProc As Long                                   ' Original process address
    Dim aStyle As Long                                     ' Original style for the tabstrip

    m_Tabs_TextColor = &HB15C07
    m_Tabstrip_BackColor = RGB(255, 255, 255)
    m_Tabstrip_MainColor = &HB15C07

    'Save the handle of tabstrip
    CtlHwnd = aTabstrip.hwnd

    ' Set the OwnerDrawn style

    ' get the original style
    aStyle = GetWindowLong(aTabstrip.hwnd, GWL_STYLE)
    ' add in the ownerdrawn style
    aStyle = aStyle Or TCS_OWNERDRAWFIXED
    ' replace the style with our "ownerdrawn" one
    SetWindowLong aTabstrip.hwnd, GWL_STYLE, aStyle

    ' Subclass the tabstrip - used to change the background color of the tabstrip
    'Credit Elite  VB - Garrett Sever -
    ' Redirect our messages to the function "TabStripProc"
    ' Swap de l'adresse de la routine standard Windows par l'adresse de notre routine WindowProc.
    ' origProc récupère l'adresse de la routine standard.
    origProc = SetWindowLong(aTabstrip.hwnd, GWL_WNDPROC, AddressOf TabStripProc)
    ' Store the original process address against the tabstrip's handle
    SetProp aTabstrip.hwnd, "OrigTabStripProc", origProc
    ' Store a "tabheight" value so we can make the tab background a different color
    ' than the rest of the tabstrip control (want the tab backgrounds, i.e. the
    ' parts that aren't really tabs to be the same color as the form)
    
    'SetProp aTabstrip.hwnd, "TabHeight", (aTabstrip.Height - aTabstrip.ClientHeight) / Screen.TwipsPerPixelY
   
    ' subclass the parent - used to capture the WM_DRAWITEM message

    ' Redirect our messages for the tabstrip's parent to "TabStripOwnerProc"
    ' Swap de l'adresse de la routine standard Windows par l'adresse de notre routine WindowProc.
    ' origProc récupère l'adresse de la routine standard.
    origProc = SetWindowLong(GetParent(aTabstrip.hwnd), GWL_WNDPROC, AddressOf WindowProc)
    ' Store the original window process's address against its handle (we're assuming
    '  only in name that its owner is a form - it could be a picturebox or something else)
    SetProp GetParent(aTabstrip.hwnd), "OrigWindowProc", origProc
    ' Store a pointer to the tabstrip's form so we can implement our slimy hack
    '  "safe" subclassing method
    SetProp GetParent(aTabstrip.hwnd), "FormHwnd", ObjPtr(aTabstrip.Parent)
End Sub

 Public Sub Unhook(ByRef aTabstrip As Object)

    Dim origProc As Long

    'Unhook the form

    ' Get the original process address for the tabstrip's parent
    origProc = GetProp(GetParent(aTabstrip.hwnd), "OrigWindowProc")
    ' Redirect all messages back to this parent
    SetWindowLong GetParent(aTabstrip.hwnd), GWL_WNDPROC, origProc
    ' Remove the entries from windows' internal database
    RemoveProp GetParent(aTabstrip.hwnd), "OrigWindowProc"
    RemoveProp GetParent(aTabstrip.hwnd), "FormHwnd"

    'unhook the tabstrip control

    ' Get the original process for the tabstrip
    origProc = GetProp(aTabstrip.hwnd, "OrigTabStripProc")
    ' Redirect all messages back to its original process
    SetWindowLong aTabstrip.hwnd, GWL_WNDPROC, origProc
    ' Remove the entries from windows' internal database
    RemoveProp aTabstrip.hwnd, "OrigTabStripProc"
    RemoveProp aTabstrip.hwnd, "TabHeight"

End Sub

 Public Function WindowProc(ByVal hW As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tDis As DRAWITEMSTRUCT
    Dim bSelected As Boolean
    Dim lTab As Long
    Dim oldWndProc As Long                                 ' original windows process for the form

    ' Get the original
    oldWndProc = GetProp(hW, "OrigWindowProc")

    ' This is the callback procedure used when a message is received by this form.
    ' The desired message is processed and all others are passed back to the
    ' original procedure associated with the form.

    Select Case uMsg
        Case WM_DRAWITEM
            'Debug.Print "DrawItem: "; "hw:"; hW; " uMsg: "; "&H"; Hex(uMsg), "wp:"; wParam, "lp: "; lParam
            If wParam = 0 Then
                'Copy DRAWINFOSTRUCT data to local variable
                CopyMemory tDis, ByVal lParam, Len(tDis)
                If tDis.CtlType = ODT_TAB Then             'Check if Tab Control
                    'Debug.Print tDis.CtlID, tDis.CtlType, tDis.itemID, tDis.itemData, tDis.itemAction, tDis.itemState
                    lTab = tDis.itemID
                    'ODS_SELECTED   This bit is set if the item’s status is selected.
                    bSelected = ((tDis.itemState And ODS_SELECTED) = ODS_SELECTED)
                    DrawItem lTab, tDis.hDC, bSelected, tDis.rcItem
                    Exit Function
                End If
            End If
    End Select
    WindowProc = CallWindowProc(oldWndProc, hW, uMsg, wParam, lParam)
End Function


Public Sub DrawItem(ByVal lTab As Long, ByVal lhDC As Long, ByVal bSelected As Boolean, ByRef tR As RECT)

    Dim hBr As Long
    Dim cx As Long, cy As Long
    Dim lX As Long, lY As Long
    Dim lImage As Long
    Dim tTI As TCITEM
    Dim Ihwnd As Long
    Dim m_hIml As Long                                     'Handle de l'imagelist

    'Remplissage des onglets
    ' Fill back color
    'Choose the different option if you need - not done in this sample
    If bSelected Then
        'hBr = GetSysColorBrush(vbButtonFace And &H1F&)
        'Add your background color here
        hBr = CreateSolidBrush(m_Tabstrip_BackColor)
    Else
        'hBr = GetSysColorBrush(vbButtonShadow And &H1F&)
        hBr = CreateSolidBrush(m_Tabstrip_BackColor)
    End If
    
    FillRect lhDC, tR, hBr
    DeleteObject hBr

    ' Icon: Looking for icon to be displayed
    lImage = GetTabImage(lTab + 1)        'Recherche de l'icone à afficher/
    If lImage > -1 Then
        ' draw the icon
        'Doesn't works with the MSCOMCTL ( VB6 SP4 )
        'Recupere le handle de l'image list (ihwnd et tTI inutiles )
        m_hIml = SendMessage(CtlHwnd, TCM_GETIMAGELIST, Ihwnd, tTI)
        If Not m_hIml = 0 Then
            ImageList_GetIconSize m_hIml, cx, cy
            If bSelected Then
                lX = tR.Left + 6
            Else
                lX = tR.Left + 2
            End If
            lY = tR.Top + (tR.Bottom - tR.Top - cy) \ 2
            ImageList_Draw m_hIml, lImage, lhDC, lX, lY, ILD_TRANSPARENT
            tR.Left = lX + cx + 1
        End If
    End If

    ' Looking for  a text to be displayed
    Dim MsgTab As String
    SetBkMode lhDC, TRANSPARENT
    MsgTab = GetTabText(lTab + 1)            'Recherche du texte à afficher
    If m_Tabs_TextColor = 0 Then m_Tabs_TextColor = vbBlack  'MR 13-mai-05
    SetTextColor lhDC, m_Tabs_TextColor       'Text color 'MR 12-mai-05
    DrawText lhDC, MsgTab, -1, tR, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER

End Sub

Private Function TabStripProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim oldWndProc      As Long                            ' Original process address for the tabstrip
    Dim aRect           As RECT                            ' Rectangle structure for drawing and filling with colors
    Dim abrush          As Long                            ' brush object for filling with colors
    Dim aTabHeight      As Long                            ' height of the tab items


    'Piece of code from Vb Elite credits to  Garrett Sever
    'updated 2005 by M.Rodenas

    'Default value
    'Tabstrip backcolor = Form.backcolor
    'Tabstrip main backcolor = form.backcolor

    ' Get the original tabstrip process address
    oldWndProc = GetProp(hwnd, "OrigTabStripProc")

    ' If its going to "print" the tabstrip shape on the client area, we
    '  need to paint the background colors first
    If wMsg = WM_PRINTCLIENT Then
        ' Retrieve the height of the tabs as stored when the tabstrip was first subclassed
        aTabHeight = GetProp(hwnd, "TabHeight")
        ' Get the dimensions for the tabstrip's area
        GetClientRect hwnd, aRect

        'Init default value if no value
        If m_Tabstrip_MainColor = 0 Then m_Tabstrip_MainColor = GetBkColor(GetDC(GetParent(hwnd)))
        ' Create a brush with our "mainTabstrip" background color. this is the
        '  color that we're filling the main body of the tabstrip with
        abrush = CreateSolidBrush(m_Tabstrip_MainColor)

        ' Adjust our fill area to down to the top of the tabs
        aRect.Top = aRect.Top + aTabHeight + 1             'MR 12-mai-05
        ' Fill the main area of the tabstrip
        FillRect wParam, aRect, abrush
        ' clean up the brush object
        DeleteObject abrush

        ' Set the fill area to only the tab area. This doesn't really fill the
        '  tabs as much as it makes the area around the tabs the right color.
        '  that way we can match it to the backcolor of our form.
        aRect.Top = 0                                      'MR 12-mai-05
        aRect.Left = 0                                     'MR 12-mai-05
        'arect.Right=
        aRect.Bottom = aTabHeight                          'MR 12-mai-05

        'Init default value to the form.backcolor value
        If m_Tabstrip_BackColor = 0 Then m_Tabstrip_BackColor = GetBkColor(GetDC(GetParent(hwnd)))

        ' Create the brush the same color as the backcolor of our form
        abrush = CreateSolidBrush(m_Tabstrip_BackColor)
        ' Fill our selected area
        FillRect wParam, aRect, abrush
        ' clean up the brush object
        DeleteObject abrush
    End If

    ' Invoke whatever the default process for our tabstrip, including the
    '  WM_PRINTCLIENT message that will draw the tabstrip's border
    TabStripProc = CallWindowProc(oldWndProc, hwnd, wMsg, wParam, lParam)

End Function

Private Property Get CtlHwnd() As Long

    CtlHwnd = m_Tabstrip_hwnd
End Property

Private Property Let CtlHwnd(ByVal vNewValue As Long)

    m_Tabstrip_hwnd = vNewValue
End Property

Public Property Get Tabs_TextColor() As Long

    Tabs_TextColor = m_Tabs_TextColor
End Property

Public Property Let Tabs_TextColor(ByVal vNewValue As Long)

    m_Tabs_TextColor = vNewValue
End Property

Public Property Get Tabstrip_BackColor() As Long

    Tabstrip_BackColor = m_Tabstrip_BackColor
End Property

Public Property Let Tabstrip_BackColor(ByVal vNewValue As Long)

    m_Tabstrip_BackColor = vNewValue
End Property

Public Property Get Tabstrip_MainColor() As Long

    Tabstrip_MainColor = m_Tabstrip_MainColor
End Property

Public Property Let Tabstrip_MainColor(ByVal vNewValue As Long)

    m_Tabstrip_MainColor = vNewValue
End Property

