VERSION 5.00
Begin VB.UserControl jeffLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   780
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   2835
   End
End
Attribute VB_Name = "jeffLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#If bStandAloneControls Then
    
Private Declare Function GetTextFace _
                                Lib "gdi32" _
                                Alias "GetTextFaceA" ( _
                            ByVal hdc As Long, _
                            ByVal nCount As Long, _
                            ByVal lpFacename As String) _
                        As Long

Private Declare Function GetTextMetrics _
                                Lib "gdi32" _
                                Alias "GetTextMetricsA" ( _
                            ByVal hdc As Long, _
                            lpMetrics As TEXTMETRIC) _
                        As Long

Private Declare Function GetDeviceCaps _
                                Lib "gdi32" ( _
                            ByVal hdc As Long, _
                            ByVal nIndex As Long) _
                        As Long

Private Declare Function MulDiv _
                                Lib "kernel32" ( _
                            ByVal nNumber As Long, _
                            ByVal nNumerator As Long, _
                            ByVal nDenominator As Long) _
                        As Long

Private Declare Function SelectObject _
                                Lib "gdi32" ( _
                            ByVal hdc As Long, _
                            ByVal hObject As Long) _
                        As Long
Private Declare Function CreateFontIndirect _
                                Lib "gdi32" _
                                Alias "CreateFontIndirectA" ( _
                            lpLogFont As LOGFONT) _
                        As Long

Private Declare Function SetTextColor _
                                Lib "gdi32" ( _
                            ByVal hdc As Long, _
                            ByVal crColor As Long) _
                        As Long

Private Declare Function GetTextColor _
                                Lib "gdi32" ( _
                            ByVal hdc As Long) _
                        As Long

Private Declare Function DrawText _
                                Lib "user32" _
                                Alias "DrawTextA" ( _
                            ByVal hdc As Long, _
                            ByVal lpStr As String, _
                            ByVal nCount As Long, _
                            lpRect As Rect, _
                            ByVal wFormat As Long) _
                        As Long

Private Declare Function GetDesktopWindow _
                                Lib "user32" () _
                        As Long

Private Declare Function ShellExecute _
                                    Lib "Shell32.dll" _
                                    Alias "ShellExecuteA" ( _
                                ByVal hWnd As Long, _
                                ByVal lpOperation As String, _
                                ByVal lpFile As String, _
                                ByVal lpParameters As String, _
                                ByVal lpDirectory As String, _
                                ByVal nShowCmd As enumShowWindow) _
                            As Long


Private Const LF_FACESIZE = 32

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Enum enumDeviceCaps
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
    BITSPIXEL = 12         '  Number of bits per pixel
End Enum

Private Enum enumTMPF
    TMPF_DEVICE = &H8
    TMPF_FIXED_PITCH = &H1
    TMPF_TRUETYPE = &H4
    TMPF_VECTOR = &H2
End Enum

Private Enum enumTextAlignment
    TA_BASELINE = 24
    TA_BOTTOM = 8
    TA_CENTER = 6
    TA_LEFT = 0
    TA_NOUPDATECP = 0
    TA_RIGHT = 2
    TA_TOP = 0
    TA_UPDATECP = 1
    TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
End Enum

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Enum enumDrawTextFormats
    [DT_BOTTOM] = &H8
    [DT_CALCRECT] = &H400
    [DT_CENTER] = &H1
    [DT_EXPANDTABS] = &H40
    [DT_EXTERNALLEADING] = &H200
    [DT_INTERNAL] = &H1000
    [DT_LEFT] = &H0
    [DT_NOCLIP] = &H100
    [DT_NOPREFIX] = &H800
    [DT_RIGHT] = &H2
    [DT_SINGLELINE] = &H20
    [DT_TABSTOP] = &H80
    [DT_TOP] = &H0
    [DT_VCENTER] = &H4
    [DT_WORDBREAK] = &H10
End Enum

Private Enum enumShellWindows
    SW_ERASE = &H4
    SW_HIDE = 0
    SW_INVALIDATE = &H2
    SW_MAX = 10
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
    SW_OTHERUNZOOM = 4
    SW_OTHERZOOM = 2
    SW_PARENTCLOSING = 1
    SW_PARENTOPENING = 3
    SW_RESTORE = 9
    SW_SCROLLCHILDREN = &H1
    SW_SHOW = 5
    SW_SHOWDEFAULT = 10
    SW_SHOWMAXIMIZED = 3
    SW_SHOWMINIMIZED = 2
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_SHOWNOACTIVATE = 4
    SW_SHOWRESTORE = 5
    SW_SHOWNORMAL = 1
End Enum

Private Enum enumShowWindow
    esw_SHOWDEFAULT = SW_SHOWDEFAULT
    esw_SHOWMAXIMIZED = SW_SHOWMAXIMIZED
    esw_SHOWMINIMIZED = SW_SHOWMINIMIZED
    esw_SHOWMINNOACTIVE = SW_SHOWMINNOACTIVE
    esw_SHOWNA = SW_SHOWNA
    esw_SHOWNOACTIVATE = SW_SHOWNOACTIVATE
    esw_SHOWRESTORE = SW_SHOWRESTORE
    esw_SHOWNORMAL = SW_SHOWNORMAL
    
End Enum

Private Const ERROR_BAD_FORMAT = 11&

Private Enum enumShellExecuteErrors
    SE_ERR_FNF = 2&
    SE_ERR_PNF = 3&
    SE_ERR_ACCESSDENIED = 5&
    SE_ERR_OOM = 8&
    SE_ERR_DLLNOTFOUND = 32&
    SE_ERR_SHARE = 26&
    SE_ERR_ASSOCINCOMPLETE = 27&
    SE_ERR_DDETIMEOUT = 28&
    SE_ERR_DDEFAIL = 29&
    SE_ERR_DDEBUSY = 30&
    SE_ERR_NOASSOC = 31&
    SE_ERROR_BAD_FORMAT = ERROR_BAD_FORMAT
End Enum

#End If  ' bStandAloneControls
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=






Public Enum enumLabelAppearance
    elaThreeD = 1
    elaFlat = 0
End Enum

Public Enum enumLabelBackstyle
    elbOpaque = 1
    elbTransparent = 0
End Enum

Public Enum enumLabelBorderstyle
    elrFixedSingle = 1
    elrNone = 0
End Enum




'Event Declarations:
Event Click() 'MappingInfo=lbl(0),lbl,0,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lbl(0),lbl,0,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lbl(0),lbl,0,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lbl(0),lbl,0,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lbl(0),lbl,0,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change() 'MappingInfo=lbl(0),lbl,0,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=lbl(0),lbl,0,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lbl(0),lbl,0,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=lbl(0),lbl,0,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=lbl(0),lbl,0,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=lbl(0),lbl,0,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=lbl(0),lbl,0,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."

Private mOffsetX As Double
Private mOffsetY As Double
Private mURL As String

Private bReadingProps As Boolean
Private bReadProps As Boolean

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#If bStandAloneControls Then

Public Function shellLaunchDoc(ByVal DocFileName As String, Optional ByVal PathName As String) As Boolean
    
    Dim lReturn As Long
    Dim lWinHandle As Long
    lWinHandle = GetDesktopWindow()
    If Trim(PathName) = "" Then
        PathName = CurDir
    End If
    
'    lReturn = ShellExecute(Me.hwnd, vbNullString, "assocapps.vbp", vbNullString, CurDir, 1)
    lReturn = ShellExecute(lWinHandle, "Open", DocFileName, "", PathName, SW_SHOWNORMAL)
    
    Dim sError As String
    If lReturn <= 32 Then
        'There was an error
        Select Case lReturn
            Case SE_ERR_FNF
                sError = "File not found"
            Case SE_ERR_PNF
                sError = "Path not found"
            Case SE_ERR_ACCESSDENIED
                sError = "Access denied"
            Case SE_ERR_OOM
                sError = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                sError = "DLL not found"
            Case SE_ERR_SHARE
                sError = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                sError = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                sError = "DDE Time out"
            Case SE_ERR_DDEFAIL
                sError = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                sError = "DDE busy"
            Case SE_ERR_NOASSOC
                sError = "No association for file extension"
            Case ERROR_BAD_FORMAT
                sError = "Invalid EXE file or error in EXE image"
            Case Else
                sError = "Unknown error"
        End Select
        MsgBox sError, vbCritical + vbOKOnly, "File Execute Error"
        shellLaunchDoc = False
    Else
        shellLaunchDoc = True
    End If
        
End Function

Public Function sNT( _
                        ByVal sString As String) _
                As String
                
    Dim iNullLoc As Integer
    iNullLoc = InStr(sString, Chr(0))
    If iNullLoc > 0 Then
        sNT = Left(sString, iNullLoc - 1)
    Else
        sNT = sString
    End If
End Function

Private Function pixelsX(ByVal TwipsIn As Long) As Integer
    pixelsX = TwipsIn / Screen.TwipsPerPixelX
End Function

Private Function pixelsY(ByVal TwipsIn As Long) As Integer
    pixelsY = TwipsIn / Screen.TwipsPerPixelY
End Function

Private Function twipsX( _
                        ByVal PixelsIn As Variant) _
                As Long
    twipsX = PixelsIn * Screen.TwipsPerPixelX
End Function
Private Function twipsY( _
                        ByVal PixelsIn As Variant) _
                As Long
    twipsY = PixelsIn * Screen.TwipsPerPixelY
End Function

Private Function dcSetFont( _
                        ByVal hdc As Long, _
                        ByVal oFont As StdFont, _
                        Optional ByVal lFontColor As OLE_COLOR = -99)
    
    Dim oLogFont As LOGFONT
    With oLogFont
        .lfCharSet = oFont.Charset
        Dim l As Long
        For l = 1 To Len(oFont.Name)
            .lfFaceName(l) = Asc(Mid(oFont.Name, l, 1))
        Next l
        .lfHeight = -MulDiv(GetDeviceCaps(hdc, LOGPIXELSY), oFont.Size, 72)
        .lfItalic = oFont.Italic
        .lfStrikeOut = oFont.Strikethrough
        .lfUnderline = oFont.Underline
        .lfWeight = oFont.Weight
        .lfQuality = 2
        
    End With
    
    SelectObject hdc, CreateFontIndirect(oLogFont)
    
    If lFontColor <> -99 Then
        dcSetTextColor hdc, lFontColor
    End If

    
End Function

Private Function dcGetFont(ByVal hdc As Long) As StdFont
    
    Set dcGetFont = New StdFont
    
    Dim sFontName As String
    sFontName = Space(1024)
    GetTextFace hdc, Len(sFontName), sFontName
    
    Dim tMetrics As TEXTMETRIC
    GetTextMetrics hdc, tMetrics
    
    Dim dY As Long
    dY = GetDeviceCaps(hdc, LOGPIXELSY)
    
    With dcGetFont
        .Name = sNT(sFontName)
        .Bold = (tMetrics.tmWeight = 700)
        .Charset = tMetrics.tmCharSet
        .Italic = tMetrics.tmItalic
        .Size = (tMetrics.tmAscent / tMetrics.tmDigitizedAspectY) * 72
'        Select Case True
'            Case tMetrics.tmPitchAndFamily And TMPF_FIXED_PITCH
'                .Size = (tMetrics.tmAscent / tMetrics.tmDigitizedAspectY) * 72
'            Case tMetrics.tmPitchAndFamily And TMPF_TRUETYPE
'                .Size = (tMetrics.tmAscent + tMetrics.tmInternalLeading) / 2
'        End Select
        .Strikethrough = tMetrics.tmStruckOut
        .Underline = tMetrics.tmUnderlined
        .Weight = tMetrics.tmWeight
    End With
    
End Function

Private Function dcGetTextColor(ByVal hdc As Long) As OLE_COLOR
    dcGetTextColor = GetTextColor(hdc)
End Function

Private Function dcSetTextColor(ByVal hdc As Long, ByVal lColor As OLE_COLOR)
    SetTextColor hdc, lColor
End Function


#End If  ' bStandAloneControls
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=



Public Property Get OffsetX() As Double
    OffsetX = mOffsetX
End Property

Public Property Let OffsetX(ByVal new_x As Double)
    mOffsetX = new_x
    PropertyChanged "OffsetX"
    UserControl_Paint
End Property

Public Property Get OffsetY() As Double
    OffsetY = mOffsetY
End Property

Public Property Let OffsetY(ByVal new_y As Double)
    mOffsetY = new_y
    PropertyChanged "OffsetY"
    UserControl_Paint
End Property



Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = lbl(1).ForeColor
End Property

Public Property Let ShadowColor(ByVal eColor As OLE_COLOR)
    lbl(1).ForeColor = eColor
    PropertyChanged "ShadowColor"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lbl(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lbl(0).ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = lbl(0).Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lbl(0).Enabled() = New_Enabled
    lbl(1).Enabled() = New_Enabled
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lbl(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lbl(0).Font = New_Font
    Set lbl(1).Font = New_Font
    PropertyChanged "Font"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As enumLabelBackstyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As enumLabelBackstyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As enumLabelBorderstyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As enumLabelBorderstyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    lbl(0).Refresh
End Sub

Private Sub lbl_Click(Index As Integer)
    RaiseEvent Click
    If Trim(URL) <> "" Then
        shellLaunchDoc URL
    End If
End Sub

Private Sub lbl_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    UserControl_Paint
End Sub


Private Sub UserControl_Click()
    RaiseEvent Click
    If Trim(URL) <> "" Then
        shellLaunchDoc URL
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
    
End Sub

Private Sub UserControl_Initialize()
'    lbl(1).ZOrder 0
    lbl(0).ZOrder 0
End Sub

Private Sub UserControl_InitProperties()
    mOffsetX = 1
    mOffsetY = 1
    lbl(1).ForeColor = vbWhite
    bReadProps = True
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = lbl(0).Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    lbl(0).Alignment() = New_Alignment
    lbl(1).Alignment() = New_Alignment
    PropertyChanged "Alignment"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As enumLabelAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As enumLabelAppearance)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = lbl(0).AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    lbl(0).AutoSize() = New_AutoSize
    lbl(1).AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lbl(0).Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lbl(0).Caption() = New_Caption
    lbl(1).Caption() = New_Caption
    PropertyChanged "Caption"
'    UserControl_Paint
    Redraw
End Property

Private Function Redraw()
    If bReadProps And UserControl.BackStyle = 0 Then
        On Error Resume Next
        UserControl.Parent.Cls
        Dim l As Long
        For l = 0 To UserControl.ParentControls.Count - 1
            If TypeOf UserControl.ParentControls(l) Is jeffLabel Then
                UserControl.ParentControls(l).Repaint
            End If
        Next l
    Else
        UserControl_Paint
    End If
    
End Function

Private Sub lbl_Change(Index As Integer)
    RaiseEvent Change
End Sub

'''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''''MappingInfo=lbl(0),lbl,0,DataFormat
''''Public Property Get DataFormat() As IStdDataFormatDisp
''''    Set DataFormat = lbl(0).DataFormat
''''End Property
''''
''''Public Property Set DataFormat(ByVal New_DataFormat As IStdDataFormatDisp)
''''    Set lbl(0).DataFormat = New_DataFormat
''''    Set lbl(1).DataFormat = New_DataFormat
''''    PropertyChanged "DataFormat"
''''End Property
''''
'''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''''MappingInfo=lbl(0),lbl,0,DataSource
''''Public Property Get DataSource() As DataSource
''''    Set DataSource = lbl(0).DataSource
''''End Property
''''
''''Public Property Set DataSource(ByVal New_DataSource As DataSource)
''''    Set lbl(0).DataSource = New_DataSource
''''    Set lbl(1).DataSource = New_DataSource
''''    PropertyChanged "DataSource"
''''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = lbl(0).FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lbl(0).FontBold() = New_FontBold
    lbl(1).FontBold() = New_FontBold
    PropertyChanged "FontBold"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = lbl(0).FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lbl(0).FontItalic() = New_FontItalic
    lbl(1).FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = lbl(0).FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lbl(0).FontName() = New_FontName
    lbl(1).FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = lbl(0).FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lbl(0).FontSize() = New_FontSize
    lbl(1).FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = lbl(0).FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lbl(0).FontStrikethru() = New_FontStrikethru
    lbl(1).FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = lbl(0).FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lbl(0).FontUnderline() = New_FontUnderline
    lbl(1).FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HyperLink
Public Property Get HyperLink() As HyperLink
Attribute HyperLink.VB_Description = "Returns a Hyperlink object used for browser style navigation."
    Set HyperLink = UserControl.HyperLink
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkExecute
Public Sub LinkExecute(ByVal Command As String)
Attribute LinkExecute.VB_Description = "Sends a command string to the source application in a DDE conversation."
    lbl(0).LinkExecute Command
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
    LinkItem = lbl(0).LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    lbl(0).LinkItem() = New_LinkItem
    lbl(1).LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkMode
Public Property Get LinkMode() As Integer
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
    LinkMode = lbl(0).LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As Integer)
    lbl(0).LinkMode() = New_LinkMode
    lbl(1).LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "Transfers contents of Label, PictureBox, or TextBox to source application in DDE conversation."
    lbl(0).LinkPoke
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "Asks the source DDE application to update the contents of a Label, PictureBox, or Textbox control."
    lbl(0).LinkRequest
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "Transfers contents of PictureBox to destination application in DDE conversation."
    lbl(0).LinkSend
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
    LinkTimeout = lbl(0).LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    lbl(0).LinkTimeout() = New_LinkTimeout
    lbl(1).LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
    LinkTopic = lbl(0).LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    lbl(0).LinkTopic() = New_LinkTopic
    lbl(1).LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = lbl(0).MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set lbl(0).MouseIcon = New_MouseIcon
    Set lbl(1).MouseIcon = New_MouseIcon
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = lbl(0).MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    lbl(0).MousePointer() = New_MousePointer
    lbl(1).MousePointer() = New_MousePointer
    UserControl.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub lbl_OLECompleteDrag(Index As Integer, Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    lbl(0).OLEDrag
End Sub

Private Sub lbl_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub lbl_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = lbl(0).OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    lbl(0).OLEDropMode() = New_OLEDropMode
    lbl(1).OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub lbl_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub lbl_OLESetData(Index As Integer, Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub lbl_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = lbl(0).RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    lbl(0).RightToLeft() = New_RightToLeft
    lbl(1).RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = lbl(0).ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    lbl(0).ToolTipText() = New_ToolTipText
    lbl(1).ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,UseMnemonic
Public Property Get UseMnemonic() As Boolean
Attribute UseMnemonic.VB_Description = "Returns/sets a value that specifies whether an & in a Label's Caption property defines an access key."
    UseMnemonic = lbl(0).UseMnemonic
End Property

Public Property Let UseMnemonic(ByVal New_UseMnemonic As Boolean)
    lbl(0).UseMnemonic() = New_UseMnemonic
    lbl(1).UseMnemonic() = New_UseMnemonic
    PropertyChanged "UseMnemonic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
    WordWrap = lbl(0).WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    lbl(0).WordWrap() = New_WordWrap
    lbl(1).WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
    Redraw
End Property

Private Sub UserControl_Paint()
    
    Static bPainting As Boolean
    If Not bPainting Then
        bPainting = True
        
        Dim lTopX As Long
        Dim lTopY As Long
        
        If UserControl.BorderStyle = 1 Then
            lTopX = twipsX(2)
            lTopY = twipsY(2)
        Else
            lTopX = 0
            lTopY = 0
        End If
        lbl(1).Move lTopX + twipsX(mOffsetX), lTopY + twipsY(mOffsetY)
'        lbl(0).ZOrder 1
        lbl(0).Move lTopX, lTopY
        'lbl(0).ZOrder 0
        lbl(0).Refresh
        
        UserControl_Resize
        
        ' Draw On Parent If Transparent
        Dim oPaint
        Dim rec As Rect
        
        Dim xMargin As Long
        Dim yMargin As Long
        If UserControl.BorderStyle <> 0 Then
            xMargin = 3
            yMargin = 3
        End If
        
        If UserControl.BackStyle <> 0 Then
            lbl(0).Visible = True
            lbl(1).Visible = True
        Else
            
            Set oPaint = UserControl.Parent
            rec.Left = pixelsX(UserControl.Extender.Left) + xMargin
            rec.Top = pixelsY(UserControl.Extender.Top) + yMargin
'        Else
'            Set oPaint = Me
'            rec.Left = pixelsX(lbl(0).Left)
'            rec.Top = pixelsY(lbl(0).Top)
'
'        End If
        rec.Right = rec.Left + pixelsX(lbl(0).Width) - xMargin
        rec.Bottom = rec.Top + pixelsY(lbl(0).Height) - yMargin
        
            lbl(0).Visible = False
            lbl(1).Visible = False
            
            Dim shadowrec As Rect
            shadowrec = rec
            With shadowrec
                .Left = .Left + OffsetX
                .Right = .Right + OffsetX
                .Top = .Top + OffsetY
                .Bottom = .Bottom + OffsetY
            End With
            
            Dim oFont As Font
            Set oFont = dcGetFont(oPaint.hdc)
            
            Dim oColor As OLE_COLOR
            oColor = oPaint.ForeColor
            
            Dim oOptions As enumDrawTextFormats
            If lbl(1).WordWrap Then
                oOptions = oOptions Or DT_WORDBREAK
            End If
            Select Case True
                Case lbl(0).Alignment = vbRightJustify
                    oOptions = oOptions Or DT_RIGHT
                Case lbl(0).Alignment = vbLeftJustify
                    oOptions = oOptions Or DT_LEFT
                Case lbl(0).Alignment = vbCenter
                    oOptions = oOptions Or DT_CENTER
            End Select
            
            dcSetFont oPaint.hdc, lbl(1).Font, lbl(1).ForeColor
            
            DrawText Parent.hdc, lbl(1).Caption, Len(lbl(1).Caption), shadowrec, oOptions
            
            dcSetFont oPaint.hdc, lbl(0).Font, lbl(0).ForeColor
            
            DrawText Parent.hdc, lbl(0).Caption, Len(lbl(0).Caption), rec, oOptions
            
            dcSetFont oPaint.hdc, oFont, oColor
        End If
        
        bPainting = False
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Not bReadingProps Then
        Debug.Print "UserControl_ReadProperties"
        bReadingProps = True
        Me.WordWrap = PropBag.ReadProperty("WordWrap", False)
        Me.AutoSize = PropBag.ReadProperty("AutoSize", False)
        
        UserControl_Paint
        Me.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
        Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
        Me.Enabled = PropBag.ReadProperty("Enabled", True)
        Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
        UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
        UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
        Me.Alignment = PropBag.ReadProperty("Alignment", 0)
        Me.Appearance = PropBag.ReadProperty("Appearance", 1)
        Me.Caption = PropBag.ReadProperty("Caption", "lbl")
    ''''    Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    ''''    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
        Me.FontBold = PropBag.ReadProperty("FontBold", Ambient.Font.Bold)
        Me.FontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font.Italic)
        Me.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
        Me.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.Size)
        Me.FontStrikethru = PropBag.ReadProperty("FontStrikethru", Ambient.Font.Strikethrough)
        Me.FontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font.Underline)
        Me.LinkItem = PropBag.ReadProperty("LinkItem", "")
        Me.LinkMode = PropBag.ReadProperty("LinkMode", 0)
        Me.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
        Me.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
        Set Me.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
        Me.MousePointer = PropBag.ReadProperty("MousePointer", 0)
        Me.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
        Me.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
        Me.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
        Me.UseMnemonic = PropBag.ReadProperty("UseMnemonic", True)
        Me.ShadowColor = PropBag.ReadProperty("ShadowColor", vbWhite)
        mOffsetX = PropBag.ReadProperty("OffsetX", 2)
        mOffsetY = PropBag.ReadProperty("OffsetY", 2)
        mURL = PropBag.ReadProperty("URL", "")
        bReadProps = True
        UserControl_Paint
        bReadingProps = False
    End If
    
End Sub

Private Sub UserControl_Resize()
    Debug.Print "UserControl_Resize: uc = " & UserControl.Width & "," & UserControl.Height & "  lbl = " & lbl(0).Left & "," & lbl(0).Top & "," & lbl(0).Width & "," & lbl(0).Height & "  lbl1 = " & lbl(1).Left & "," & lbl(1).Top & "," & lbl(1).Width & "," & lbl(1).Height
    
    Static bResizing As Boolean
    
    Static bFirstResize As Boolean
    
        Dim lTopX As Long
        Dim lTopY As Long
        If UserControl.BorderStyle = 1 Then
            lTopX = twipsX(1)
            lTopY = twipsY(1)
        Else
            lTopX = 0
            lTopY = 0
        End If
        
        lbl(1).Move lTopX + twipsX(mOffsetX), lTopY + twipsY(mOffsetY)
        lbl(0).Move lTopX, lTopY
        'lbl(0).ZOrder 0
'        lbl(0).Refresh
        
    If Not bFirstResize And bReadProps Then
        Dim bAutoSize As Boolean
        Dim xMargin As Long
        Dim yMargin As Long
        If UserControl.BorderStyle <> 0 Then
            xMargin = twipsX(1)
            yMargin = twipsY(1)
        End If
        bAutoSize = lbl(0).AutoSize
        lbl(1).AutoSize = False
        lbl(0).AutoSize = False
        With lbl(1)
            .Left = lTopX + twipsX(mOffsetX)
            .Top = lTopY + twipsY(mOffsetY)
            .Width = UserControl.Width - .Left - xMargin
            .Height = UserControl.Height - .Top - yMargin
        End With
        With lbl(0)
            .Left = lTopX
            .Top = lTopY
            .Width = UserControl.Width - .Left - xMargin
            .Height = UserControl.Height - .Top - yMargin
        End With
        lbl(1).AutoSize = bAutoSize
        lbl(0).AutoSize = bAutoSize
        bFirstResize = True
    End If
        
    If Not bResizing And bReadProps Then
        bResizing = True
        
        Dim lMarginX As Long
        Dim lMarginY As Long
        If UserControl.BorderStyle = 1 Then
            lMarginX = twipsX(1)
            lMarginY = twipsY(1)
        End If
                
        If lbl(0).AutoSize Then
            UserControl.Width = lbl(1).Left + lbl(1).Width + lMarginX
            If lbl(0).WordWrap Then
                UserControl.Height = lbl(1).Top + lbl(1).Height + lMarginY
            End If
        Else
            With lbl(1)
                .Move .Left, .Top, UserControl.Width - .Left, UserControl.Height - .Top
            End With
            With lbl(0)
                .Move .Left, .Top, UserControl.Width - lbl(1).Left, UserControl.Height - lbl(1).Top
            End With
        End If
                
        Redraw
                
        bResizing = False
        
    End If
    
End Sub



Private Sub UserControl_Show()
    UserControl_Paint
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    If Not bReadingProps Then
    Debug.Print "UserControl_WriteProperties"
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", lbl(0).ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", lbl(0).Enabled, True)
    Call PropBag.WriteProperty("Font", lbl(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Alignment", lbl(0).Alignment, 0)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("AutoSize", lbl(0).AutoSize, False)
    Call PropBag.WriteProperty("Caption", lbl(0).Caption, "lbl")
''''    Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
''''    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("FontBold", lbl(0).FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lbl(0).FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lbl(0).FontName, "")
    Call PropBag.WriteProperty("FontSize", lbl(0).FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lbl(0).FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lbl(0).FontUnderline, 0)
    Call PropBag.WriteProperty("LinkItem", lbl(0).LinkItem, "")
    Call PropBag.WriteProperty("LinkMode", lbl(0).LinkMode, 0)
    Call PropBag.WriteProperty("LinkTimeout", lbl(0).LinkTimeout, 50)
    Call PropBag.WriteProperty("LinkTopic", lbl(0).LinkTopic, "")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lbl(0).MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", lbl(0).OLEDropMode, 0)
    Call PropBag.WriteProperty("RightToLeft", lbl(0).RightToLeft, False)
    Call PropBag.WriteProperty("ToolTipText", lbl(0).ToolTipText, "")
    Call PropBag.WriteProperty("UseMnemonic", lbl(0).UseMnemonic, True)
    Call PropBag.WriteProperty("WordWrap", lbl(0).WordWrap, False)
    Call PropBag.WriteProperty("ShadowColor", lbl(1).ForeColor, vbWhite)
    Call PropBag.WriteProperty("OffsetX", mOffsetX, 2)
    Call PropBag.WriteProperty("OffsetY", mOffsetY, 2)
    Call PropBag.WriteProperty("URL", mURL, "")
    
    End If
End Sub

Public Function Repaint()
    UserControl_Paint
End Function


Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property


Public Property Get URL() As String
    URL = mURL
End Property

Public Property Let URL(ByVal new_url As String)
    mURL = new_url
    PropertyChanged "URL"
    
End Property
