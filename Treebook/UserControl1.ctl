VERSION 5.00
Begin VB.UserControl TreeBook 
   BackColor       =   &H005A371B&
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   KeyPreview      =   -1  'True
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   407
   Begin VB.PictureBox DropDownIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   5055
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox IconBuff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   4320
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picDetails 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4680
      Picture         =   "UserControl1.ctx":04EA
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picScr 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   5760
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   6
      Top             =   0
      Width           =   255
      Begin VB.PictureBox picScroller 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   0
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   7
         Top             =   1080
         Width           =   255
         Begin VB.Line LN 
            X1              =   0
            X2              =   16
            Y1              =   112
            Y2              =   120
         End
      End
      Begin Project1.Button BtnDown 
         Height          =   270
         Left            =   -90
         TabIndex        =   21
         Top             =   4635
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
      End
      Begin Project1.Button BtnUp 
         Height          =   270
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
      End
   End
   Begin VB.PictureBox picHideDetails 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4680
      Picture         =   "UserControl1.ctx":09E8
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picScrH 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   4560
      Width           =   3495
      Begin Project1.Button BtnLeft 
         Height          =   270
         Left            =   3150
         TabIndex        =   18
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
      End
      Begin VB.PictureBox picScrollerH 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   1
         Top             =   0
         Width           =   1815
         Begin VB.Line lnH 
            X1              =   72
            X2              =   72
            Y1              =   0
            Y2              =   16
         End
      End
      Begin Project1.Button BtnRight 
         Height          =   270
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
      End
   End
   Begin VB.Timer tmrScr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   1680
   End
   Begin VB.PictureBox picMnuText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicBuff3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picExtendet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picBuffText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   12
      Top             =   120
      Width           =   3975
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox PicRename 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
         Begin VB.TextBox txtRename 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox picShdw 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin Project1.Button Btn 
         Height          =   270
         Left            =   45
         TabIndex        =   20
         Top             =   1530
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
      End
   End
   Begin VB.Image LIcon 
      Height          =   255
      Index           =   0
      Left            =   4320
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image LIconS 
      Height          =   255
      Index           =   0
      Left            =   4320
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   0
      Left            =   4320
      Picture         =   "UserControl1.ctx":0EE6
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   1
      Left            =   4320
      Picture         =   "UserControl1.ctx":13C0
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   2
      Left            =   4320
      Picture         =   "UserControl1.ctx":189A
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btnSCR 
      Height          =   135
      Index           =   3
      Left            =   4320
      Picture         =   "UserControl1.ctx":1D74
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image IconS 
      Height          =   255
      Index           =   0
      Left            =   4800
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Icon 
      Height          =   255
      Index           =   0
      Left            =   4800
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image MnuIconExtendet 
      Height          =   240
      Index           =   0
      Left            =   4680
      Top             =   1320
      Width           =   240
   End
End
Attribute VB_Name = "TreeBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


'== Image list
Private Const CLR_NONE     As Long = &HFFFFFFFF
Private Const ILC_MASK     As Long = &H1
Private Const ILC_COLORDDB As Long = &HFE

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Add Lib "Comctl32" (ByVal hImageList As Long, ByVal hBitmap As Long, ByVal hBitmapMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "Comctl32" (ByVal hImageList As Long) As Long


'Graphics
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

'AlphaBlending - to get blendet colors (buttons etc.)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Type ColorAndAlpha
    R                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


'Events
Public Event NodeClick(ByVal Node As Long)
Public Event NodeCheck(ByVal Node As Long)
Public Event NodeDblClick(ByVal Node As Long)
Public Event ListName(ListItem As String)
Public Event ButtonDown(Clicked As Boolean)
Public Event DbClick()
Public Event Resize()

Public TopLine As Long
Public Listwidth As Long
Public ListText As String
Public ListIndex As Long

'Colors
Public SelBorderColor As OLE_COLOR
Public SelBackColor As OLE_COLOR
Public SelForeColor As OLE_COLOR
Public oBackColor As OLE_COLOR
Public oForeColor As OLE_COLOR

'Dim dPath As String
'Dim DoGen As Boolean
'Dim qk As Boolean

Private Type Whatever
    name As String
    Extension As String
    FullName As String
    Selected As Boolean
    ShowDetails As Boolean
    Details As String
    IconIndex As Long
    selIcon As PictureBox
    index As Long
    sKey As String
End Type

'Dim qPathExists As Boolean
Dim TVDispInfo() As Whatever

Dim iTop As Long
Dim iLeft As Integer
Dim iHeight As Long
Dim SelHeight As Integer
Dim selDetHeight As Integer
Public DoNotGenerate As Boolean
Dim MouseY As Single
Dim MouseX As Single
Dim bSelected As Long
Dim sSelected As Long

Public ListCount As Long
Public Function AddBitmap( _
                ByVal hBitmap As Long, _
                Optional ByVal MaskColor As Long = CLR_NONE _
                ) As Long
    
    If (m_hImageList) Then
        If (MaskColor <> CLR_NONE) Then
            AddBitmap = ImageList_AddMasked(m_hImageList, hBitmap, MaskColor)
          Else
            AddBitmap = ImageList_Add(m_hImageList, hBitmap, 0)
        End If
        m_lImageListCount = ImageList_GetImageCount(m_hImageList)
    End If
End Function


Public Sub ExpandAll()
  For ICNT = 0 To ListCount
      TVDispInfo(ICNT).ShowDetails = True
  Next ICNT
  UserControl_Resize
End Sub
Public Sub CollapseAll()
  For ICNT = 0 To ListCount
      TVDispInfo(ICNT).ShowDetails = False
  Next ICNT
  UserControl_Resize
End Sub
Public Sub Expand(ByVal key As Long)
  TVDispInfo(key).ShowDetails = True
  UserControl_Resize
End Sub
Public Sub Collapse(ByVal key As Long)
  TVDispInfo(key).ShowDetails = False
  UserControl_Resize
End Sub
Public Function GetCaption(ByVal key As Long) As String
    GetCaption = TVDispInfo(key).FullName
End Function
Public Sub Clear()
  
  ReDim TVDispInfo(0 To 1)
  For i = 1 To ListCount
    Unload Icon(i)
    Unload IconS(i)
    Unload LIcon(i)
    Unload LIconS(i)
  Next i

  '-- Reset count
  ListCount = 0
  UserControl_Resize
End Sub

Public Function AddNode( _
                Optional ByVal key As String, _
                Optional ByVal Text As String, _
                Optional ByVal SubText As String, _
                Optional ByVal Image As StdPicture, _
                Optional ByVal SmalIcon As StdPicture) As Long
    
  Dim ICNT As Integer
  List1.AddItem Text
  ListCount = ListCount + 1
  
  ReDim Preserve TVDispInfo(ListCount)

  ICNT = ListCount
  TVDispInfo(ICNT).sKey = key
  TVDispInfo(ICNT).name = Text
  TVDispInfo(ICNT).Details = SubText
  TVDispInfo(ICNT).FullName = Text
  TVDispInfo(ICNT).Selected = False
  TVDispInfo(ICNT).ShowDetails = False
  
  TVDispInfo(ICNT).index = ListCount

  If ICNT > 0 Then
        Load Icon(ICNT)
        Load IconS(ICNT)
        Load LIcon(ICNT)
        Load LIconS(ICNT)
  
  End If
  TVDispInfo(ICNT).IconIndex = Icon().Count - 1
  Set Icon(ICNT) = SmalIcon
  Set IconS(ICNT) = SmalIcon
  Set LIcon(ICNT) = Image
  Set LIconS(ICNT) = Image
  
'Resize
  UserControl_Resize
  AddNode = ListCount

  Exit Function

errH:
    
    AddNode = 0
End Function

Public Property Get SelectedNode() As Long
  SelectedNode = bSelected
End Property
Public Property Let SelectedNode(ByVal hNode As Long)
  TVDispInfo(bSelected).Selected = False
  TVDispInfo(hNode).Selected = True
  bSelected = hNode
  iSelected = False
  UserControl_Resize
End Property

Public Property Get NodeText(ByVal hNode As Long) As String
    NodeText = TVDispInfo(hNode).name
End Property
Public Property Let NodeText(ByVal hNode As Long, ByVal New_NodeText As String)
  TVDispInfo(hNode).name = New_NodeText
  UserControl_Resize
End Property

Public Function GetKeyNode(ByVal key As String) As Long
  GetKeyNode = m_cKey(key)
End Function

Public Function GetNodeKey(ByVal hNode As Long) As String
  GetNodeKey = TVDispInfo(bSelected).sKey
End Function
Public Function GetPhoto(ByVal hNode As Long) As StdPicture
  Set GetPhoto = LIcon(hNode)
End Function

Private Function AlphaBlend(ByVal FirstColor As Long, ByVal SecondColor As Long, ByVal AlphaValue As Long) As Long
    Dim iForeColor         As ColorAndAlpha
    Dim iBackColor         As ColorAndAlpha
    
    OleTranslateColor FirstColor, 0, VarPtr(iForeColor)
    OleTranslateColor SecondColor, 0, VarPtr(iBackColor)
    With iForeColor
        .R = (.R * AlphaValue + iBackColor.R * (255 - AlphaValue)) / 255
        .G = (.G * AlphaValue + iBackColor.G * (255 - AlphaValue)) / 255
        .B = (.B * AlphaValue + iBackColor.B * (255 - AlphaValue)) / 255
    End With
    CopyMemory VarPtr(AlphaBlend), VarPtr(iForeColor), 4
    
End Function

Public Function CheckHeight() As Long
  On Error Resume Next
  'here we get the height of thelist
  Dim ICNT As Integer
  CheckHeight = 0
  If dView = 3 Then
      CheckHeight = picBuff.Height * ColumnCount
  Else
      For ICNT = 1 To ListCount
          If TVDispInfo(ICNT).ShowDetails = True Then
              CheckHeight = CheckHeight + selDetHeight
          Else
              CheckHeight = CheckHeight + SelHeight
          End If
      Next ICNT
  End If

End Function

Private Sub DrawBackgrounds()
'DrawsTheBackground

    With picBackground
        .Cls
        .Width = PicView.Width
        If dView = 3 Then
            .Height = picBuff.Height * 2
        Else
            .Height = SelHeight * 2
            picBuff.Height = .Height / 2
        End If
        
        If dView = 1 Or dView = 2 Then

            .Width = ColumnWidth
        End If
        If dView = 0 Then picBuff.Width = .Width
    End With
    
    With picBuff
        .Cls
        .BackColor = oBackColor
        .ForeColor = oForeColor
    End With
    
    BitBlt picBackground.hDC, 0, 0, picBackground.Width, picBackground.Height / 2, picBuff.hDC, 0, 0, SRCCOPY
     
    With picBuff
        .Cls
        .BackColor = SelBackColor
        .ForeColor = SelBorderColor
    End With
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
        
    BitBlt picBackground.hDC, 0, picBackground.Height / 2, picBackground.Width, picBackground.Height / 2, picBuff.hDC, 0, 0, SRCCOPY

    picBackground.Refresh
    
    If dView = 0 Then DrawExtendet
     
End Sub

Private Sub DrawExtendet()
'DrawsExtendetBackground

    With picExtendet
        .Cls
        .Width = PicView.Width
        .Height = selDetHeight * 2
        picBuff.Height = selDetHeight
        picBuff.Width = .Width
    End With
    
    With picBuff
        .Cls
        .BackColor = AlphaBlend(oBackColor, vbBlack, 245)
        .ForeColor = AlphaBlend(oBackColor, vbBlack, 225)
    End With
        
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    
    BitBlt picExtendet.hDC, 0, 0, picExtendet.Width, picExtendet.Height / 2, picBuff.hDC, 0, 0, SRCCOPY
    
    picBuff.BackColor = oBackColor
    BitBlt picExtendet.hDC, 1, 1, 54, picExtendet.Height / 2 - 2, picBuff.hDC, 0, 0, SRCCOPY
    
    With picBuff
        .Cls
        .BackColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
        .ForeColor = SelBorderColor
    End With
    
    picBuff.Line (0, 0)-(picBuff.Width, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
        
    BitBlt picExtendet.hDC, 0, picExtendet.Height / 2, picExtendet.Width, picExtendet.Height / 2, picBuff.hDC, 0, 0, SRCCOPY
    
    picBuff.BackColor = SelBackColor
    BitBlt picExtendet.hDC, 1, picExtendet.Height / 2 + 1, 55, picExtendet.Height / 2 - 2, picBuff.hDC, 0, 1, SRCCOPY

    picExtendet.Refresh
    picBuff.Height = SelHeight
     
End Sub

Private Sub DrawScrollers()
  If dScrollersType = 0 Then
    'Draws buttons - just read it line by line - you should see what it does
    picBuff.Width = 15
    picBuff.Height = 15
    picBuff.Cls
    
    picBuff.BackColor = &H8000000F
    
    picBuff.ForeColor = &H80000014
    picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height - 1)
    
    picBuff.ForeColor = &H80000010
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
    
    IconBuff.BackColor = picBuff.BackColor
    IconBuff.Picture = btnSCR(0).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnUp.NormalImage = picBuff.Image
    Set BtnUp.FocusedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(1).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnDown.NormalImage = picBuff.Image
    Set BtnDown.FocusedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(2).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnLeft.NormalImage = picBuff.Image
    Set BtnLeft.FocusedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(3).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnRight.NormalImage = picBuff.Image
    Set BtnRight.FocusedImage = picBuff.Image
    
    picBuff.ForeColor = &H80000010
    picBuff.Line (0, 0)-(picBuff.Width - 1, 0)
    picBuff.Line (0, 0)-(0, picBuff.Height - 1)
    picBuff.Line (picBuff.Width - 1, 0)-(picBuff.Width - 1, picBuff.Height)
    picBuff.Line (0, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

    IconBuff.Picture = btnSCR(0).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnUp.PressedImage = picBuff.Image

    IconBuff.Picture = btnSCR(1).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnDown.PressedImage = picBuff.Image
  
    IconBuff.Picture = btnSCR(2).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnLeft.PressedImage = picBuff.Image
    
    IconBuff.Picture = btnSCR(3).Picture
    
    BitBlt picBuff.hDC, (picBuff.Width - IconBuff.Width) / 2 + 1, (picBuff.Height - IconBuff.Height) / 2 + 1, IconBuff.Width, IconBuff.hwnd, IconBuff.hDC, 0, 0, SRCCOPY

    Set BtnRight.PressedImage = picBuff.Image
    
    'Draw scrollers - follow it line by line;p
    picScr.Width = BtnUp.Width
    picScroller.Width = BtnUp.Width
     
    picScroller.Height = 1700 ' this is somehow the maximum visible size on the screen - we need to set this, so it draws the whole scroller - because of bitblt
    
    picScroller.BackColor = &H8000000F
    
    picScroller.ForeColor = &H80000014
    
    picScroller.Line (0, 0)-(picScroller.Width - 1, 0)
    picScroller.Line (0, 0)-(0, picScroller.Height)
    
    picScroller.ForeColor = &H80000010
    picScroller.Line (picScroller.Width - 1, 0)-(picScroller.Width - 1, picScroller.Height)

    picScrH.Height = BtnLeft.Height
    picScrollerH.Height = BtnLeft.Height
    
    LN.X1 = 0
    LN.X2 = picScroller.Width
    LN.BorderColor = &H80000010
    
    picScrollerH.Width = 2050 ' this is somehow the maximum visible size on the screen - we need to set this, so it draws the whole scroller - because of bitblt
    
    picScrollerH.BackColor = &H8000000F
    
    picScrollerH.ForeColor = &H80000014
    
    picScrollerH.Line (0, 0)-(picScrollerH.Width, 0)
    picScrollerH.Line (0, 0)-(0, picScrollerH.Height - 1)
    
    picScrollerH.ForeColor = &H80000010
    picScrollerH.Line (0, picScrollerH.Height - 1)-(picScrollerH.Width, picScrollerH.Height - 1)

    lnH.Y1 = 0
    lnH.Y2 = picScrollerH.Height
    lnH.BorderColor = &H80000010
    
    
  End If

End Sub

Private Sub btnDOWN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
RaiseEvent ButtonDown(False)
If Button = vbLeftButton Then
    tmrScr.Tag = "D"
    tmrScr.Interval = 1500
    tmrScr_Timer
    tmrScr.Enabled = True
End If

End Sub

Private Sub btnDOWN_TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrScr.Enabled = False

End Sub

Private Sub btnUP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
RaiseEvent ButtonDown(False)
If Button = vbLeftButton Then
    tmrScr.Tag = "U"
    tmrScr.Interval = 1500
    tmrScr_Timer
    tmrScr.Enabled = True
End If

End Sub

Private Sub btnUP_TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
tmrScr.Enabled = False

End Sub

Public Sub DeleteSelected()
  Dim ICNT As Long
  Dim ref As Boolean
  For ICNT = 1 To ListCount
      If TVDispInfo(ICNT).Selected = True Then
  
      End If
  Next ICNT
  If ref = True Then Refresh

End Sub

Public Sub Generate()
  PicView.Cls
  PicView.BackColor = oBackColor

  If DoNotGenerate = True Then GoTo CCEND
  
  SelHeight = 22        'NormalviewHeight
  selDetHeight = 96     '"Extendet" dView Mode

      
  Dim FromIndex As Long
  Dim ToIndex As Long

If dView = 0 Then
    
    If Not PicView.Height = UserControl.Height / Screen.TwipsPerPixelX - 2 Then PicView.Height = UserControl.Height / Screen.TwipsPerPixelX - 2
    picScrH.Visible = False
    
    IconBuff.Height = 16
    IconBuff.Width = 16
    
    FontSize = 10
    
    'SetHeight
    iHeight = CheckHeight
    
    DrawBackgrounds
    
    picBuffText.Height = picBuff.Height - 2
    
    ToIndex = ListCount
    Dim xTop As Integer
    Dim cHeight As Integer
    Dim cLeft As Integer
    Dim fTop As Long
    FromIndex = 0
    On Error GoTo loopend
    'iTop = 20
    Do Until fTop > -iTop
        FromIndex = FromIndex + 1
        xTop = fTop + iTop
        If TVDispInfo(FromIndex).ShowDetails = True Then
    
            fTop = fTop + selDetHeight
    
        Else
            fTop = fTop + SelHeight
        End If
    Loop
loopend:

    If FromIndex < 1 Then FromIndex = 1
    fTop = 0
    'ListCount = 20
    For ICNT = 1 To ListCount
        If TVDispInfo(ICNT).ShowDetails = True Then
            fTop = fTop + selDetHeight
        Else
            fTop = fTop + SelHeight
        End If
        
        If fTop + iTop >= PicView.Height Then
            ToIndex = ICNT
            Exit For
        End If
    Next ICNT
    
    On Error Resume Next 'GoTo NoDraw

    For ICNT = FromIndex To ToIndex
    If ListCount > 0 Then
    
        If TVDispInfo(ICNT).ShowDetails = True Then
            picBuffText.Font.Size = 9
            picBuffText.Font.Bold = True
            

            
            If TVDispInfo(ICNT).Selected = True Then
                'plaats photo in selected area
                IconBuff.Picture = LIconS(TVDispInfo(ICNT).IconIndex).Picture
                
                picBuff.Height = picExtendet.Height / 2
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picExtendet.hDC, 0, picExtendet.Height / 2, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(TVDispInfo(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                picBuffText.ForeColor = SelForeColor
                picBuffText.Font = "Arial"
                picBuffText.FontSize = 8
                picBuffText.Print TVDispInfo(ICNT).name
                
                cLeft = 1
                picHideDetails.BackColor = picBuff.BackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = 100
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuffText.Font.Size = picBuffText.Font.Size - 2
                picBuffText.Font.Bold = True
                
                picBuffText.Font = "TAHOMA"
                picBuffText.FontSize = 9
                picBuffText.Width = picBuff.Width - cLeft - 8
                picBuffText.Height = picBuffText.TextHeight("gŽ") * 4
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                picBuffText.ForeColor = AlphaBlend(SelBackColor, SelForeColor, 110)
                

                Dim asd() As String
                asd() = Split(TVDispInfo(ICNT).Details, "|")
                picBuffText.Print asd(0)
                picBuffText.Print asd(1)
                picBuffText.Print asd(2)
                picBuffText.Print asd(3)
                
                picBuffText.Font.Size = picBuffText.Font.Size + 2
                picBuffText.Font.Bold = False
                cLeft = cLeft + 7
                BitBlt picBuff.hDC, cLeft, picBuffText.Height / 4 + 6, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY

                
                picBuffText.Height = picBuffText.TextHeight("gŽ")
                
                picBuff.Refresh
                
                
                
                'the following code "deletes" the upper and lower lines in multiple selections
                If ICNT > 1 Then
                    If TVDispInfo(ICNT - 1).Selected = True Then
                        picBuff.ForeColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                        picBuff.Line (56, 0)-(picBuff.Width - 1, 0)
                        If TVDispInfo(ICNT - 1).ShowDetails = True Then picBuff.ForeColor = SelBackColor
                        picBuff.Line (1, 0)-(56, 0)
                    End If
                End If
    
                If ICNT < ListCount Then
                    If TVDispInfo(ICNT + 1).Selected = True Then
                        picBuff.ForeColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                    
                    End If
                End If
                                
                BitBlt PicView.hDC, 0, xTop, PicView.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height
            Else
                IconBuff.Picture = LIcon(TVDispInfo(ICNT).IconIndex).Picture
                
                picBuff.Height = picExtendet.Height / 2
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picExtendet.hDC, 0, 0, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(TVDispInfo(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(oBackColor, vbBlack, 245)
                picBuffText.ForeColor = ForeColor
                picBuffText.Print TVDispInfo(ICNT).name
                
                cLeft = 1
                picHideDetails.BackColor = oBackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = 100
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuffText.Font.Size = picBuffText.Font.Size - 2
                picBuffText.Font.Bold = True
                
                picBuffText.Width = picBuff.Width - cLeft - 8
                picBuffText.Height = picBuffText.TextHeight("gŽ") * 4
                picBuffText.Cls
                picBuffText.BackColor = AlphaBlend(oBackColor, vbBlack, 245)
                picBuffText.ForeColor = AlphaBlend(oBackColor, ForeColor, 110)
                
                asd() = Split(TVDispInfo(ICNT).Details, "|")
                
                picBuffText.Print asd(0)
                picBuffText.Print asd(1)
                picBuffText.Print asd(2)
                picBuffText.Print asd(3)
                
                picBuffText.Font.Size = picBuffText.Font.Size + 2
                picBuffText.Font.Bold = False
                cLeft = cLeft + 7
                BitBlt picBuff.hDC, cLeft, picBuffText.Height / 4 + 6, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY

                
                picBuffText.Height = picBuffText.TextHeight("gŽ")
                
                picBuff.Refresh

                'the following code "deletes" the upper and lower lines in multiple selections
                If ICNT > 1 Then
                    If TVDispInfo(ICNT - 1).Selected = False Then
                        If TVDispInfo(ICNT - 1).ShowDetails = True Then picBuff.ForeColor = AlphaBlend(oBackColor, vbBlack, 245) Else picBuff.ForeColor = AlphaBlend(oBackColor, vbBlack, 225)
                        picBuff.Line (0, 0)-(picBuff.Width, 0)
                        picBuff.ForeColor = oBackColor
                        If TVDispInfo(ICNT - 1).ShowDetails = True Then picBuff.Line (0, 0)-(55, 0)
                    End If
                End If
    
                If ICNT < ListCount Then
                    If TVDispInfo(ICNT + 1).Selected = False Then
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                    
                    End If
                End If
                                
                BitBlt PicView.hDC, 0, xTop, PicView.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height

            End If
            
        picBuff.Height = SelHeight
        picBuffText.Font.Size = 8
        picBuffText.Font.Bold = False
        
        Else
        
            'to add an icon infront of the caption
            If TVDispInfo(ICNT).Selected = True Then
                

                
                IconBuff.Picture = IconS(TVDispInfo(ICNT).IconIndex).Picture
    
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, picBackground.Height / 2, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(TVDispInfo(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = SelBackColor
                picBuffText.ForeColor = SelForeColor
                picBuffText.Print TVDispInfo(ICNT).name
                
                cLeft = 1
                picDetails.BackColor = picBuff.BackColor
                DropDownIcon.BackColor = picBuff.BackColor
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                BitBlt picBuff.hDC, PicView.Width - 12, (picBuff.Height - DropDownIcon.Height) / 2, DropDownIcon.Width, DropDownIcon.Height, DropDownIcon.hDC, 0, 0, SRCCOPY
                
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.Refresh
                
                picBuff.ForeColor = SelBackColor
                
                'the following code "deletes" the upper and lower lines in multiple selections
                If ICNT > 1 Then
                    If TVDispInfo(ICNT - 1).Selected = True Then
                        picBuff.Line (1, 0)-(picBuff.Width - 1, 0)
                    
                    End If
                End If
    
                If ICNT < ListCount Then
                    If TVDispInfo(ICNT + 1).Selected = True Then
                        picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)
                    
                    End If
                End If
                                
                BitBlt PicView.hDC, 0, xTop, PicView.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height
            Else
                IconBuff.Picture = Icon(TVDispInfo(ICNT).IconIndex).Picture
            
                BitBlt picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, picBackground.hDC, 0, 0, SRCCOPY
                
                picBuffText.Width = picBuffText.TextWidth(TVDispInfo(ICNT).name)
                picBuffText.Cls
                picBuffText.BackColor = oBackColor
                picBuffText.ForeColor = oForeColor
                picBuffText.Print TVDispInfo(ICNT).name
                
                cLeft = 1
                
                picDetails.BackColor = picBuffText.BackColor
                DropDownIcon.BackColor = picBuffText.BackColor
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - picDetails.Height) / 2, picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                BitBlt picBuff.hDC, PicView.Width - 12, (picBuff.Height - DropDownIcon.Height) / 2, DropDownIcon.Width, DropDownIcon.Height, DropDownIcon.hDC, 0, 0, SRCCOPY

                
                cLeft = cLeft + 1 + picDetails.Width
                
                BitBlt picBuff.hDC, cLeft, (picBuff.Height - IconBuff.Height) / 2, IconBuff.Width, IconBuff.Height, IconBuff.hDC, 0, 0, SRCCOPY
                cLeft = cLeft + 1 + IconBuff.Width
                
                BitBlt picBuff.hDC, cLeft, 1, picBuffText.Width, picBuffText.Height, picBuffText.hDC, 0, 0, SRCCOPY
                
                picBuff.ForeColor = AlphaBlend(SelBackColor, SelBorderColor, 230)
                picBuff.Line (1, picBuff.Height - 1)-(picBuff.Width - 1, picBuff.Height - 1)

                
                picBuff.Refresh
    
    
                BitBlt PicView.hDC, 0, xTop, PicView.Width, picBuff.Height, picBuff.hDC, 0, 0, SRCCOPY
                
                xTop = xTop + picBuff.Height
                
            End If
        
        End If
    End If
    Next ICNT
    
  End If
NoDraw:
CCEND:

  PicView.Refresh

End Sub

Private Sub btn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
PicView_MouseDown Button, Shift, X + Btn.Left, Y + Btn.Top

PicBuff3.ForeColor = vbHighlight
PicBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
PicBuff3.Cls
PicBuff3.Picture = LoadPicture("")
If TVDispInfo(Btn.Tag).ShowDetails = False Then
     PicBuff3.Height = SelHeight - 4
     PicBuff3.Width = picDetails.Width - 2
     
     picDetails.BackColor = PicBuff3.BackColor
     BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
     
     PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
     PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
     PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
     PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)

     Set Btn.NormalImage = PicBuff3.Image
     Set Btn.FocusedImage = PicBuff3.Image
     
     PicBuff3.ForeColor = vbHighlight
     PicBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

     picDetails.BackColor = PicBuff3.BackColor
     BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
     
     PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
     PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
     PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
     PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)

     Set Btn.PressedImage = PicBuff3.Image
     
Else
     PicBuff3.Cls
     PicBuff3.Height = selDetHeight - 4
     PicBuff3.Width = picDetails.Width - 2
     
     picHideDetails.BackColor = PicBuff3.BackColor
     BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
     
     PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
     PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
     PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
     PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)
     
     Set Btn.NormalImage = PicBuff3.Image
     Set Btn.FocusedImage = PicBuff3.Image
     
     PicBuff3.ForeColor = vbHighlight
     PicBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

     picHideDetails.BackColor = PicBuff3.BackColor
     BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picHideDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
     
     PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
     PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
     PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
     PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)

     Set Btn.PressedImage = PicBuff3.Image
 End If
 
 If Btn.Top <> GetFileTop(Btn.Tag) + iTop + 2 Then
    Btn.Top = GetFileTop(Btn.Tag) + iTop + 2
 End If
 
End Sub

Private Sub btn_MouseOut()
Btn.Visible = False

End Sub

Private Sub SetScroller()
  On Error Resume Next
  If PicView.Height * (PicView.Height - BtnUp.Height - BtnDown.Height) / (CheckHeight) > 10 Then
      picScroller.Height = PicView.Height * (PicView.Height - BtnUp.Height - BtnDown.Height) / (CheckHeight)
  Else
      picScroller.Height = 10
  End If
  
  If Not picScroller.Top = (iTop) * (PicView.Height - BtnUp.Height - BtnDown.Height - picScroller.Height) / (PicView.Height - (CheckHeight)) + BtnUp.Height Then picScroller.Top = (iTop) * (PicView.Height - BtnUp.Height - BtnDown.Height - picScroller.Height) / (PicView.Height - (CheckHeight)) + BtnUp.Height
  
  picScr.Refresh
End Sub

Private Sub SetScrollerH()
  On Error Resume Next
  If PicView.Width * (PicView.Width - BtnLeft.Width - BtnRight.Width) / (ColumnCount * picBuff.Width) > 10 Then
      picScrollerH.Width = PicView.Width * (PicView.Width - BtnLeft.Width - BtnRight.Width) / (ColumnCount * picBuff.Width)
  Else
      picScrollerH.Width = 10
  End If
  
  picScrollerH.Left = (iLeft) * (PicView.Width - BtnLeft.Width - BtnRight.Width - picScrollerH.Width) / (PicView.Width - (ColumnCount * picBuff.Width)) + BtnLeft.Width
  
  picScrH.Refresh
End Sub

Private Sub PicView_Click()
  ListIndex = bSelected
End Sub

Private Sub PicView_DblClick()
On Error Resume Next

'If bSelected > 0 Then RaiseEvent FileSelect(GetSelectedFiles)

Dim isSelcted As Boolean
Dim ICNT As Integer
isSelcted = False
For ICNT = 1 To ListCount
    If TVDispInfo(ICNT).Selected = True And Len(TVDispInfo(ICNT).FullName) > 3 And TVDispInfo(ICNT).FullName <> "#MYCOMPUTER" Then
        isSelcted = True
        ListText = TVDispInfo(ICNT).FullName
    End If
Next ICNT
'RaiseEvent DeletableItemSelected(isSelcted)
RaiseEvent ListName(TVDispInfo(bSelected).FullName)
RaiseEvent NodeDblClick(bSelected)
RaiseEvent DbClick
End Sub

Private Sub PicView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iSelected As Long
Dim ICNT As Long
Dim IcNT2 As Long
Dim Rec As RECT

If ListCount < 1 And Button = vbLeftButton Then Exit Sub
On Error Resume Next
'first we have to calculate wich item the user intented to select
If dView = 0 Then
    For ICNT = 1 To ListCount
        If GetFileTop(ICNT) <= Y - iTop And GetFileTop(ICNT + 1) > Y - iTop Then
            iSelected = ICNT
            'RaiseEvent ButtonDown(False)
            If X < picDetails.Width + 3 Then
                If Button = vbLeftButton Then
                    RaiseEvent ButtonDown(False)
                    TVDispInfo(iSelected).ShowDetails = Not TVDispInfo(iSelected).ShowDetails
                    
                    If TVDispInfo(iSelected).ShowDetails = True And GetFileTop(ICNT) > PicView.Height - selDetHeight - iTop Then
                        iTop = PicView.Height - selDetHeight - GetFileTop(ICNT)
                    ElseIf TVDispInfo(iSelected).ShowDetails = True And GetFileTop(ICNT) + iTop < 0 Then
                        iTop = -GetFileTop(ICNT)
                    End If
                    
                    If TVDispInfo(iSelected).Selected = True Then
                        UserControl_Resize
                    End If
                End If
            ElseIf X > PicView.Width - 20 Then
                If Button = vbLeftButton Then
                    TopLine = Int(Y \ 22) * 22
                    If picScr.Visible = True Then
                      Listwidth = Width - (picScr.Width * Screen.TwipsPerPixelX)
                    Else
                      Listwidth = Width
                    End If
                    ListText = TVDispInfo(ICNT).FullName
                    
                    If bSelected = ICNT And iSelected = ICNT Then
                      iSelected = 0
                      RaiseEvent ButtonDown(False)
                    Else
                      bSelected = ICNT
                      RaiseEvent ButtonDown(True)
                    End If
                    
                End If
            Else
              RaiseEvent ButtonDown(False)
              RaiseEvent NodeClick(ICNT)
              RaiseEvent NodeCheck(Me.SelectedNode)
            End If
            Exit For
        End If
    Next ICNT
End If

'now we have to select items accordingly to type of selection

If dView = 0 Then
    If iSelected < 1 Then
        If GetFileTop(ListCount) <= Y - iTop And CheckHeight > Y - iTop Then
            iSelected = ListCount
        End If
    End If
End If

If Shift = 0 Or MultiSelect = False Then  'Normal type - single item selection
    If bSelected = iSelected Then GoTo ending
    For ICNT = 1 To ListCount
        If ICNT = iSelected Then
            TVDispInfo(ICNT).Selected = True
            TopLine = Int(Y \ 22) * 22
            
        Else
            TVDispInfo(ICNT).Selected = False
        End If
    Next ICNT
    bSelected = iSelected
    sSelected = iSelected
    
ElseIf Shift = 1 Then   'Shift selection
    If bSelected < 1 Then
        For ICNT = 1 To ListCount
            If ICNT = iSelected Then
                TVDispInfo(ICNT).Selected = True
            Else
                TVDispInfo(ICNT).Selected = False
            End If
        Next ICNT
        bSelected = iSelected
    Else
        If bSelected > iSelected Then
            For ICNT = 1 To ListCount
                If ICNT >= iSelected And ICNT <= bSelected Then
                    TVDispInfo(ICNT).Selected = True
                Else
                    TVDispInfo(ICNT).Selected = False
                End If
            Next ICNT
        Else
            For ICNT = 1 To ListCount
                If ICNT <= iSelected And ICNT >= bSelected Then
                    TVDispInfo(ICNT).Selected = True
                Else
                    TVDispInfo(ICNT).Selected = False
                End If
            Next ICNT
        End If
    End If

ElseIf Shift = 2 Then ' Ctrl selection
    TVDispInfo(iSelected).Selected = Not TVDispInfo(iSelected).Selected
    bSelected = iSelected
End If

List1.Clear

For ICNT = 1 To ListCount
    If TVDispInfo(ICNT).Selected = True And TVDispInfo(ICNT).Extension <> "#FOLDER" Then List1.AddItem TVDispInfo(ICNT).name
Next ICNT

'RaiseEvent FilePreSelect(List1)

UserControl_Resize 'draw selections
ending:

If GetFileTop(ICNT) <= Y - iTop And GetFileTop(ICNT + 1) > Y - iTop Then
End If

If dView > 0 Or X > picDetails.Width + 2 Then
    If Button = vbRightButton Then
    
        If TVDispInfo(iSelected).ShowDetails = True And GetFileTop(ICNT) > PicView.Height - selDetHeight - iTop Then
            iTop = PicView.Height - selDetHeight - GetFileTop(ICNT)
            Resize
        End If

        GetWindowRect UserControl.hwnd, Rec
        
        mx = (Rec.Left + X + 1) * Screen.TwipsPerPixelX
        my = (Rec.Top + Y + 1) * Screen.TwipsPerPixelY
        
        PicMnu.Top = my
        PicMnu.Left = mx
        'SetUpMnu iSelected
        PicMnu.Top = -PicMnu.Height
        
        PicMnu.Visible = True
    End If
End If
Dim isSelcted As Boolean
isSelcted = False
For ICNT = 1 To ListCount
    If TVDispInfo(ICNT).Selected = True And Len(TVDispInfo(ICNT).FullName) > 3 And TVDispInfo(ICNT).FullName <> "#MYCOMPUTER" Then
        isSelcted = True
    End If
Next ICNT
'RaiseEvent DeletableItemSelected(isSelcted)

End Sub
Private Function GetFileTop(index As Long) As Long
On Error Resume Next
'here we get distance of the line from the top
Dim ICNT As Integer
GetFileTop = 0
For ICNT = 1 To index - 1
    If TVDispInfo(ICNT).ShowDetails = True Then
        GetFileTop = GetFileTop + selDetHeight
    Else
        GetFileTop = GetFileTop + SelHeight
    End If
Next ICNT



End Function
Private Sub PicView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ICNT As Long
Dim i As Long
Btn.Visible = False
PicView.Refresh
If dView = 0 And Button = 0 Then
    If X < picDetails.Width + 1 And X > 1 Then
        i = 0
        For ICNT = 1 To ListCount
            If GetFileTop(ICNT) <= Y - iTop And GetFileTop(ICNT + 1) > Y - iTop Then
                i = ICNT
                Exit For
            End If
        Next ICNT
        Dim sh As Integer
        If TVDispInfo(i).ShowDetails = False Then
            sh = SelHeight - 4
        Else
            sh = selDetHeight - 4
        End If
        
        If GetFileTop(i) + 2 + iTop <= Y And GetFileTop(i) + 2 + iTop + sh >= Y Then
        
        Else
            Exit Sub
        End If
        
        If i > 0 Then
            PicBuff3.ForeColor = vbHighlight
            PicBuff3.BackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
            PicBuff3.Cls
            PicBuff3.Picture = LoadPicture("")
            If TVDispInfo(i).ShowDetails = False Then
                PicBuff3.Height = SelHeight - 4
                PicBuff3.Width = picDetails.Width - 2
                
                picDetails.BackColor = PicBuff3.BackColor
                BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                
                PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
                PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
                PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
                PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)

                Set Btn.NormalImage = PicBuff3.Image
                Set Btn.FocusedImage = PicBuff3.Image
                
                PicBuff3.ForeColor = vbHighlight
                PicBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

                picDetails.BackColor = PicBuff3.BackColor
                BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2), picDetails.Width, picDetails.Height, picDetails.hDC, 0, 0, SRCCOPY
                
                PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
                PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
                PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
                PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)

                Set Btn.PressedImage = PicBuff3.Image
                
           Else
                PicBuff3.Cls
                PicBuff3.Height = selDetHeight - 4
                PicBuff3.Width = picDetails.Width - 2
                
                picHideDetails.BackColor = PicBuff3.BackColor
                BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                
                PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
                PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
                PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
                PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)
                
                Set Btn.NormalImage = PicBuff3.Image
                Set Btn.FocusedImage = PicBuff3.Image
                
                PicBuff3.ForeColor = vbHighlight
                PicBuff3.BackColor = AlphaBlend(AlphaBlend(vbHighlight, AlphaBlend(vbHighlight, vbWindowBackground, 70), 128), AlphaBlend(vbHighlight, vbWindowBackground, 70), 128)

                picHideDetails.BackColor = PicBuff3.BackColor
                BitBlt PicBuff3.hDC, (PicBuff3.Width - picDetails.Width) / 2, Int((PicBuff3.Height - picDetails.Height) / 2) + 1, picDetails.Width, picHideDetails.Height, picHideDetails.hDC, 0, 0, SRCCOPY
                
                PicBuff3.Line (0, 0)-(PicBuff3.Width, 0)
                PicBuff3.Line (0, PicBuff3.Height - 1)-(PicBuff3.Width, PicBuff3.Height - 1)
                PicBuff3.Line (0, 0)-(0, PicBuff3.Height - 1)
                PicBuff3.Line (PicBuff3.Width - 1, 0)-(PicBuff3.Width - 1, PicBuff3.Height - 1)

                Set Btn.PressedImage = PicBuff3.Image
            End If
                
                
                Btn.Top = GetFileTop(i) + 2 + iTop
                Btn.Left = 2
                Btn.Tag = i
                Btn.Visible = True
                Btn.SetFocus
        
        End If
    Else
        'PicBtnMnu.visible = False
    End If
End If
If PicRename.Visible = False Then PicView.SetFocus

End Sub


Private Sub picScr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent ButtonDown(False)
If Button = vbLeftButton Then
    If Y >= picScroller.Top And Y < picScroller.Top + picScroller.Height Then
        MouseY = -picScroller.Top + Y
    Else
        MouseY = picScroller.Height / 2
        picScr_MouseMove Button, Shift, X, Y
    End If
End If

End Sub


Private Sub picScr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Y - MouseY <= BtnUp.Height Then
        picScroller.Top = BtnUp.Height
        picScr.Refresh
        
        If Not iTop = 0 Then
            iTop = 0
            Generate
        End If
        
    ElseIf Y - MouseY >= PicView.Height - BtnDown.Height - picScroller.Height Then
        picScroller.Top = PicView.Height - BtnDown.Height - picScroller.Height
        picScr.Refresh
        
        If Not iTop = PicView.Height - CheckHeight Then
            iTop = PicView.Height - CheckHeight
            Generate
        End If
        
    Else
        picScroller.Top = Y - MouseY
        picScr.Refresh
        
        If Not iTop = Int((picScroller.Top - BtnUp.Height) * (PicView.Height - CheckHeight) / ((PicView.Height - BtnUp.Height - BtnDown.Height) - picScroller.Height)) Then
            iTop = Int((picScroller.Top - BtnUp.Height) * (PicView.Height - CheckHeight) / ((PicView.Height - BtnUp.Height - BtnDown.Height) - picScroller.Height))
            Generate
        End If
        
    End If
End If
End Sub


Private Sub picScrH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If X >= picScrollerH.Left And X < picScrollerH.Left + picScrollerH.Width Then
        MouseX = -picScrollerH.Left + X
    Else
        MouseX = picScrollerH.Width / 2
        picScrH_MouseMove Button, Shift, X, Y
    End If
End If

End Sub

Private Sub picScrH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If X - MouseX <= BtnLeft.Width Then
        picScrollerH.Left = BtnLeft.Width
        picScrH.Refresh
        
        If Not iLeft = 0 Then
            iLeft = 0
            Generate
        End If
    ElseIf X - MouseX >= PicView.Width - BtnRight.Width - picScrollerH.Width Then
        picScrollerH.Left = PicView.Width - BtnRight.Width - picScrollerH.Width
        picScrH.Refresh
        
        If Not iLeft = PicView.Width - ColumnCount * picBuff.Width Then
            iLeft = PicView.Width - ColumnCount * picBuff.Width
            Generate
        End If
    Else
        picScrollerH.Left = X - MouseX
        picScrH.Refresh

        If Not iLeft = Int((picScrollerH.Left - BtnLeft.Width) / ((PicView.Width - BtnLeft.Width - BtnRight.Width - picScrollerH.Width) / (PicView.Width - (ColumnCount * ColumnWidth)))) Then
            iLeft = Int((picScrollerH.Left - BtnLeft.Width) / ((PicView.Width - BtnLeft.Width - BtnRight.Width - picScrollerH.Width) / (PicView.Width - (ColumnCount * ColumnWidth))))
            Generate
        End If
    End If
End If

End Sub


Private Sub picScroller_Resize()
LN.Y1 = picScroller.Height - 1
LN.Y2 = picScroller.Height - 1

End Sub


Private Sub picScrollerH_Resize()
  lnH.X1 = picScrollerH.Width - 1
  lnH.X2 = picScrollerH.Width - 1

End Sub


Private Sub tmrScr_Timer()
Dim cc As Integer
Dim gTop As Long
tmrScr.Interval = 25

If tmrScr.Tag = "L" Or tmrScr.Tag = "R" Then GoTo LeFTrIGHT
cc = 15
If tmrScr.Tag = "D" Then
    If iTop > PicView.Height - CheckHeight - cc Then
        gTop = iTop - cc
    Else
        gTop = PicView.Height + CheckHeight
    End If
ElseIf tmrScr.Tag = "U" Then
    If iTop < -cc Then
        gTop = iTop + cc
    Else
        gTop = 0
    End If
End If

If gTop <> iTop Then
    iTop = gTop
    UserControl_Resize
End If
    
Exit Sub

LeFTrIGHT:
cc = 20
Dim gLeft As Long
If tmrScr.Tag = "L" Then
    If iLeft < -cc Then
        gLeft = iLeft + cc
    Else
        gLeft = 0
    End If
ElseIf tmrScr.Tag = "R" Then
    If iLeft > PicView.Width - ColumnCount * picBuff.Width - cc Then
        gLeft = iLeft - cc
    Else
        gLeft = PicView.Width - ColumnCount * picBuff.Width
    End If
End If

If iLeft <> gLeft Then
    iLeft = gLeft
    UserControl_Resize
End If


End Sub

Private Sub UserControl_Initialize()
  SelBorderColor = vbHighlight
  SelBackColor = AlphaBlend(vbHighlight, vbWindowBackground, 70)
  SelForeColor = vbBlack
  oBackColor = vbWhite
  oForeColor = vbBlack
  FontSize = 10

  PicView.Top = 1
  PicView.Left = 1
  
  picScr.Top = 1
  picScrH.Left = 1
  
  BtnUp.Left = 0
  BtnUp.Top = 0
  
  DrawScrollers
  UserControl.BackColor = AlphaBlend(AlphaBlend(vbButtonFace, vbWhite, 170), vb3DDKShadow, 70)
  BtnDown.Left = 0
  DoNotGenerate = True


  SetWindowLong picShdw.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
  SetParent picShdw.hwnd, 0

  SelHeight = 22      'NormalviewHeight
  selDetHeight = 96   '"Extendet" dView Mode

  
  picScr.BackColor = AlphaBlend(vb3DHighlight, vbButtonFace, 128)
  picScrH.BackColor = AlphaBlend(vb3DHighlight, vbButtonFace, 128)

  ListCount = 0
  ReDim Preserve TVDispInfo(ListCount)
  

End Sub


Private Sub UserControl_Resize()
  DoNotGenerate = True
  If dView = 0 Then
    PicView.Height = UserControl.Height / Screen.TwipsPerPixelY - 2
    picScrH.Visible = False
    picScr.Left = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 1
    picScr.Height = PicView.Height
    
    'On Error ResuUserControl Next
    If CheckHeight > UserControl.Height / Screen.TwipsPerPixelY - 2 Then
        If iTop + CheckHeight < PicView.Height Then iTop = PicView.Height - CheckHeight
        SetScroller
        BtnDown.Top = picScr.Height - BtnDown.Height
        picScr.Visible = True
        PicView.Width = UserControl.Width / Screen.TwipsPerPixelX - picScr.Width - 2
    Else
        iTop = 0
        picScr.Visible = False
        PicView.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    End If
  ElseIf dView = 1 Or dView = 2 Then
    PicView.Width = UserControl.Width / Screen.TwipsPerPixelX - 2
    
    picScr.Visible = False
    picScrH.Top = UserControl.Height / Screen.TwipsPerPixelY - picScrH.Height - 1
    picScrH.Width = PicView.Width
  End If
  DoNotGenerate = False
  Generate
  RaiseEvent Resize
End Sub

'== Node count (Total/Children)
Public Property Get NodeCount() As Long
    NodeCount = ListCount
End Property
Public Sub Refresh()

DrawBackgrounds

End Sub
Public Sub Resize()
SetScroller
SetScrollerH
Generate

End Sub
