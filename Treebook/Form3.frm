VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6705
      Picture         =   "Form3.frx":4CF62
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   45
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6450
      Picture         =   "Form3.frx":4D2A4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   45
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   1755
      TabIndex        =   14
      Top             =   2055
      Width           =   75
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   225
      Index           =   5
      Left            =   3570
      TabIndex        =   13
      Top             =   2055
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   1755
      TabIndex        =   12
      Top             =   1635
      Width           =   75
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   225
      Index           =   4
      Left            =   3570
      TabIndex        =   11
      Top             =   1635
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   225
      Index           =   3
      Left            =   3585
      TabIndex        =   10
      Top             =   1395
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   225
      Index           =   2
      Left            =   3585
      TabIndex        =   9
      Top             =   1155
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   225
      Index           =   1
      Left            =   3585
      TabIndex        =   8
      Top             =   915
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   225
      Index           =   0
      Left            =   3585
      TabIndex        =   7
      Top             =   675
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   1770
      TabIndex        =   6
      Top             =   1395
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   1770
      TabIndex        =   5
      Top             =   1155
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1770
      TabIndex        =   4
      Top             =   915
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1770
      TabIndex        =   3
      Top             =   675
      Width           =   75
   End
   Begin VB.Shape Shape1 
      Height          =   1905
      Left            =   0
      Top             =   0
      Width           =   7050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marijke Beerden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   5010
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   255
      Picture         =   "Form3.frx":4D5E6
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const SW_SHOW = 5
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub Picture2_Click()
  Call OnTop(Me, False)
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Rs_Detail                      As New Recordset
  Dim Details                 As String
  Dim key As Long
  
  key = Val(Form2.TreeBook.GetNodeKey(Form2.TreeBook.SelectedNode))
  Details = ""
  Rs_Detail.CursorLocation = adUseClient
  Rs_Detail.Open "SELECT AddressBookFields.* FROM AddressBookFields WHERE (((AddressBookFields.OwnerID)=" & key & ")) ORDER BY AddressBookFields.RNo;", CN, adOpenStatic, adLockOptimistic
  index = 0
  Set Image1.Picture = Form2.TreeBook.GetPhoto(Form2.TreeBook.SelectedNode)
  Do While Not Rs_Detail.EOF
    Label2(index).Caption = "" & Rs_Detail.Fields("Fieldname").value
    Label3(index).Caption = "" & Rs_Detail.Fields("Contents").value
    
    Rs_Detail.MoveNext
    index = index + 1
  Loop
  
  Rs_Detail.Close
  Set Rs_Detail = Nothing
  Call OnTop(Me, True)
  
  
End Sub
' When this property is set, the window is made top most
Public Sub OnTop(frmTop As Form, bSetOnTop As Boolean)
    If bSetOnTop = True Then
      ' Set the window to topmost window
      Call SetWindowPos(frmTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Else
      ' Set the window to not topmost window
      Call SetWindowPos(frmTop.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
End Sub

Private Sub Form_Resize()
  Picture1.Left = Me.ScaleWidth - (Picture1.Width + Picture2.Width + 40)
  Picture2.Left = Me.ScaleWidth - (Picture2.Width + 40)
  Shape1.Width = Me.ScaleWidth - 20
End Sub
