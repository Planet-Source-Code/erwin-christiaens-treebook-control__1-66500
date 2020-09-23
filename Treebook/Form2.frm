VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "FrmTest"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Image_DB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DataField       =   "Photo"
      DataSource      =   "DataPictures"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1365
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7305
      Width           =   405
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load DB"
      Height          =   330
      Left            =   7125
      TabIndex        =   8
      Top             =   180
      Width           =   780
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Select"
      Height          =   330
      Left            =   5265
      TabIndex        =   7
      Top             =   150
      Width           =   690
   End
   Begin VB.CommandButton Command4 
      Caption         =   "List"
      Height          =   345
      Left            =   4725
      TabIndex        =   6
      Top             =   150
      Width           =   510
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Collapse"
      Height          =   300
      Left            =   3780
      TabIndex        =   5
      Top             =   330
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Expand"
      Height          =   315
      Left            =   3780
      TabIndex        =   4
      Top             =   30
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   765
      TabIndex        =   3
      Top             =   195
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   30
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   7305
      Width           =   1215
   End
   Begin Project1.TreeBook TreeBook 
      Height          =   5355
      Left            =   90
      TabIndex        =   1
      Top             =   705
      Width           =   7935
      _extentx        =   13996
      _extenty        =   9446
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0734
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0ACE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Graphics
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020


Private Sub Command1_Click()
  Dim Details As String
  
  Details = "722 Moss Bay Blvd."
  Details = Details & "|98033 Kirkland"
  Details = Details & "|Tel.  1-239-368-7924"
  Details = Details & "|gsm   1-239-368-7924"
  
    With TreeBook
      pNode = .AddNode("Key", Text1.Text, Details, Picture1.Picture, ImageList1.ListImages(4).Picture)
      '.Expand pNode
    End With
End Sub

Private Sub Command2_Click()
TreeBook.ExpandAll
End Sub

Private Sub Command3_Click()
TreeBook.CollapseAll
End Sub

Private Sub Command4_Click()
  For i = 1 To TreeBook.NodeCount
    MsgBox TreeBook.GetCaption(i)
  Next i
End Sub

Private Sub Command5_Click()
  TreeBook.SelectedNode = 3
End Sub

Private Sub Command6_Click()
  Dim Rs                      As New Recordset
  Dim Rs_Detail               As New Recordset
  Dim Details                 As String
  
  TreeBook.Clear
  Picture1.Cls
  Rs.CursorLocation = adUseClient
  Rs.Open "SELECT AddressBook.ID, AddressBook.Title, AddressBook.Type, AddressBook.Photo FROM AddressBook WHERE ((AddressBook.Type)= '1') ORDER BY AddressBook.Title;", CN, adOpenKeyset, adLockOptimistic
  Set Image_DB.DataSource = Rs
  Do While Not Rs.EOF
    Details = ""
    Rs_Detail.CursorLocation = adUseClient
    Rs_Detail.Open "SELECT AddressBookFields.* FROM AddressBookFields WHERE (((AddressBookFields.OwnerID)=" & Rs.Fields("ID") & ")) ORDER BY AddressBookFields.RNo;", CN, adOpenStatic, adLockOptimistic

    Do While Not Rs_Detail.EOF
      m_field = Left$(Rs_Detail.Fields("Fieldname").value + Space(15), 15)
      m_content = Rs_Detail.Fields("Contents").value
      Details = Details & m_field & Space(8) & m_content
      Details = Details & "|"
      
      Rs_Detail.MoveNext
    Loop
        '-- Draw Picture1
        BitBlt Picture1.hDC, 0, 0, Image_DB.Width, Image_DB.Height, Image_DB.hDC, 0, 0, SRCCOPY
        '-- Draw  border
        Picture1.Line (0, 0)-(Picture1.ScaleWidth - 1, Picture1.ScaleHeight - 1), , B

    DoEvents
    With TreeBook
        pNode = .AddNode(Rs.Fields("ID").value, Rs.Fields("Title").value, Details, Picture1.Image, ImageList1.ListImages(4).Picture)
    End With
    Rs_Detail.Close
    Set Rs_Detail = Nothing
    Rs.MoveNext
  Loop
  Rs.Close
End Sub


Private Sub Form_Click()
  Unload Form3
End Sub

Private Sub Form_Load()
  CN.CursorLocation = adUseClient
  CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\organizer.mdb" & ";Persist Security Info=False"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Unload Form3
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  TreeBook.Width = Me.ScaleWidth - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CN.Close
  Unload Form3
  Unload Me
End Sub

Private Sub TreeBook_ButtonDown(Clicked As Boolean)
  If Clicked = True Then
    If Form3.Visible = True Then Unload Form3
    
    Form3.Top = Form2.Top + Form2.TreeBook.Top + ((TreeBook.TopLine + 96) * Screen.TwipsPerPixelY)
    Form3.Width = TreeBook.Listwidth - 60
    Form3.Left = Form2.Left + TreeBook.Left + 200
    Form3.Label1.Caption = Form2.TreeBook.ListText
    Form3.Show
    Form3.ZOrder 0
  Else
    Unload Form3
  End If
End Sub

Private Sub TreeBook_DbClick()
  'MsgBox TreeBook.ListIndex
  Unload Form3
  Form4.Show vbModal
End Sub

Private Sub TreeBook_NodeClick(ByVal Node As Long)
  'MsgBox TreeBook.GetCaption(Node)
End Sub


