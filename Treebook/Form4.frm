VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   360
      Left            =   8085
      TabIndex        =   30
      Top             =   5805
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   6540
      TabIndex        =   29
      Top             =   5805
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   4995
      TabIndex        =   28
      Top             =   5820
      Width           =   1350
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   3750
      TabIndex        =   27
      Text            =   "0"
      Top             =   5265
      Width           =   630
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   3735
      TabIndex        =   26
      Top             =   4905
      Width           =   1305
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   795
      Index           =   6
      Left            =   5610
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   4050
      Width           =   3720
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   795
      Index           =   5
      Left            =   1785
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   4035
      Width           =   3720
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   4
      Left            =   1800
      TabIndex        =   23
      Top             =   2100
      Width           =   7515
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   3
      Left            =   1800
      TabIndex        =   22
      Top             =   1785
      Width           =   7515
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   1800
      TabIndex        =   21
      Top             =   1470
      Width           =   7515
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   1800
      TabIndex        =   20
      Top             =   1155
      Width           =   7515
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   165
      ScaleHeight     =   1395
      ScaleWidth      =   1185
      TabIndex        =   19
      Top             =   3765
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delete empty fields:"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   6
      Left            =   45
      TabIndex        =   18
      Top             =   5850
      Value           =   1  'Checked
      Width           =   2130
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Remind before:"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   1770
      TabIndex        =   17
      Top             =   5280
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date of Birth:"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   1770
      TabIndex        =   16
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Phone (Work)"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   225
      TabIndex        =   15
      Top             =   2115
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Phone (Home)"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   225
      TabIndex        =   14
      Top             =   1815
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "First name"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   225
      TabIndex        =   13
      Top             =   1500
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Last name"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   12
      Top             =   1170
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.TextBox TxtDatafield 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1425
      TabIndex        =   1
      Top             =   120
      Width           =   7920
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   165
      Top             =   510
      Width           =   9150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      Height          =   195
      Index           =   10
      Left            =   5970
      TabIndex        =   11
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   195
      Index           =   9
      Left            =   2145
      TabIndex        =   10
      Top             =   3840
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load photo"
      Height          =   195
      Index           =   8
      Left            =   165
      TabIndex        =   9
      Top             =   5355
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   195
      Index           =   7
      Left            =   1860
      TabIndex        =   8
      Top             =   915
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field name:"
      Height          =   195
      Index           =   6
      Left            =   225
      TabIndex        =   7
      Top             =   915
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move down"
      Height          =   195
      Index           =   5
      Left            =   6690
      TabIndex        =   6
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move up"
      Height          =   195
      Index           =   4
      Left            =   5520
      TabIndex        =   5
      Top             =   585
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remove field"
      Height          =   195
      Index           =   3
      Left            =   3930
      TabIndex        =   4
      Top             =   585
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add new field"
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   585
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Templates..."
      Height          =   195
      Index           =   1
      Left            =   690
      TabIndex        =   2
      Top             =   585
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   285
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Just to show DubbleClick Treebook
'Form2 shows how to fill in the fields
'This is only a example of what je can do with Vb
'Try to make it better and share it for all Vb users on the planet sourcecode
'Thanks
'
