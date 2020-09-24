VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   ControlBox      =   0   'False
   Icon            =   "WARNING.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   1560
      Picture         =   "WARNING.frx":27A2
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   5280
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "TESTED ON POPULAR EXE'S"
      Top             =   4200
      Width           =   5295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "TESTED ON POPULAR EXE'S"
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      Height          =   5145
      Left            =   0
      Top             =   0
      Width           =   6815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FAILURE LIST..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SUCCESS LIST..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"WARNING.frx":370C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem "IEXPLORE.EXE-Ver 5.50.4134.100 internet explorer"
Combo1.AddItem "WMPLAYER.EXE-Ver 7.0.0.1440 media player"
Combo1.AddItem "MSIMN.EXE-Ver 5.50.4133.2400 outlook express"
Combo1.AddItem "NAVW32.EXE-Ver 7.00.00.51 norton antivirus"
Combo1.AddItem "PINBALL.EXE-Ver 4.90.3000.0 cinimatronics"
Combo1.AddItem "WZSEP32.EXE-Ver 1.0.0.0 winzip,nico mac computing"
Combo1.AddItem "music converter.exe-Ver 4.0.0.28 dBPowerAMP"
Combo1.AddItem "Lots more......."
Combo2.AddItem "WORDPAD.EXE-Ver 5.0.1691.1 word pad"
Combo2.AddItem "TELNET.EXE-Ver 4.90.3000.0 tel net"
Combo2.AddItem "NOTEPAD.EXE-Ver 4.90.0.3000 note pad"
Combo2.AddItem "PROGMAN.EXE-Ver 4.90.0.3000 program manager"

End Sub

