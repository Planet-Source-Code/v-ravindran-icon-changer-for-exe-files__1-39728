VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "iconpatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmD2 
      Left            =   3600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.ico"
      Filter          =   "ico"
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CmD1 
      Left            =   3000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.exe"
      Filter          =   "exe"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SELECTED"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Text            =   "C:\WINDOWS\DESKTOP\SHIELD.ICO"
      Top             =   720
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SELECTED"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7223
      _Version        =   393217
      BackColor       =   12640511
      ScrollBars      =   3
      TextRTF         =   $"iconpatch.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "c:\windows\desktop\exe.exe"
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "WRITING STASTUS..."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "EXE STRINGS..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "SELECT AN ICON FILE TO BE WRITTEN IN EXE..."
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT YOUR EXE FILE TO CHANGE  THE ICON..."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim WRITESTR As String
Dim WriteDllPos(0 To 100) As Long
Dim GetStartVal As Long
Dim writestring As String
Dim lenght1, m As Integer
Private Sub Command1_Click()
Dim hkl, WRT As Integer
Dim kp1, kp2, kp3, kp4, kp5, KP6, KP7, KP8, KP9, loc As Double
hkl = 0
Dim Countee As String

 

 Dim i As Double, iCount As Long

    'Initialize our variables
    WRT = 0
    iCount = 0
    i = 1

    
ProgressBar1.Max = Len(writestring)
ProgressBar1.Min = 1
For i = 1 To Len(writestring)

If i = 1 And hkl = 1 Then
GoTo juu:
End If

If i = 1 Then
hkl = 1

End If



ProgressBar1.Value = i
          
           
               DoEvents
                i = InStr(i, writestring, Chr((HexToDec("28"))), vbTextCompare)
             DoEvents
            kp4 = i
            If i Then
                DoEvents
                i = InStr(i, writestring, Chr((HexToDec("00"))), vbTextCompare)
              DoEvents
            
              kp1 = i
           
               If i = kp4 + 1 Then
               DoEvents
               i = InStr(i + 1, writestring, Chr((HexToDec("00"))), vbTextCompare)
             
               kp2 = i
              
     DoEvents
 If i = kp1 + 1 Then
  
   
   DoEvents
   i = InStr(i + 1, writestring, Chr((HexToDec("00"))), vbTextCompare)
     
    kp3 = i
    
  
   DoEvents
   If i = kp2 + 1 Then
    DoEvents
     i = InStr(i + 1, writestring, Chr((HexToDec("20"))), vbTextCompare)
     DoEvents
     kp5 = i
     
     If i = kp3 + 1 Then
     i = InStr(i + 1, writestring, Chr((HexToDec("00"))), vbTextCompare)
     DoEvents
     KP6 = i
      
      If i = kp5 + 1 Then
         i = InStr(i + 1, writestring, Chr((HexToDec("00"))), vbTextCompare)
     DoEvents
     KP7 = i
    
    If i = KP6 + 1 Then
       i = InStr(i + 1, writestring, Chr((HexToDec("00"))), vbTextCompare)
     DoEvents
     KP8 = i
  
  If i = KP7 + 1 Then
     i = InStr(i + 1, writestring, Chr((HexToDec("40"))), vbTextCompare)
     DoEvents
    
  
  WRT = 1
  
  FileNumber = FreeFile
  
  
  Open Text1.Text For Binary As FileNumber
  
  Put FileNumber, kp4, "(" & WRITESTR
  
  Close FileNumber
          
        
     End If
           End If
         
  End If
  End If
  End If
  
   End If
            End If
                
             
                      iCount = iCount + 1
                
               End If

 
Next i
        sCount = iCount
   
Close FileNumber

juu:

         sCount = 0

If WRT = 1 Then
MsgBox "Icon changed, Click refrish to view the change"
Else
MsgBox "Finished... unable to change the icon..."
End If
End Sub

Public Sub Command2_Click()
GetStartVal = 1
 FileNumber = FreeFile
 Open Text1.Text For Binary As FileNumber
 
 writestring = Space(FileLen(Text1.Text))
Get FileNumber, GetStartVal, writestring

RichTextBox1.Text = writestring
Close FileNumber
End Sub

Public Sub Command3_Click()
Dim kl As Integer
 Dim FileInformation As FILE_INFORMATION
 Call GetFileInformation(Text4, FileInformation, True)
 If CInt(FileInformation.nFileSize) = 2238 Then
 MsgBox "You have selected 8 Bit 256 colours icon, if you are very sure that the icon in the exe file is of the same pattern then you can go ahead,else select 4 Bit 16 Colour icons which are always safe"
 
 End If
 If CInt(FileInformation.nFileSize) = 3266 Then
 MsgBox "You have selected 24 Bit True colour icon, if you are very sure that the icon in the exe file is of the same pattern then you can go ahead,else select 4 Bit 16 Colour icons which are always safe"
 
 End If
If CInt(FileInformation.nFileSize) = 326 Then
 MsgBox "You have selected 1 Bit B/W icon, if you are very sure that the icon in the exe file is of the same pattern then you can go ahead,else select 4 Bit 16 Colour icons which are always safe"
 
 End If
 If CInt(FileInformation.nFileSize) = 766 Then
 MsgBox "You have selected 4 Bit 16 colour icon, "
  End If
Picture1.Picture = LoadPicture(Text4.Text)
FileNumber = FreeFile
GetStartVal = 1
Open Text4.Text For Binary As FileNumber
WRITESTR = Space(FileLen(Text4.Text))
DoEvents
Get FileNumber, GetStartVal, WRITESTR
DoEvents
Close FileNumber

kl = InStr(1, WRITESTR, "(", vbTextCompare)

GetStartVal = kl + 1
Open Text4.Text For Binary As FileNumber
WRITESTR = Space(FileLen(Text4.Text))
DoEvents
Get FileNumber, GetStartVal, WRITESTR
DoEvents
Close FileNumber

End Sub

Private Sub Command4_Click()
CmD2.ShowOpen
Text4.Text = CmD2.FileName
End Sub

Private Sub Command5_Click()
CmD1.ShowOpen

Text1.Text = CmD1.FileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub
