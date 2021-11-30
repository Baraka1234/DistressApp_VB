VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distress Management System"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Home.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMControl1 
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "My Recordings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   4
      Top             =   5040
      Width           =   6735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send My Device Coordinates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View My Emergency Contact"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register Emergency Contact"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Distress Management System"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
Call ConnectStringAc
    
    Set recordsetsearch = New ADODB.Recordset
     
        recordsetsearch.LockType = adLockOptimistic
        recordsetsearch.CursorType = adOpenDynamic
        recordsetsearch.CursorLocation = adUseClient
        
        recordsetsearch.Open "Select * from details", ConnectStringAc

If recordsetsearch.EOF Then
MsgBox "No Contact Saved. Please Register Emergency Contact", vbCritical, "No Contact"
Exit Sub

Else

Form3.Text1.FontBold = True
Form3.Text1.FontName = "Arial"
Form3.Text1.FontSize = 12

Form3.Text2.FontBold = True
Form3.Text2.FontName = "Arial"
Form3.Text2.FontSize = 12

Form3.Text3.FontBold = True
Form3.Text3.FontName = "Arial"
Form3.Text3.FontSize = 12

Form3.Text4.FontBold = True
Form3.Text4.FontName = "Arial"
Form3.Text4.FontSize = 12

Form3.Text5.FontBold = True
Form3.Text5.FontName = "Arial"
Form3.Text5.FontSize = 12

Form3.Text6.FontBold = True
Form3.Text6.FontName = "Arial"
Form3.Text6.FontSize = 12

Form3.Text1.Text = recordsetsearch.Fields("username")
Form3.Text2.Text = recordsetsearch.Fields("useraddress")
Form3.Text3.Text = recordsetsearch.Fields("userphone")
Form3.Text4.Text = recordsetsearch.Fields("contactname")
Form3.Text5.Text = recordsetsearch.Fields("contactaddress")
Form3.Text6.Text = recordsetsearch.Fields("contactphone")


Me.Hide
Form3.Show
End If
End Sub

Private Sub Command3_Click()
Call ConnectStringAc
    
    Set recordsetsearch = New ADODB.Recordset
     
        recordsetsearch.LockType = adLockOptimistic
        recordsetsearch.CursorType = adOpenDynamic
        recordsetsearch.CursorLocation = adUseClient
        
        recordsetsearch.Open "Select * from details", ConnectStringAc

If recordsetsearch.EOF Then
MsgBox "No Contact Saved. Please Register Emergency Contact", vbCritical, "No Contact"
Exit Sub

Else

Form4.Label7.Caption = recordsetsearch.Fields("username")
Form4.Label8.Caption = recordsetsearch.Fields("contactname")
Form4.Label9.Caption = recordsetsearch.Fields("userphone")
Form4.Label10.Caption = recordsetsearch.Fields("contactphone")
Form4.WebBrowser1.Navigate ("www.google.com/maps")


Me.Hide
Form4.Show

End If
End Sub

Private Sub Command4_Click()
  With MMControl1
        .DeviceType = "WaveAudio"
        .FileName = App.Path & "\NewAudio1.WAV"
        .RecordMode = mciRecordOverwrite
        'mciRecordOverwrite
        .UpdateInterval = 2
        .Command = "Open"
        .Command = "Play"
    End With
End Sub
