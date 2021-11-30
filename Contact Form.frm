VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emergency Contact Registration"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Contact Form.frx":0000
   ScaleHeight     =   6480
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   1080
      TabIndex        =   14
      Top             =   5280
      Width           =   7935
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete Contact"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Store Contact"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Emergency Contact Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   1080
      TabIndex        =   1
      Top             =   2880
      Width           =   7935
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Phone No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Full Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User's Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Names:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter A Valid Username", vbCritical, "Can't Save Info"
Text1.SetFocus
Else

If Text3.Text = "" Then
MsgBox "Please Enter A Valid Phone Number", vbCritical, "Can't Save Info"
Text3.SetFocus
Else

If Text4.Text = "" Then
MsgBox "Enter Emergency Contact Name", vbCritical, "Can't Save Info"
Text4.SetFocus
Else

If Text6.Text = "" Then
MsgBox "Enter Emergency Contact Phone", vbCritical, "Can't Save Info"
Text6.SetFocus
Else

Call ConnectStringAc
    
    Set recordsetsearch = New ADODB.Recordset
     
        recordsetsearch.LockType = adLockOptimistic
        recordsetsearch.CursorType = adOpenDynamic
        recordsetsearch.CursorLocation = adUseClient
        
        recordsetsearch.Open "Select * from details", ConnectStringAc

Do Until recordsetsearch.EOF
If recordsetsearch.Fields("username").Value = Text1.Text And recordsetsearch.Fields("userphone").Value = Text3.Text Then

recordsetsearch.Fields("username").Value = Text1.Text
recordsetsearch.Fields("useraddress").Value = Text2.Text
recordsetsearch.Fields("userphone").Value = Text3.Text

recordsetsearch.Fields("contactname").Value = Text4.Text
recordsetsearch.Fields("contactaddress").Value = Text5.Text
recordsetsearch.Fields("contactphone").Value = Text6.Text
recordsetsearch.Update
MsgBox "Contact Updated Successfully", vbInformation, "Updated"

End If
Exit Sub
'Else: recordsetsearch.movenexx


Loop

recordsetsearch.AddNew
recordsetsearch.Fields("username").Value = Text1.Text
recordsetsearch.Fields("useraddress").Value = Text2.Text
recordsetsearch.Fields("userphone").Value = Text3.Text

recordsetsearch.Fields("contactname").Value = Text4.Text
recordsetsearch.Fields("contactaddress").Value = Text5.Text
recordsetsearch.Fields("contactphone").Value = Text6.Text
recordsetsearch.Update


   Dim objXML As Object
    Dim mymessage As String
    Dim myusername As String
    Dim mypassword As String
    Dim mysender As String
    Dim URL As String

Recipient = "234" + Mid$(Text3.Text, 2)
mysender = "Distress"
mymessage = "Dear " + Text1.Text + ", You Have Successfully Registered Your Emergency Contact. Thank you for using this service"

URL = "https://www.bulksmsnigeria.com/api/v1/sms/create?api_token=ckSXSPDKIuR3Qos9WQ92N4OjEB8gEOCxyvdLeeP71SzZHaxlO7BzSvFyVYGP&from=" & mysender & "&to=" & Recipient & "&body=" & mymessage & "&dnd= 2"

'URL = "http://vencube.com/api/no_reseller/bulksms.aspx?username=" & myusername & "&password=" & mypassword & "&sender=" & mysender & "&to=" & Recipient & "&message=" & URLEncode(mymessage)
Set objXML = CreateObject("Microsoft.XMLHTTP")
objXML.Open "POST", URL, False
objXML.send
MsgBox "Contact Captured Successfully", vbInformation, "Contact Saved"


End If
End If
End If
End If


End Sub

Private Sub Command2_Click()
confirm = MsgBox("Are you sure you want to delete your emergency contact? This Action Cannot Be Undone", vbQuestion + vbYesNo, "Delete Contact")
If confirm = vbYes Then

Call ConnectStringAc
    
    Set recordsetsearch = New ADODB.Recordset
     
        recordsetsearch.LockType = adLockOptimistic
        recordsetsearch.CursorType = adOpenDynamic
        recordsetsearch.CursorLocation = adUseClient
        
        recordsetsearch.Open "Delete * from details", ConnectStringAc
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Form1.Show
End Sub
