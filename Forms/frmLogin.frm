VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H80000014&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "User's Login"
   ClientHeight    =   2850
   ClientLeft      =   2850
   ClientTop       =   3375
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1683.874
   ScaleMode       =   0  'User
   ScaleWidth      =   5506.917
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Password"
      Text            =   "*******"
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   0
      Tag             =   "UserName"
      Top             =   1080
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   5535
      TabIndex        =   9
      Top             =   2160
      Width           =   5535
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   8
      Top             =   742
      Width           =   9375
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   391
      TabIndex        =   5
      Top             =   0
      Width           =   5865
      Begin VB.Image Image2 
         Height          =   660
         Left            =   0
         Picture         =   "frmLogin.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plese enter your username and password in the space provided bellow to login."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   420
         Left            =   840
         TabIndex        =   7
         Top             =   510
         Width           =   5760
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User's Log-in"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   840
         TabIndex        =   6
         Top             =   180
         Width           =   1875
      End
   End
   Begin VB.PictureBox bgHWND 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   -3090
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   -120
      Width           =   15
   End
   Begin lvButton.lvButtons_H cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   4440
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "Log-In"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   330
      Left            =   3480
      TabIndex        =   11
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attemp :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label lblatemp 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   780
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsData               As ADODB.Recordset
Dim rsTotal              As ADODB.Recordset
Dim attempt As Integer
 
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
If is_empty(txtUserName, True) Then Exit Sub
If is_empty(txtPassword, True) Then Exit Sub
Set rsData = New ADODB.Recordset
Set rsData = CN.Execute("Select * from tblAccount where username='" & txtUserName & "'")
    With rsData
        If .BOF = True And .EOF = True Then
             verify_attempt
             attempt = attempt + 1
             lblatemp.Caption = attempt
              MsgBox "Invalid Username!" & vbCrLf & "Please type the correct Username!", vbExclamation, "Warning...."
             txtUserName.SetFocus
             SendKeys "{Home}+{End}"
        Else
               If txtPassword = !Password Then
                CurrUser.UserNAME = !CompleteName
                CurrUser.UserPK = !UserCode
               ' CurrUser.UserisAdmin = !Usertype
               ' CurrUser.IsDate = FormatDateTime(Date, vbShortDate)
               ' CurrUser.Status = !Usertype
                MDIForm1.lblCurrentUser = CurrUser.UserNAME
               ' Call SaveUserLogs(CurrUser.UserNAME, "In")
                Unload Me
            Else
                verify_attempt
                attempt = attempt + 1
                lblatemp.Caption = attempt
                MsgBox "Invalid password!" & vbCrLf & "Please type the correct password!", vbExclamation, "Warning...."
                txtPassword.SetFocus
                SendKeys "{Home}+{End}"
            End If
        End If
    End With

End Sub

Private Sub cmdLog_Click()

End Sub

Private Sub Form_Load()
attemp = 0
End Sub
 Private Sub verify_attempt()
If attempt = 3 Then
    MsgBox "This will terminate the applicatin" & vbCrLf & "You already used all attempt", vbCritical, "System Information"
     End
 End If
End Sub

 
 

Private Sub txtPassword_Click()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtPassword_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

 
Private Sub txtUserName_Click()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtUserName_GotFocus()
 SendKeys "{Home}+{End}"
End Sub
 

'Public Sub Total()
'Set rsTotal = New ADODB.Recordset
'rsTotal.Open "Select SUM(TotalAmount) as TotalSales FROM tblSalesFuel WHERE Createdby='" & CurrUser.UserNAME & "' and DateCreated='" & FormatDateTime(Date, vbShortDate) & "'", CN, adOpenStatic, adLockPessimistic
'CurrUser.DailySales = rsTotal!totalSales
'Text1.Text = CurrUser.DailySales
'End Sub

