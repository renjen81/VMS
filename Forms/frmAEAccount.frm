VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAEAccount 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Account"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "frmAEAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   60
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   7890
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   515
      TabIndex        =   10
      Top             =   0
      Width           =   7725
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Account"
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
         TabIndex        =   12
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fill all fields or fields with '*' then click 'Save' button to update."
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
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   630
         Width           =   3900
      End
      Begin VB.Image Image2 
         Height          =   1305
         Left            =   -240
         Picture         =   "frmAEAccount.frx":57E2
         Stretch         =   -1  'True
         Top             =   -240
         Width           =   1410
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Tag             =   "UserName"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Password"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Tag             =   "Complete Name"
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   60
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   7410
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   10530
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   330
      Left            =   5280
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "&Save "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   330
      Left            =   6720
      TabIndex        =   9
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   2580
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   4551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilList"
      SmallIcons      =   "ilList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Complete Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Admin"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   5280
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAEAccount.frx":6A37
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   330
      Left            =   4320
      TabIndex        =   19
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "N&ew"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Labels 
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1815
      Width           =   915
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "frmAEAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsData           As ADODB.Recordset
 
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    Set rsData = New ADODB.Recordset
    rsData.Open "Select * From tblAccount", CN, adOpenStatic, adLockPessimistic
    With rsData
        .AddNew
        .Fields("UserCode") = Text4.Text
        .Fields("username") = Text1.Text
        .Fields("Password") = Encode(Text2.Text)
        .Fields("CompleteName") = Text3.Text
        .Fields("Type") = changeYNValue(Check1.Value)
        .Fields("DateAdded") = Format(Date, "MMM-dd-yyyy")
        .Update
    End With
    FillListViews
    MsgBox "New record has been successfully saved.", vbInformation
    Resetfields
End Sub

Sub Resetfields()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Check1.Value = Unchecked
    lvButtons_H2.Enabled = True
   cmdSave.Enabled = False
End Sub

Sub FillListViews()
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblAccount", CN, adOpenStatic, adLockPessimistic
With rsData
  Do While Not .EOF
                Dim xy As ListItem
                Set xy = ListView1.ListItems.Add(, , rsData!UserCode, 1, 1)
                    xy.SubItems(1) = rsData!UserNAME
                    xy.SubItems(2) = rsData!CompleteName
                    xy.SubItems(3) = rsData!Type
                    xy.SubItems(4) = rsData!PK
    .MoveNext
Loop
End With
End Sub

Private Sub Form_Load()
FillListViews
cmdSave.Enabled = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
End Sub
 
Sub GenerateNumber()
     Dim PK As Long
     PK = getIndex("tblAccount")
     Me.Text4.Text = GenerateID(PK, "ACCNT", "0000")
End Sub
 


Private Sub Form_Unload(Cancel As Integer)
Set rsData = Nothing
End Sub

Private Sub lvButtons_H1_Click()
        If ListView1.ListItems.Count < 1 Then
                MsgBox "No Records", vbExclamation
             
              Else
             If ListView1.SelectedItem.ListSubItems(4) = CurrUser.UserPK Then
                  MsgBox "" & ListView1.SelectedItem.ListSubItems(2).Text & " is currently using this account.Can't proceed to the operation!", vbExclamation
                Else
             
                Dim wheng As VbMsgBoxResult
                wheng = MsgBox("You are about to delete (1) record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
                If wheng = vbYes Then
                CN.Execute ("Delete from  tblAccount where UserCode= '" & ListView1.SelectedItem.Text & "'")
                Call Form_Load
                MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
            End If
              End If
              End If
              
End Sub

Private Sub lvButtons_H2_Click()
cmdSave.Enabled = True
lvButtons_H2.Enabled = False
GenerateNumber
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
End Sub
 
Private Sub Text3_LostFocus()
 Text3.Text = StrConv(Text3, vbProperCase)
End Sub

 
