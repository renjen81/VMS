VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmProduct 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "frmProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cboBarangay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      TabIndex        =   30
      Tag             =   "Unit Price"
      Top             =   4080
      Width           =   3210
   End
   Begin VB.TextBox txtMiddleName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Tag             =   "Unit Price"
      Top             =   2760
      Width           =   3210
   End
   Begin VB.TextBox txtFirstName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Tag             =   "Unit Price"
      Top             =   2400
      Width           =   3210
   End
   Begin VB.TextBox txtLastName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      TabIndex        =   24
      Tag             =   "Unit Price"
      Top             =   2040
      Width           =   3210
   End
   Begin VB.ComboBox cboLeader 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   1080
      TabIndex        =   22
      Tag             =   "Supplier"
      Top             =   4440
      Width           =   3210
   End
   Begin VB.ComboBox cboPrecint 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   1080
      TabIndex        =   20
      Tag             =   "Supplier"
      Top             =   3720
      Width           =   3210
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   13
      Top             =   742
      Width           =   9375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   8655
      TabIndex        =   12
      Top             =   5490
      Width           =   8655
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
      ScaleWidth      =   591
      TabIndex        =   4
      Top             =   0
      Width           =   8865
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6720
         TabIndex        =   19
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   7200
         TabIndex        =   18
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
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
         Left            =   720
         TabIndex        =   6
         Top             =   510
         Width           =   3900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voter"
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
         Left            =   720
         TabIndex        =   5
         Top             =   180
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmProduct.frx":57E2
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   1890
   End
   Begin VB.TextBox txtCompleteName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   525
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   3120
      Width           =   3210
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H80000014&
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   3045
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   3810
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   330
      Left            =   6600
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "&Save"
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
      Left            =   7920
      TabIndex        =   15
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "&Close"
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
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
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
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Leader"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precint No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Caption         =   "Other Informations "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000010&
      Caption         =   "Voter Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voters Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Tag             =   "Description"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   4560
      X2              =   4560
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Label lblRC 
      AutoSize        =   -1  'True
      BackColor       =   &H00F0F0F0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   360
      TabIndex        =   7
      Top             =   5160
      Width           =   2805
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public State            As FormState
Public EditFlag         As Boolean
Public PK               As String
Public rsData           As ADODB.Recordset

 

Sub dDataCombo()


openRec rsData, "tblLeader"
While rsData.EOF <> True
            cboLeader.AddItem rsData![LeaderName]
         rsData.MoveNext
Wend

openRec rsData, "tblPrecint"
While rsData.EOF <> True
            cboPrecint.AddItem rsData![PrecintName]
         rsData.MoveNext
Wend


End Sub



Private Sub cboPrecint_Click()
Set rsData = New ADODB.Recordset

sql = "Select * from tblPrecint where PrecintName = '" & Me.cboPrecint.Text & "'"
rsData.Open sql, CN, adOpenStatic, adLockOptimistic
With rsData
    Me.cboBarangay.Text = .Fields("Brgy")

End With

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If is_empty(txtFirstName, True) = True Then Exit Sub
If is_empty(txtLastName, True) = True Then Exit Sub
If is_empty(txtMiddleName, True) = True Then Exit Sub
If is_empty(txtCompleteName, True) = True Then Exit Sub
If is_empty(cboBarangay, True) = True Then Exit Sub
If is_empty(cboPrecint, True) = True Then Exit Sub
If is_empty(cboLeader, True) = True Then Exit Sub

If EditFlag = True Then
    Set rsData = New ADODB.Recordset
    rsData.Open "tblVoter", CN, adOpenKeyset, adLockOptimistic
    With rsData
            .AddNew
            .Fields("VoterCode") = txtCode.Text
            .Fields("FullName") = txtCompleteName.Text
            .Fields("LastName") = txtLastName.Text
            .Fields("MiddleName") = txtMiddleName.Text
            .Fields("FirstName") = txtFirstName.Text
            .Fields("Brgy") = cboBarangay.Text
            .Fields("PrecintNo") = cboPrecint.Text
            .Fields("Leader") = cboLeader.Text
            .Fields("Remark") = txtRemarks.Text
            .Fields("Active") = chkActive.Value
            .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
            .Fields("CreatedBy") = CurrUser.UserNAME
        
        
        .Update
    End With
    MsgBox "New record has been successfully added.", vbInformation
    Dim REPLY As Integer
    REPLY = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
    If REPLY = vbYes Then
            Resetfields
    
    Else
    
    frmProducts.LoadEntries
            Unload Me
    End If
Else

    Set rsData = New ADODB.Recordset
    
    rsData.Open "Select * from tblVoter where VoterCode = '" & Me.txtCode.Text & "' ", CN, adOpenKeyset, adLockOptimistic
    With rsData
            .Fields("VoterCode") = txtCode.Text
            .Fields("FullName") = txtCompleteName.Text
            .Fields("LastName") = txtLastName.Text
            .Fields("MiddleName") = txtMiddleName.Text
            .Fields("FirstName") = txtFirstName.Text
            .Fields("Brgy") = cboBarangay.Text
            .Fields("PrecintNo") = cboPrecint.Text
            .Fields("Leader") = cboLeader.Text
            .Fields("Remark") = txtRemarks.Text
            .Fields("Active") = chkActive.Value
            .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
            .Fields("CreatedBy") = CurrUser.UserNAME
        
        .Update
    End With


    MsgBox "Changes in record has been successfully saved.", vbInformation
    Unload Me
End If
End Sub

Sub Resetfields()
clearText Me
GeneratePK

txtLastName.SetFocus
frmProducts.LoadEntries
End Sub
Sub GeneratePK()
Dim iCode  As Long
    iCode = getIndex("tblVoter")
    txtCode.Text = GenerateID(iCode, "PCODE-", "00000")
End Sub
 
Private Sub Form_Load()
If EditFlag = True Then
         dDataCombo
         GeneratePK
Else
    Caption = "Edit Existing"
    dDataCombo
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

frmProducts.LoadEntries
frmWelcome.LOAD_MY_URL

Set rsData = Nothing
Exit Sub
err:
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub
 
Private Sub txtMiddleName_Change()
Me.txtCompleteName = Me.txtLastName.Text & ", " & Me.txtFirstName.Text & " " & Left(Me.txtMiddleName.Text, 1) & "."
End Sub
