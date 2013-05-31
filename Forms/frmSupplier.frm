VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSupplier 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Entry"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboBrgy 
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
      Left            =   1380
      TabIndex        =   23
      Tag             =   "Category"
      Top             =   2400
      Width           =   3690
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   9135
      TabIndex        =   22
      Top             =   4200
      Width           =   9135
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   21
      Top             =   742
      Width           =   9375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      TabIndex        =   9
      Tag             =   "Address"
      Top             =   2040
      Width           =   3720
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      TabIndex        =   8
      Tag             =   "Supplier Name"
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      TabIndex        =   7
      Tag             =   "Contact No"
      Top             =   1680
      Width           =   3735
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
      ScaleWidth      =   624
      TabIndex        =   4
      Top             =   0
      Width           =   9360
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmSupplier.frx":57E2
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leader"
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
         TabIndex        =   6
         Top             =   180
         Width           =   960
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
         TabIndex        =   5
         Top             =   510
         Width           =   3900
      End
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2115
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      TabIndex        =   2
      Tag             =   "Position"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      TabIndex        =   1
      Tag             =   "Contact Person"
      Top             =   2760
      Width           =   3735
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
      TabIndex        =   0
      Top             =   3720
      Value           =   1  'Checked
      Width           =   795
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   330
      Left            =   7080
      TabIndex        =   11
      Top             =   4440
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
      Left            =   8400
      TabIndex        =   12
      Top             =   4440
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
      TabIndex        =   24
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5400
      X2              =   5400
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   1
      Left            =   150
      TabIndex        =   20
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leader Name"
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
      Index           =   0
      Left            =   150
      TabIndex        =   19
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leader Code"
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
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
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
      TabIndex        =   17
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   5760
      TabIndex        =   16
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Left            =   150
      TabIndex        =   15
      Top             =   2760
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Index           =   3
      Left            =   150
      TabIndex        =   14
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label lblRC 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   960
      TabIndex        =   13
      Top             =   3720
      Width           =   1485
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditFlag         As Boolean
Public PK               As String
Public rsData           As ADODB.Recordset
 

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo error:

If is_empty(Text2, True) = True Then Exit Sub
If is_empty(Text3, True) = True Then Exit Sub
If is_empty(Text4, True) = True Then Exit Sub
If is_empty(Text5, True) = True Then Exit Sub
If is_empty(Text6, True) = True Then Exit Sub
If is_empty(cboBrgy, True) = True Then Exit Sub
  
  
With frmSuppliers.rsData
    If EditFlag = True Then .AddNew
        .Fields("LeaderCode") = Text1.Text
        .Fields("LeaderName") = Text2.Text
        .Fields("ContactNos") = Text3.Text
        .Fields("Address") = Text4.Text
        .Fields("Brgy") = cboBrgy.Text
        .Fields("ContactPerson") = Text5.Text
        .Fields("Positions") = Text6.Text
        .Fields("Active") = chkActive.Value
        .Fields("Remarks") = Text7.Text
        .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
        .Fields("CreatedBy") = CurrUser.UserNAME
    .Update
End With

    If EditFlag = True Then
        Call FindRec(frmSuppliers.rsData, "LeaderCode", True, Text1.Text, 0)
        MsgBox "New record has been successfully added.", vbInformation
        Dim REPLY As Integer
        REPLY = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
        If REPLY = vbYes Then
            Resetfields
        Else
            frmSuppliers.LoadEntries
            Unload Me
        End If
    Else
        Call FindRec(frmSuppliers.rsData, "LeaderCode", True, Text1.Text, 0)
        MsgBox "Changes in record has been successfully saved.", vbInformation
        frmSuppliers.LoadEntries
        Unload Me
    End If
    
Exit Sub
error:
    MsgBox err.Description, vbExclamation
    Unload Me
End Sub

Sub Resetfields()
clearText Me
GeneratePK
Text2.SetFocus
frmSuppliers.LoadEntries
End Sub
Sub GeneratePK()
Dim iCode  As Long
    iCode = getIndex("tblLeader")
    Text1.Text = GenerateID(iCode, "LEAD-", "00000")
End Sub
 
Private Sub Form_Load()

openRec rsData, "tblBarangay"
While rsData.EOF <> True
            cboBrgy.AddItem rsData![BrgyName]
         rsData.MoveNext
Wend

If EditFlag = False Then
        
        Caption = "Edit Existing"
Else
        clearText Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If EditFlag = True Then frmSuppliers.LoadEntries
frmWelcome.LOAD_MY_URL
Set rsData = Nothing
End Sub
 
Private Sub Text2_LostFocus()
ToUpper Text2
End Sub
 
Private Sub Text5_GotFocus()
 Text5.Text = StrConv(Text5, vbProperCase)
End Sub

Private Sub Text5_LostFocus()
 Text5.Text = StrConv(Text5, vbProperCase)
End Sub

Private Sub Text6_GotFocus()
 Text6.Text = StrConv(Text6, vbProperCase)
End Sub

 Private Sub Text6_LostFocus()
 Text6.Text = StrConv(Text6, vbProperCase)
End Sub

Private Sub Text4_GotFocus()
 Text4.Text = StrConv(Text4, vbProperCase)
End Sub

Private Sub Text4_LostFocus()
 Text4.Text = StrConv(Text4, vbProperCase)
End Sub
Private Sub cboBrgy_LostFocus()
 cboBrgy.Text = StrConv(cboBrgy, vbProperCase)
End Sub

