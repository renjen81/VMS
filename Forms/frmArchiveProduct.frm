VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmProductArchive 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "frmArchiveProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSizes 
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
      Left            =   1200
      TabIndex        =   36
      Tag             =   "Supplier"
      Top             =   3600
      Width           =   3090
   End
   Begin VB.TextBox txtReorder 
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
      Left            =   1200
      TabIndex        =   34
      Tag             =   "Quantity"
      Top             =   5760
      Width           =   1290
   End
   Begin VB.TextBox txtCritical 
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
      Left            =   1200
      TabIndex        =   32
      Tag             =   "Quantity"
      Top             =   5400
      Width           =   1290
   End
   Begin VB.ComboBox txtUnit 
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
      Left            =   1200
      TabIndex        =   30
      Tag             =   "Supplier"
      Top             =   3240
      Width           =   3090
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   6570
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
      TabIndex        =   9
      Top             =   0
      Width           =   8865
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6720
         TabIndex        =   29
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   7200
         TabIndex        =   28
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
         TabIndex        =   11
         Top             =   510
         Width           =   3900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Archive"
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
         TabIndex        =   10
         Top             =   180
         Width           =   2265
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmArchiveProduct.frx":57E2
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1890
   End
   Begin VB.TextBox Text2 
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
      Height          =   765
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   1920
      Width           =   3090
   End
   Begin VB.TextBox Text5 
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
      Left            =   1200
      TabIndex        =   3
      Tag             =   "Quantity"
      Top             =   5040
      Width           =   1290
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
      TabIndex        =   7
      Top             =   6240
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.TextBox Text6 
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
      Left            =   5880
      TabIndex        =   4
      Tag             =   "Unit Price"
      Top             =   1560
      Width           =   2850
   End
   Begin VB.TextBox Text7 
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
      Left            =   5880
      TabIndex        =   5
      Tag             =   "Selling Price"
      Top             =   1920
      Width           =   2850
   End
   Begin VB.TextBox Text8 
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
      Height          =   2205
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2520
      Width           =   3810
   End
   Begin VB.ComboBox Text3 
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
      Left            =   1200
      TabIndex        =   1
      Tag             =   "Category"
      Top             =   2880
      Width           =   3090
   End
   Begin VB.ComboBox Text4 
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
      Left            =   1200
      TabIndex        =   2
      Tag             =   "Supplier"
      Top             =   3960
      Width           =   3090
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   330
      Left            =   6600
      TabIndex        =   23
      Top             =   6840
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
      TabIndex        =   24
      Top             =   6840
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
      Caption         =   "Sizes/Term"
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
      TabIndex        =   37
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reorder Level"
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
      Index           =   8
      Left            =   120
      TabIndex        =   35
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Critical Level"
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
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Type"
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
      TabIndex        =   31
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Caption         =   " Pricing Informations "
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
      TabIndex        =   27
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000010&
      Caption         =   " Product Information"
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
      TabIndex        =   26
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000010&
      Caption         =   "Qty Informations "
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
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "General Name"
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
      TabIndex        =   19
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
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
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Specific Name"
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
      TabIndex        =   17
      Tag             =   "Description"
      Top             =   1920
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
      TabIndex        =   16
      Top             =   2280
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
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price"
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
      Index           =   5
      Left            =   4920
      TabIndex        =   15
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      Index           =   4
      Left            =   4920
      TabIndex        =   14
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
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
      Left            =   1080
      TabIndex        =   12
      Top             =   6240
      Width           =   2805
   End
End
Attribute VB_Name = "frmProductArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public State            As FormState
Public EditFlag         As Boolean
Public PK               As String
Public rsData           As ADODB.Recordset

 

Sub dDataCombo()
openRec rsData, "tblCategory"
While rsData.EOF <> True
            Text3.AddItem rsData![CategoryName]
         rsData.MoveNext
Wend

openRec rsData, "tblSupplier"
While rsData.EOF <> True
            Text4.AddItem rsData![SupplierName]
         rsData.MoveNext
Wend

openRec rsData, "tblUnit"
While rsData.EOF <> True
            txtUnit.AddItem rsData![UnitName]
         rsData.MoveNext
Wend


End Sub

Private Sub cboSizes_Change()

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If is_empty(Text2, True) = True Then Exit Sub
If is_empty(Text3, True) = True Then Exit Sub
If is_empty(Text4, True) = True Then Exit Sub
If is_empty(Text5, True) = True Then Exit Sub
If is_empty(Text6, True) = True Then Exit Sub
If is_empty(Text7, True) = True Then Exit Sub
If is_empty(txtCritical, True) = True Then Exit Sub
If is_empty(txtReorder, True) = True Then Exit Sub
If is_empty(cboSizes, True) = True Then Exit Sub

If is_empty(txtUnit, True) = True Then Exit Sub
If EditFlag = True Then
    Set rsData = New ADODB.Recordset
    rsData.Open "tblProduct", CN, adOpenKeyset, adLockOptimistic
    With rsData
            .AddNew
            .Fields("ProductCode") = Text1.Text
            .Fields("Description") = Text2.Text
            .Fields("CategoryName") = Text3.Text
            .Fields("SupplierName") = Text4.Text
            .Fields("Qty") = Text5.Text
            .Fields("QtyRemain") = Text9.Text
            .Fields("QtySold") = Text10.Text
            .Fields("UnitPrice") = Text6.Text
            .Fields("SellingPrice") = Text7.Text
            .Fields("Remarks") = Text8.Text
            .Fields("Active") = chkActive.Value
            .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
            .Fields("UnitName") = txtUnit.Text
            .Fields("CreatedBy") = CurrUser.UserNAME
            .Fields("CriticalLevel") = Me.txtCritical.Text
            .Fields("ReorderLevel") = Me.txtReorder.Text
            .Fields("Sizes") = Me.cboSizes.Text
        
        
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
    
    rsData.Open "Select * from tblProduct where ProductCode = '" & Me.Text1.Text & "' ", CN, adOpenKeyset, adLockOptimistic
    With rsData
            .Fields("ProductCode") = Text1.Text
            .Fields("Description") = Text2.Text
            .Fields("CategoryName") = Text3.Text
            .Fields("SupplierName") = Text4.Text
            .Fields("Qty") = Text5.Text
            .Fields("QtyRemain") = Text9.Text
            .Fields("QtySold") = Text10.Text
            .Fields("UnitPrice") = Text6.Text
            .Fields("SellingPrice") = Text7.Text
            .Fields("Remarks") = Text8.Text
            .Fields("Active") = chkActive.Value
            .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
            .Fields("UnitName") = txtUnit.Text
            .Fields("CreatedBy") = CurrUser.UserNAME
            .Fields("CriticalLevel") = Me.txtCritical.Text
            .Fields("ReorderLevel") = Me.txtReorder.Text
            .Fields("Sizes") = Me.cboSizes.Text
        
        .Update
    End With


    MsgBox "Changes in record has been successfully saved.", vbInformation
    Unload Me
End If
End Sub

Sub Resetfields()
clearText Me
GeneratePK
Text3.ListIndex = -1
Text4.ListIndex = -1
Text2.SetFocus
frmProducts.LoadEntries
End Sub
Sub GeneratePK()
Dim iCode  As Long
    iCode = getIndex("tblProduct")
    Text1.Text = GenerateID(iCode, "PCODE-", "00000")
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
 
Private Sub Text2_LostFocus()
ToUpper Text2
End Sub

 
Private Sub Text3_Change()

End Sub

Private Sub Text4_Change()

End Sub

Private Sub Text5_Change()
    If EditFlag = True Then
        Text9.Text = Text5.Text
    ElseIf EditFlag = False Then
        Text9.Text = (Val(Text5.Text) - (Val(Text10.Text)))
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
Text6.Text = Format(Text6, "###,#0.00")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
Text6.Text = Format(Text6, "###,#0.00")
End Sub

Private Sub Text7_GotFocus()
Text7.Text = Format(Text7, "###,#0.00")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub Text7_LostFocus()
Text7.Text = Format(Text7, "###,#0.00")
End Sub

Private Sub txtUnit_Change()

End Sub
