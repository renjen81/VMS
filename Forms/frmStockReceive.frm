VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmStockReceive 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Entry"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "frmStockReceive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5760
      TabIndex        =   16
      Tag             =   "Total Qty"
      Text            =   "0"
      Top             =   1305
      Width           =   1815
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
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Tag             =   "Price"
      Text            =   "0.00"
      Top             =   960
      Width           =   1815
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
      TabIndex        =   12
      Tag             =   "Supplier"
      Top             =   2160
      Width           =   3330
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   7455
      TabIndex        =   9
      Top             =   2700
      Width           =   7455
   End
   Begin VB.ComboBox Text1 
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
      TabIndex        =   5
      Tag             =   "Product Code"
      Top             =   960
      Width           =   3330
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "Description"
      Top             =   1320
      Width           =   3330
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
      ScaleWidth      =   512
      TabIndex        =   1
      Top             =   0
      Width           =   7680
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmStockReceive.frx":57E2
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
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
         TabIndex        =   3
         Top             =   180
         Width           =   1110
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
         TabIndex        =   2
         Top             =   510
         Width           =   3900
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   742
      Width           =   9375
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   330
      Left            =   5400
      TabIndex        =   10
      Top             =   2880
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
      Left            =   6720
      TabIndex        =   11
      Top             =   2880
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
   Begin VB.Label Label5 
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
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   4800
      TabIndex        =   15
      Top             =   1305
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price(Each)"
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
      Left            =   4800
      TabIndex        =   14
      Top             =   960
      Width           =   855
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
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Tag             =   "Description"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      TabIndex        =   7
      Top             =   960
      Width           =   1095
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "frmStockReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditFlag         As Boolean
Public PK               As String
Public rsData           As ADODB.Recordset



Sub dDataCombo()
openRec rsData, "tblPRoduct"
While rsData.EOF <> True
            Text1.AddItem rsData![Description]
         rsData.MoveNext
Wend

openRec rsData, "tblSupplier"
While rsData.EOF <> True
            Text3.AddItem rsData![SupplierName]
         rsData.MoveNext
Wend
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If is_empty(Text1, True) = True Then Exit Sub
If is_empty(Text2, True) = True Then Exit Sub
If is_empty(Text3, True) = True Then Exit Sub
If is_empty(Text4, True) = True Then Exit Sub
If is_empty(Text5, True) = True Then Exit Sub

If toNumber(Text5.Text) = 0 Then
        MsgBox "Please enter quantity to receive.", vbExclamation
        Text5.SetFocus
        HighL Text5
        Exit Sub
End If
 
With frmStockRecieves.rsData
    If EditFlag = True Then .AddNew
        .Fields("ProductCode") = Text2.Text
        .Fields("Description") = Text1.Text
        .Fields("SupplierName") = Text3.Text
        .Fields("SellingPrice") = Text4.Text
        .Fields("Qty") = Text5.Text
        .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
        .Fields("CreatedBy") = CurrUser.UserNAME
    .Update
End With
 
    If EditFlag = True Then
        Call FindRec(frmStockRecieves.rsData, "ProductCode", True, Text2.Text, 0)
        MsgBox "New record has been successfully added.", vbInformation
        
        Call ProductQtyAdd
        
        Dim REPLY As Integer
        REPLY = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
        If REPLY = vbYes Then
            Resetfields
        Else
            frmStockRecieves.LoadEntries
            Unload Me
        End If
    Else
        Call FindRec(frmStockRecieves.rsData, "ProductCode", True, Text2.Text, 0)
        MsgBox "Changes in record has been successfully saved.", vbInformation
        frmStockRecieves.LoadEntries
        Unload Me
    End If
End Sub

Sub Resetfields()
clearText Me
Text1.SetFocus
frmStockRecieves.LoadEntries
End Sub
 
 
Private Sub Form_Load()
If EditFlag = False Then
        Caption = "Edit Existing"
        dDataCombo
Else
         clearText Me
         dDataCombo
End If
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
If EditFlag = True Then frmStockRecieves.LoadEntries
Set rsData = Nothing
End Sub
 

Private Sub Text1_Click()
On Error Resume Next
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblProduct where Description  = '" & Text1.Text & "'", CN, adOpenStatic, adLockPessimistic
Text2.Text = rsData!ProductCode
Text4.Text = rsData!SellingPrice

End Sub

Private Sub Text4_GotFocus()
Text4.Text = Format(Text4, "###,#0.00")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub
Private Sub Text4_LostFocus()
Text4.Text = Format(Text4, "###,#0.00")
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub


Public Sub ProductQtyAdd()
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblProduct where ProductCode = '" & Text2.Text & "'", CN, adOpenStatic, adLockPessimistic
With rsData
         EditFlag = True
        .Fields("ProductCode") = Text2.Text
        .Fields("Description") = Text1.Text
        .Fields("SellingPrice") = Text4.Text
        .Fields("Qty") = .Fields("Qty") + Val(Text5.Text)
        .Fields("QtyRemain") = .Fields("QtyRemain") + Val(Text5.Text)
    .Update
End With
End Sub
