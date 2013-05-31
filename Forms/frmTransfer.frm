VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmTransfer 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer Items"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "frmTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCategory 
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtSupplier 
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   1320
      TabIndex        =   13
      Tag             =   "Quantity"
      Text            =   "0"
      Top             =   2160
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   1680
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   240
      ScaleHeight     =   45
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   2760
      Width           =   5775
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
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   1320
      Width           =   4650
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   1890
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
      ScaleWidth      =   410
      TabIndex        =   2
      Top             =   0
      Width           =   6150
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmTransfer.frx":57E2
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
         TabIndex        =   4
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
         TabIndex        =   3
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
      TabIndex        =   1
      Top             =   742
      Width           =   9375
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   330
      Left            =   3840
      TabIndex        =   9
      Top             =   3000
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
      Left            =   5160
      TabIndex        =   10
      Top             =   3000
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
      Caption         =   "Total Qty"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse / Store"
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
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1680
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
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Tag             =   "Description"
      Top             =   1320
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
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditFlag         As Boolean
Public PK               As String
Public rsData           As ADODB.Recordset
Dim adQty, adqtyRemain  As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    If Combo1.Text = "" Then
    MsgBox "Please enter Warehouse/Store", vbExclamation
    Combo1.SetFocus
    HighL Combo1.Text
    Exit Sub
    End If


Dim getQty As Integer
Dim getWar As String

getQty = Me.Text3.Text

If Check1.Value = 1 Then

sql = "INSERT INTO tblStockTransfer(ProductCode,Description,CategoryName,SupplierName,Qty,SellingPrice,Warehouse,DateTransferred,TransferredBy) " & _
      " VALUES ('" & Me.Text1.Text & "','" & Me.Text2.Text & "','" & Me.txtCategory.Text & "','" & Me.txtSupplier.Text & "','" & Me.Text3.Text & "','" & Me.txtPrice.Text & "','" & Me.Combo1.Text & "','" & Format(Now, "YYYY-MM-DD") & "','Renje')   "
sql1 = "Delete from tblProduct where ProductCode = '" & Me.Text1.Text & "'"
Set RS = CN.Execute(sql1)

Else
getWar = "Transferred at " & Me.Combo1.Text
sql = "UPDATE tblProduct Set QtyRemain = " & getQty & ",Qty = " & getQty & ", Notes = '" & getWar & "',Createdby ='Renje',DateCreated = '" & Format(Now, "YYYY-MM-DD") & "' where ProductCode = '" & Me.Text1.Text & "'"
EditFlag = True
End If

Set rsData = CN.Execute(sql)


''With frmMonitoring.rsData
''adQty = .Fields("Qty")
''adqtyRemain = .Fields("QtyRemain")
         
''        .Fields("Qty") = .Fields("Qty") + Val(Text4.Text)
''        .Fields("QtyRemain") = .Fields("QtyRemain") - Val(Text4.Text)
''        .Fields("Notes") = "Transferred as of " & Format(Now, "YYYY-MM-DD")
''
''
''    .Update
''End With
''
    If EditFlag = True Then
        Call FindRec(frmMonitoring.rsData, "ProductCode", True, Text1.Text, 0)
        MsgBox "Update in stock record has been successfull.", vbInformation
        frmMonitoring.LoadEntries
        Unload Me
    Else
        MsgBox "Stock record has been successfully transferred.", vbInformation
        frmMonitoring.LoadEntries
        Unload Me
    
    End If
End Sub

 
 
Private Sub Form_Unload(Cancel As Integer)
If EditFlag = True Then frmMonitoring.LoadEntries
Set rsData = Nothing
End Sub

Private Sub Text4_GotFocus()
HighL Text4.Text
End Sub

Private Sub Text4_LostFocus()
HighL Text4.Text
End Sub

