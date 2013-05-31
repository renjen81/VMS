VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSalesDaily 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Daily Sales"
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   10455
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   17775
      Begin VB.Label Label19 
         Caption         =   "Under Construction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4680
         TabIndex        =   13
         Top             =   3960
         Width           =   13815
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   114819073
      CurrentDate     =   41371
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9765
      TabIndex        =   8
      Top             =   7440
      Width           =   9765
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No Record"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   255
         Left            =   77
         TabIndex        =   9
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9765
      TabIndex        =   0
      Top             =   0
      Width           =   9765
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   2295
      End
      Begin VB.Image cmdExit 
         Height          =   360
         Left            =   9320
         MouseIcon       =   "frmSalesDaily.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmSalesDaily.frx":0152
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   -27
         Width           =   360
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   3975
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   35
      TabIndex        =   3
      Top             =   1020
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilList"
      SmallIcons      =   "ilList"
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmSalesDaily.frx":083C
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice NO"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Code"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   4128
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total Amount"
         Object.Width           =   4128
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Remarks"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "DateCreated"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CreatedBy"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   0
      Left            =   45
      TabIndex        =   4
      Top             =   360
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1085
      Caption         =   "New"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSalesDaily.frx":099E
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmSalesDaily.frx":3730
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   1
      Left            =   1035
      TabIndex        =   5
      Top             =   360
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1085
      Caption         =   "Print"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSalesDaily.frx":3892
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmSalesDaily.frx":45EC
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   4
      Left            =   3285
      TabIndex        =   6
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1085
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSalesDaily.frx":474E
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmSalesDaily.frx":54A8
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
      Caption         =   "Refresh"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483642
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   16119285
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSalesDaily.frx":560A
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmSalesDaily.frx":73A4
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   8640
      Top             =   360
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
            Picture         =   "frmSalesDaily.frx":7506
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Today:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmSalesDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcRecord               As String
Dim srcRecCD                As Variant
Public rsData               As ADODB.Recordset
Public Total                As Currency
Public gtotal               As Currency
Public xQty                 As Single
Public tempx                As Integer
Dim sql                     As String
Public Sub LoadEntries()
tempx = 0
xQty = 0
gtotal = 0
Total = 0
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblSalesInvoice", CN, adOpenStatic, adLockPessimistic
With rsData
    While Not .EOF
                Dim LS As ListItem
                Set LS = ListView1.ListItems.Add(, , rsData!InvoiceNO, 1, 1)
                    LS.SubItems(1) = rsData!ProductCode
                    LS.SubItems(2) = rsData!Description
                    LS.SubItems(3) = rsData!Qty
                    LS.SubItems(4) = Format(rsData!Price, "###,#0.00")
                    LS.SubItems(5) = Format(rsData!TotalAmount, "###,#0.00")
                    LS.SubItems(6) = rsData!Remarks
                    LS.SubItems(7) = rsData!DateCreated
                    LS.SubItems(8) = rsData!CreatedBy
                    
                    xQty = xQty + rsData!Qty
                    
                    gtotal = gtotal + Format(rsData!Price, "###,#0.00")
                    Total = Total + Format(rsData!TotalAmount, "###,#0.00")
                    
                    LS.ListSubItems(1).Bold = True
                    LS.ListSubItems(2).Bold = True
            .MoveNext
       Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    Wend
    End With
 tempx = tempx + 1
 Set LS = ListView1.ListItems.Add(, , "")
  LS.SubItems(3) = "Total : " & xQty
  
  LS.SubItems(4) = "Total : " & Format(gtotal, "###,#0.00")
  LS.SubItems(5) = "Total : " & Format(Total, "###,#0.00")
  
  ListView1.ListItems(tempx).ListSubItems(3).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(4).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(5).ForeColor = vbBlue
  
  LS.ListSubItems(3).Bold = True
  LS.ListSubItems(4).Bold = True
  LS.ListSubItems(5).Bold = True
End Sub

Private Sub cmdButtons_Click(Index As Integer)
Select Case Index

Case 0: Command "New"

Case 1: Command "Print"

Case 3: Command "Refresh"

Case 4: Command "Close"

End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub DTPicker1_Change()
LoadEntriesFill
End Sub

Private Sub DTPicker1_Click()
LoadEntriesFill
End Sub

Private Sub Form_Load()
Call LoadEntries
End Sub

 
Private Sub Form_Unload(Cancel As Integer)
     MDIForm1.mDelete.Enabled = True
     MDIForm1.mEdit.Enabled = True
     MDIForm1.mNew.Enabled = True
     MDIForm1.mPrint.Enabled = False
     loadForm frmWelcome
End Sub

Private Sub ListView1_DblClick()
Command "Edit"
End Sub

 
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
               On Error Resume Next
               Command "New"
               
     Case vbKeyDelete
               On Error Resume Next
               Command "Delete"
               
    Case vbKeyP Or (KeyCode = 109 And Shift = 2)
               On Error Resume Next
               Command "Print"
               
    Case vbKeyF5
               On Error Resume Next
               Command "Refresh"
               
    Case 67 And Shift = 2
                On Error Resume Next
                Command "Close"
End Select
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then ListView1_Click
End Sub

Private Sub Picture1_Resize()
  cmdExit.Left = Picture1.Width - cmdExit.Width ' - -23
End Sub


Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.ScaleWidth
ListView1.Height = (Me.ScaleHeight - Picture2.Height) - ListView1.Top
 End Sub

Sub RefreshRecords()
Form_Load
End Sub

 
Private Sub ListView1_Click()
    If Trim(srcRecord) = vbNullString Then
         Label1.Caption = "No Record"
    Else
       Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
      MDIForm1.mEdit.Enabled = False
     PopupMenu MDIForm1.mAction
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortLV ListView1
End Sub

 Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcRecCD = ListView1.SelectedItem.Index
srcRecord = ListView1.ListItems.Item(srcRecCD).Text
End Sub

 
Public Sub Command(cmd As String)
Select Case cmd

Case "New"
                With frmSalesInvoice
                    .Show vbModal
                End With
Case "Refresh"
  Call LoadEntries
  
Case "Print"
    Set rsData = New ADODB.Recordset
    rsData.Open "tblSalesInvoice", CN, adOpenStatic, adLockPessimistic
    Set rptDailySales.DataSource = rsData
    rptDailySales.Show
    
Case "Close"
      Unload Me


End Select
End Sub

Public Sub LoadEntriesFill()
tempx = 0
xQty = 0
gtotal = 0
Total = 0
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
sql = "Select * from tblSalesInvoice where month(DateCreated) = '" & Month(Me.DTPicker1.Value) & "' and day(DateCreated) = '" & Day(Me.DTPicker1.Value) & "' and year(DateCreated) = '" & Year(Me.DTPicker1.Value) & "'  "
rsData.Open sql, CN, adOpenStatic, adLockPessimistic
With rsData
    While Not .EOF
                Dim LS As ListItem
                Set LS = ListView1.ListItems.Add(, , rsData!InvoiceNO, 1, 1)
                    LS.SubItems(1) = rsData!ProductCode
                    LS.SubItems(2) = rsData!Description
                    LS.SubItems(3) = rsData!Qty
                    LS.SubItems(4) = Format(rsData!Price, "###,#0.00")
                    LS.SubItems(5) = Format(rsData!TotalAmount, "###,#0.00")
                    LS.SubItems(6) = rsData!Remarks
                    LS.SubItems(7) = rsData!DateCreated
                    LS.SubItems(8) = rsData!CreatedBy
                    
                    xQty = xQty + rsData!Qty
                    
                    gtotal = gtotal + Format(rsData!Price, "###,#0.00")
                    Total = Total + Format(rsData!TotalAmount, "###,#0.00")
                    
                    LS.ListSubItems(1).Bold = True
                    LS.ListSubItems(2).Bold = True
            .MoveNext
       Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    Wend
    End With
 tempx = tempx + 1
 Set LS = ListView1.ListItems.Add(, , "")
  LS.SubItems(3) = "Total : " & xQty
  
  LS.SubItems(4) = "Total : " & Format(gtotal, "###,#0.00")
  LS.SubItems(5) = "Total : " & Format(Total, "###,#0.00")
  
  ListView1.ListItems(tempx).ListSubItems(3).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(4).ForeColor = vbBlue
  ListView1.ListItems(tempx).ListSubItems(5).ForeColor = vbBlue
  
  LS.ListSubItems(3).Bold = True
  LS.ListSubItems(4).Bold = True
  LS.ListSubItems(5).Bold = True
End Sub



