VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmStockRecieves 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9855
      TabIndex        =   9
      Top             =   7770
      Width           =   9855
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
         TabIndex        =   10
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
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Image cmdExit 
         Height          =   360
         Left            =   9320
         MouseIcon       =   "frmStockRecieves.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmStockRecieves.frx":0152
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   -27
         Width           =   360
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Receive"
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
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Receive"
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
      Left            =   45
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
      MouseIcon       =   "frmStockRecieves.frx":083C
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Code"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5715
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Supplier"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Selling Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Created"
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
      Image           =   "frmStockRecieves.frx":099E
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmStockRecieves.frx":3730
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   1
      Left            =   1065
      TabIndex        =   5
      Top             =   360
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1085
      Caption         =   "Edit"
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
      Image           =   "frmStockRecieves.frx":3892
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmStockRecieves.frx":562C
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   2
      Left            =   2085
      TabIndex        =   6
      Top             =   360
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1085
      Caption         =   "Del"
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
      Image           =   "frmStockRecieves.frx":578E
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmStockRecieves.frx":7528
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   3
      Left            =   3120
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
      Image           =   "frmStockRecieves.frx":768A
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmStockRecieves.frx":9424
   End
   Begin lvButton.lvButtons_H cmdButtons 
      Height          =   615
      Index           =   4
      Left            =   4380
      TabIndex        =   8
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
      Image           =   "frmStockRecieves.frx":9586
      ImgSize         =   32
      cBack           =   16119285
      mPointer        =   99
      mIcon           =   "frmStockRecieves.frx":A2E0
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   6360
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
            Picture         =   "frmStockRecieves.frx":A442
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStockRecieves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcRecord               As String
Dim srcRecCD                As Variant
Public rsData               As ADODB.Recordset

Public Sub LoadEntries()
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblStockReceive", CN, adOpenStatic, adLockPessimistic
With rsData
    While Not .EOF
                Dim LS As ListItem
                Set LS = ListView1.ListItems.Add(, , rsData!ProductCode, 1, 1)
                    LS.SubItems(1) = rsData!Description
                    LS.SubItems(2) = rsData!SupplierName
                    LS.SubItems(3) = rsData!Qty
                    LS.SubItems(4) = Format(rsData!SellingPrice, "###,#0.00")
                    LS.SubItems(5) = rsData!DateCreated
        .MoveNext
       Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    Wend
    End With
End Sub

 

Private Sub cmdButtons_Click(Index As Integer)
Select Case Index

Case 0
On Error Resume Next
Command "New"

Case 1
On Error Resume Next
Command "Edit"

Case 2
On Error Resume Next
Command "Delete"

Case 3
On Error Resume Next
Command "Refresh"

Case 4
On Error Resume Next
Command "Close"


End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()

Call LoadEntries
End Sub

 

Private Sub Form_Unload(Cancel As Integer)
     MDIForm1.mPrint.Enabled = True
     loadForm frmWelcome
End Sub

Private Sub ListView1_DblClick()
Command "Edit"
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then ListView1_Click
End Sub

Private Sub Picture1_Resize()
  cmdExit.Left = Picture1.Width - cmdExit.Width - 23
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

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
               On Error Resume Next
               Command "New"
               
    Case vbKeyF2
               On Error Resume Next
               Command "Edit"
               
    Case vbKeyDelete
               On Error Resume Next
               Command "Delete"
               
    Case vbKeyF5
               On Error Resume Next
               Command "Refresh"
               
    Case 67 And Shift = 2
                On Error Resume Next
                Command "Close"
End Select
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
     MDIForm1.mPrint.Enabled = False
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
                With frmStockReceive
                    .EditFlag = True
                    .Show vbModal
                End With
                
Case "Edit"
           If Trim(srcRecord) = vbNullString Then
                    MsgBox "Invalid selection.Can't proceed to the operation!", vbExclamation
                    Exit Sub
           Else
           If ListView1.ListItems.Count < 1 Then Exit Sub
           Set rsData = New ADODB.Recordset
           rsData.Open "Select * from tblStockReceive where ProductCode = '" & ListView1.SelectedItem.Text & "'", CN, adOpenStatic, adLockPessimistic
           If rsData.RecordCount < 1 Then Exit Sub
            With frmStockReceive
              .EditFlag = False
              .Text2.Text = rsData.Fields("ProductCode")
              .Text1.Text = rsData.Fields("Description")
              .Text3.Text = rsData.Fields("Suppliername")
              .Text4.Text = rsData.Fields("sellingPrice")
              .Text5.Text = rsData.Fields("Qty")
              .PK = srcRecord
              .Show vbModal
            End With
            End If
Case "Delete"
           If ListView1.ListItems.Count < 1 Then
            MsgBox "No Records", vbExclamation
            Exit Sub
            End If
            
            If Trim(srcRecord) = vbNullString Then
            MsgBox "Invalid selection.Can't proceed to the operation!", vbExclamation
            Exit Sub
            End If
            
            Dim wheng As VbMsgBoxResult
            wheng = MsgBox("You are about to delete (1) record." & vbCrLf & _
                     "If you click Yes, you won't be able to undo this delete operation." & _
                    vbCrLf & vbCrLf & _
                     "Are you sure you want to delete this record ?", vbCritical + vbYesNo)
            If wheng = vbYes Then
            CN.Execute ("Delete from  tblStockReceive where ProductCode= '" & ListView1.SelectedItem.Text & "'  ")
            LoadEntries
            MsgBox "Deleteng Records Successfully Made", vbInformation
         End If
         
Case "Refresh"
 Me.LoadEntries
 
Case "Close"
      Unload Me


End Select
End Sub





