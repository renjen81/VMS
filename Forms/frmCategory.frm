VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCategory 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Entry"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   Icon            =   "frmCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   2640
      Width           =   5775
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
      Height          =   885
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "Description"
      Top             =   1560
      Width           =   4530
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
      TabIndex        =   0
      Tag             =   "Category Name"
      Top             =   1200
      Width           =   4530
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
      TabIndex        =   11
      Top             =   840
      Width           =   1890
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   10
      Top             =   720
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
      ScaleWidth      =   403
      TabIndex        =   7
      Top             =   0
      Width           =   6045
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
         TabIndex        =   9
         Top             =   510
         Width           =   3900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay"
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
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmCategory.frx":57E2
         Top             =   120
         Width           =   480
      End
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   325
      Left            =   3720
      TabIndex        =   2
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
      Height          =   325
      Left            =   5040
      TabIndex        =   3
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay Code"
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
      TabIndex        =   6
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay Name"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1560
      Width           =   795
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public State            As FormState
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
 
With frmCategorys.rsData
    If EditFlag = True Then .AddNew
        .Fields("BrgyCode") = Text1.Text
        .Fields("BrgyName") = Text2.Text
        .Fields("Description") = Text3.Text
      .Update
End With

    If EditFlag = True Then
        Call FindRec(frmCategorys.rsData, "BrgyCode", True, Text1.Text, 0)
        MsgBox "New record has been successfully added.", vbInformation
        Dim REPLY As Integer
        REPLY = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
        If REPLY = vbYes Then
            Resetfields
        Else
            frmCategorys.LoadEntries
            Unload Me
        End If
    Else
        Call FindRec(frmCategorys.rsData, "BrgyCode", True, Text1.Text, 0)
        MsgBox "Changes in record has been successfully saved.", vbInformation
        frmCategorys.LoadEntries
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
frmCategorys.LoadEntries
End Sub
Sub GeneratePK()
Dim iCode  As Long
    iCode = getIndex("tblBrgy")
    Text1.Text = GenerateID(iCode, "BRGY-", "00000")
End Sub
 
Private Sub Form_Load()
If EditFlag = False Then
        Caption = "Edit Existing"
Else
        clearText Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If EditFlag = True Then frmCategorys.LoadEntries
Set rsData = Nothing
End Sub

 
Private Sub Text2_LostFocus()
ToUpper Text2
End Sub
