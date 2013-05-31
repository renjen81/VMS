VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDBBackup 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Database"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "frmDBBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   330
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
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
   Begin lvButton.lvButtons_H cmdBackup 
      Default         =   -1  'True
      Height          =   330
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Caption         =   "&Create Backup"
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
      ScaleWidth      =   323
      TabIndex        =   1
      Top             =   0
      Width           =   4845
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmDBBackup.frx":57E2
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
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
         Width           =   2445
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
   Begin VB.Label lblCBK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   0
      Picture         =   "frmDBBackup.frx":6426
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make back-up on database for reference and to secure data from loss."
      Height          =   600
      Left            =   855
      TabIndex        =   6
      Top             =   135
      Width           =   2400
   End
End
Attribute VB_Name = "frmDBBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents clsBKU As clsHuffman
Attribute clsBKU.VB_VarHelpID = -1

Private Sub cmdBackup_Click()
    cmdBackup.Enabled = False
    cmdCancel.Enabled = False
    lblCBK.Caption = "Creating Database Backup..."
    BackUpDB
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

 

Private Sub clsBKU_Progress(Procent As Integer)

    progStat.Value = Procent

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsBKU = Nothing
End Sub

Public Sub BackUpDB()
On Error GoTo err:
    Dim FSO As New FileSystemObject
    
    Dim sDBFN As String
    Dim sDBTmpFN As String
    
    If FSO.FolderExists(App.Path & "/Backup") = False Then
        FSO.CreateFolder App.Path & "/Backup"
    End If
    
    'set backup file path filename
    sDBFN = App.Path & "/Backup/" & Format$(Date, "MMM-dd-yyyy") & ".mdb"
    
    'set temporary file
    sDBTmpFN = sDBFN & Now - DateValue(Now) & GetTickCount
    
    If FSO.FileExists(sDBTmpFN) = True Then
        FSO.DeleteFile sDBTmpFN
    End If
    
    'show ctl
    progStat.Visible = True
    lblCBK.Visible = True
    DoEvents
    
    'start backup
    Set frmDBBackup.clsBKU = New clsHuffman
    frmDBBackup.clsBKU.EncodeFile DBPathFileName, sDBTmpFN
    
    'rename file
    If FSO.FileExists(sDBFN) = True Then
        FSO.DeleteFile sDBFN
    End If
    FSO.MoveFile sDBTmpFN, sDBFN
    
    
    Set FSO = Nothing

    lblCBK.Caption = "Backup Complete"
    MsgBox "Database were back-up successfully!", vbInformation, "Back-up Detail"
    cmdCancel.Enabled = True
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

 

