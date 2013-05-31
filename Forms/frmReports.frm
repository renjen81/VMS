VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmReports 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Reports"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   7095
      TabIndex        =   10
      Top             =   1510
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Caption         =   "Daily"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00584620&
         Height          =   855
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   4215
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61865985
            CurrentDate     =   40547
         End
         Begin lvButton.lvButtons_H lvButtons_H3 
            Height          =   450
            Left            =   2760
            TabIndex        =   16
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   794
            Caption         =   "Generate"
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
            cBhover         =   16777215
            LockHover       =   3
            cGradient       =   16119285
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            cBack           =   16119285
            mPointer        =   99
            mIcon           =   "frmReports.frx":57E2
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Caption         =   "Weekly-Monthly"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00584620&
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   6615
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   450
            Left            =   5280
            TabIndex        =   14
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   794
            Caption         =   "Generate"
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
            cBhover         =   16777215
            LockHover       =   3
            cGradient       =   16119285
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            cBack           =   16119285
            mPointer        =   99
            mIcon           =   "frmReports.frx":5944
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   960
            TabIndex        =   18
            Top             =   320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   61865987
            CurrentDate     =   39859
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3360
            TabIndex        =   19
            Top             =   320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   61865987
            CurrentDate     =   39859
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00584620&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00584620&
            Height          =   195
            Left            =   3000
            TabIndex        =   12
            Top             =   360
            Width           =   180
         End
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H8000000C&
         Height          =   2535
         Left            =   0
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   7095
      TabIndex        =   9
      Top             =   1510
      Width           =   7095
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H8000000C&
         Height          =   2535
         Left            =   0
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   7095
      TabIndex        =   6
      Top             =   1510
      Width           =   7095
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H8000000C&
         Height          =   2535
         Left            =   0
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   3
      Top             =   742
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
      ScaleWidth      =   529
      TabIndex        =   0
      Top             =   0
      Width           =   7935
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
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmReports.frx":5AA6
         Top             =   120
         Width           =   480
      End
   End
   Begin lvButton.lvButtons_H cmdFile 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   661
      Caption         =   "Product"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   9069372
      cBhover         =   13016952
      LockHover       =   3
      cGradient       =   -2147483628
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483629
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   2610
      TabIndex        =   7
      Top             =   1080
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   661
      Caption         =   "Stock Monitoring"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   9069372
      cBhover         =   13016952
      LockHover       =   3
      cGradient       =   -2147483628
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483629
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   5000
      TabIndex        =   8
      Top             =   1080
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   661
      Caption         =   "Sales"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   9069372
      cBhover         =   13016952
      LockHover       =   3
      cGradient       =   -2147483628
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483629
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   240
      Picture         =   "frmReports.frx":66EA
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make back-up on database for reference and to secure data from loss."
      Height          =   600
      Left            =   855
      TabIndex        =   4
      Top             =   135
      Width           =   2400
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS   As ADODB.Recordset
Public rsData   As ADODB.Recordset
Public SQL      As String

Private Sub cmdFile_Click()
Me.Picture3.Visible = True
Me.Picture3.Visible = False
Me.Picture4.Visible = False

End Sub

Private Sub cmdSave_Click()
Set rsData = New ADODB.Recordset
rsData.Open "SELECT * FROM tblSalesInvoice WHERE DateCreated BETWEEN #" & Me.DTPicker1.Value & "# AND #" & Me.DTPicker2.Value & "#", CN, adOpenStatic, adLockPessimistic
Set rptDailySales.DataSource = rsData
rptDailySales.Show 1

End Sub


Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rsData = Nothing
End Sub

Private Sub lvButtons_H1_Click()
Me.Picture3.Visible = True
Me.Picture1.Visible = False
Me.Picture4.Visible = False
End Sub

Private Sub lvButtons_H2_Click()
Me.Picture4.Visible = True
Me.Picture3.Visible = False
Me.Picture1.Visible = False
End Sub

Private Sub lvButtons_H3_Click()
Set rsData = New ADODB.Recordset
rsData.Open "SELECT * FROM tblSalesInvoice order by [DateCreated] ASC", CN, adOpenStatic, adLockPessimistic
Set rptDailySales.DataSource = rsData
rptDailySales.Show 1
End Sub

