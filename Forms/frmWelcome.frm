VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmWelcome 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC8661&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   538
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      Begin VB.Timer timerUT 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   6480
         Top             =   1320
      End
      Begin lvButton.lvButtons_H cmdB 
         Height          =   345
         Index           =   1
         Left            =   2265
         TabIndex        =   1
         Top             =   1635
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Caption         =   "About"
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
         cFHover         =   0
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   13403745
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   13403745
      End
      Begin lvButton.lvButtons_H cmdB 
         Height          =   345
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   1635
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Caption         =   "Quick"
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
         cFHover         =   0
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   13403745
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   13403745
      End
      Begin VB.PictureBox bgB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEFE1&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3525
         Index           =   1
         Left            =   180
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   411
         TabIndex        =   4
         Top             =   1980
         Visible         =   0   'False
         Width           =   6165
         Begin VB.PictureBox bgMe 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   130
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   1950
            Begin VB.Timer timerAniIn 
               Enabled         =   0   'False
               Interval        =   10
               Left            =   1215
               Top             =   2475
            End
            Begin VB.Image me2 
               Height          =   3540
               Left            =   0
               Picture         =   "frmWelcome.frx":0000
               Top             =   0
               Visible         =   0   'False
               Width           =   1950
            End
            Begin VB.Image me1 
               Height          =   3540
               Left            =   0
               Picture         =   "frmWelcome.frx":62AF
               Top             =   0
               Width           =   1950
            End
         End
         Begin lvButton.lvButtons_H cmdPW 
            Height          =   345
            Left            =   720
            TabIndex        =   6
            Top             =   1680
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   609
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   12648447
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16773089
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright© 2013. All Rights Reserved."
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   570
            Width           =   3975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Urbiztondo Voters Monitoring System"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2595
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact : "
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Developed by: Renje M. Nituda"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   2235
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "renjen81@gmail.com"
            Height          =   255
            Left            =   720
            TabIndex        =   12
            Top             =   1410
            Width           =   3975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " (+63) 9169178903                 "
            Height          =   195
            Left            =   720
            TabIndex        =   11
            Top             =   1200
            Width           =   2115
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Created   :"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   2490
            Width           =   1215
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Project Title     : Urbiztondo Voters Monitoring System"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   2280
            Width           =   3750
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "This system provide basic functionality in a certain Insurance. It can add,  delete , update .It can generate report."
            Height          =   855
            Left            =   240
            TabIndex        =   8
            Top             =   2850
            Width           =   4215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "May-2013"
            Height          =   195
            Left            =   1440
            TabIndex        =   7
            Top             =   2520
            Width           =   705
         End
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   5400
         Top             =   960
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
               Picture         =   "frmWelcome.frx":D7A5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox bgB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEFE1&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3495
         Index           =   0
         Left            =   180
         ScaleHeight     =   233
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   411
         TabIndex        =   3
         Top             =   1980
         Visible         =   0   'False
         Width           =   6165
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   3240
            Left            =   -30
            TabIndex        =   24
            Top             =   -120
            Width           =   5895
            ExtentX         =   10398
            ExtentY         =   5715
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address Here!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   165
         Left            =   180
         TabIndex        =   25
         Top             =   405
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name Here!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   120
         Width           =   1125
      End
      Begin VB.Image Image2 
         Height          =   1125
         Left            =   0
         Picture         =   "frmWelcome.frx":DD3F
         Stretch         =   -1  'True
         Top             =   -360
         Width           =   28560
      End
      Begin VB.Label lblSchoolAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFAEA&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   465
         Width           =   120
      End
      Begin VB.Label lblSchoolName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   225
         TabIndex        =   21
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label lblPreOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   465
         TabIndex        =   20
         Top             =   1275
         Width           =   120
      End
      Begin VB.Label lblIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   465
         TabIndex        =   19
         Top             =   1065
         Width           =   120
      End
      Begin VB.Label lblUserNames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1350
         TabIndex        =   18
         Top             =   810
         Width           =   180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   225
         TabIndex        =   17
         Top             =   810
         Width           =   1035
      End
      Begin VB.Image Image3 
         Height          =   25200
         Left            =   0
         Picture         =   "frmWelcome.frx":DE7B
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Public rsData  As ADODB.Recordset
Dim connString          As String
 

Private Sub cmdB_Click(Index As Integer)
    Dim i As Integer
    cmdB(Index).GradientColor = &HFFEFE1
    bgB(Index).Visible = True
    
    For i = 0 To cmdB.UBound
        If i <> Index Then
            cmdB(i).GradientColor = cmdB(i).BackColor
            bgB(i).Visible = False
        End If
    Next
    
    ReArrangeControls
    
    Select Case Index
        Case 1 '
            StartShowAbout
    End Select
End Sub

Private Sub cmdPW_Click()
OpenURL "www.sherwin-salazar.qapacity.com", hwnd
End Sub
 

Public Sub StartShowAbout()
    me2.Visible = False
    bgMe.Move bgB(1).Width - bgMe.Width, 0, bgMe.Width, 1
    bgMe.Visible = True
    timerAniIn.Enabled = True
End Sub

Private Sub Form_Load()
Call LOAD_MY_URL
Label12.Caption = CurrBusinessInfo.BusinessName
Label13.Caption = CurrBusinessInfo.BusinessAddress
End Sub

Private Sub Form_Resize()
 
    ReArrangeControls
End Sub
  
  Sub ReArrangeControls()
      Dim preLeft As Integer
    Dim i As Integer
    
  On Error Resume Next

    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2

    lblCurrentTime.Left = bgMain.Width - lblCurrentTime.Width - 5

    For i = 0 To bgB.UBound
        If bgB(i).Visible = True Then
            bgB(i).Move bgB(i).Left, bgB(i).Top, bgMain.Width - bgB(i).Left - 4, bgMain.Height - bgB(i).Top - 4
        End If
    Next
    
    If bgB(0).Visible = True Then
        WebBrowser1.Width = bgB(0).Width - WebBrowser1.Left * 2
        WebBrowser1.Height = bgB(0).Height - WebBrowser1.Top - WebBrowser1.Left * 2
     End If

    If bgB(1).Visible = True Then
        bgMe.Move bgB(1).Width - bgMe.Width, 0, bgMe.Width
    End If
  End Sub

 
 
 
Private Sub Form_Unload(Cancel As Integer)
Set RS = Nothing
Set RSInf = Nothing
End Sub

 
 
Private Sub timerUT_Timer()
    lblIn.Caption = "Today is " & FormatDateTime(Now, vbLongDate)
    lblPreOut.Caption = "The time is: " & Time

End Sub

 

Sub LOAD_MY_URL()
On Error GoTo errorTrap
    Dim FLD     As ADODB.Field
    Dim i       As Long

   connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & _
           "\DB.mdb;User Id=admin;Password="

    Set rsData = New ADODB.Recordset
    rsData.Open "SELECT * FROM Qry_Inventory", CN, adOpenStatic, adLockOptimistic

 
      Me.WebBrowser1.Navigate2 "About:Blank"
      Do While Me.WebBrowser1.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    With WebBrowser1.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Tahoma;} body,td{font-size:11px;}</style>")
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write ("<TABLE WIDTH=100% CELLPADDING=3 CELLSPACING=3>")
        .Write ("<TR><TD><b>Quick Inventory Records </b></TD>")
        
        .Write ("<TABLE WIDTH=6.4% BORDER=1 BORDERCOLOR=&H00CC8661& CELLPADDING=1 CELLSPACING=0>")
        .Write ("<TR><TD bgcolor=#D6E9F7>Total Supplier          <TD BGCOLOR=#FFFFD2><CENTER><b>" & rsData.Fields("totalSupp") & "</TD></TD>")
        .Write ("<TR><UL><TD bgcolor=#D6E9F7>Total Products      <TD BGCOLOR=#FFFFD2><CENTER><b>" & rsData.Fields("totalProduct") & "</TD></TD>")
        .Write ("<TR><UL><TD bgcolor=#D6E9F7>Total Sales         <TD BGCOLOR=#FFFFD2><CENTER><b>" & Format(rsData.Fields("TotalSales"), "#,#00.00") & "</TD></TD>")
        .Write ("</TR></TABLE>")
        .Write ("<BR>")
        
        
        .Write ("<TABLE WIDTH=100% BORDER=0 BORDERCOLOR=#000000 CELLPADDING=0 CELLSPACING=0>")
        .Write ("<TABLE WIDTH=100% BORDER=0 BORDERCOLOR=#215DC6 CELLPADDING=1 CELLSPACING=1>")
        .Write ("<TR><TD BGCOLOR=#D6E9F7><B>LPPA</B></TD>")
        .Write ("<TD BGCOLOR=#D6E9F7><B>Total Supplier</B></TD>")
        .Write ("<TD BGCOLOR=#D6E9F7><B>Total Products</B></TD>")
        .Write ("<TD BGCOLOR=#D6E9F7><B>Total Sales</B></TD>")
        .Write ("<TD BGCOLOR=#D6E9F7><B>DATE</B></TD></TR>")
         

        While RS.EOF <> True
        i = i + 1
            For Each FLD In RS.Fields
                If i Mod 2 <> 0 Then
                    .Write ("<TD bgcolor=#EBF0FC><FONT FACE=tahoma SIZE=1>" & FLD.Value & "</TD>")
                Else
                    .Write ("<TD bgcolor=#FFFFD2><FONT FACE=tahoma SIZE=1>" & FLD.Value & "</TD>")
                End If
            Next FLD
            .Write ("</TR>")
            RS.MoveNext
        Wend
        .Write ("</TABLE></TD></TR>")
        .Write ("</TABLE></BODY></HTML>")
        WebBrowser1.Document.Script.Document.Clear
        WebBrowser1.Document.Script.Document.Close
    End With
    
 
errorTrap:
   If err.Number = -2147217900 Then
      MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbExclamation + vbOKOnly
    End If
End Sub


