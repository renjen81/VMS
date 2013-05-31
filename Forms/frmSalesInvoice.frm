VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSalesInvoice 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17955
   ControlBox      =   0   'False
   Icon            =   "frmSalesInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   17955
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8895
      Left            =   120
      TabIndex        =   58
      Top             =   840
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
         TabIndex        =   59
         Top             =   3960
         Width           =   13815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   4200
      TabIndex        =   51
      Top             =   7920
      Width           =   5895
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000008&
         Height          =   495
         Left            =   -480
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00004080&
         Height          =   495
         Left            =   1800
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   53
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H000000FF&
         Height          =   495
         Left            =   4560
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   52
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   75
         TabIndex        =   57
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOder Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   56
         Top             =   360
         Width           =   2130
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Out of Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   55
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture234 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   8835
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1710
         TabIndex        =   1
         Top             =   2895
         Width           =   5295
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   300
         Left            =   7080
         TabIndex        =   2
         Top             =   2895
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         Caption         =   "Select"
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
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   300
         Left            =   8025
         TabIndex        =   3
         Top             =   2895
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "Refresh"
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   15
         TabIndex        =   4
         Top             =   240
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilList"
         SmallIcons      =   "ilList"
         ForeColor       =   8399906
         BackColor       =   16777215
         Appearance      =   0
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
         MouseIcon       =   "frmSalesInvoice.frx":57E2
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   3508
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7480
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UnitPrice"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Selling Price"
            Object.Width           =   2170
         EndProperty
      End
      Begin VB.Image cmdExit 
         Height          =   360
         Left            =   8505
         Picture         =   "frmSalesInvoice.frx":5944
         ToolTipText     =   "Close"
         Top             =   -30
         Width           =   360
      End
      Begin VB.Label lblVal 
         BackStyle       =   0  'Transparent
         Caption         =   "   Product Records"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   45
         Width           =   2175
      End
      Begin VB.Label lblVal 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Decription"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   2940
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   17775
      TabIndex        =   27
      Top             =   8880
      Width           =   17775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   19215
      TabIndex        =   26
      Top             =   720
      Width           =   19215
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
      ScaleWidth      =   1197
      TabIndex        =   19
      Top             =   0
      Width           =   17955
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5160
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6840
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmSalesInvoice.frx":602E
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voters"
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
         TabIndex        =   21
         Top             =   180
         Width           =   930
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
         TabIndex        =   20
         Top             =   510
         Width           =   3900
      End
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   7560
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   15840
      TabIndex        =   8
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   114819075
      CurrentDate     =   40541
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   3120
      Top             =   480
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
            Picture         =   "frmSalesInvoice.frx":6C72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5655
      Left            =   120
      TabIndex        =   22
      Top             =   1395
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ilList"
      SmallIcons      =   "ilList"
      ForeColor       =   8388608
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
      MouseIcon       =   "frmSalesInvoice.frx":720C
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice No"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty/Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   120
      ScaleHeight     =   7860
      ScaleWidth      =   17775
      TabIndex        =   9
      Top             =   840
      Width           =   17775
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   720
         Left            =   15000
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   5400
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Height          =   720
         Left            =   15000
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   6960
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Height          =   720
         Left            =   15000
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   6165
         Width           =   2655
      End
      Begin VB.TextBox TxtDesc 
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
         Left            =   33
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "Description"
         Top             =   210
         Width           =   6300
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   795
         Left            =   15120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   795
         Left            =   15120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   1680
         Width           =   2505
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   795
         Left            =   15120
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "Quantity"
         Text            =   "0"
         Top             =   720
         Width           =   2535
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   345
         Left            =   16320
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Caption         =   "S&old"
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
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   14040
         TabIndex        =   36
         Top             =   6480
         Width           =   465
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   14040
         TabIndex        =   35
         Top             =   5520
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14040
         TabIndex        =   34
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00F5F5F5&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   33
         TabIndex        =   18
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   13800
         TabIndex        =   17
         Top             =   3000
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Price (Each)"
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
         Left            =   13800
         TabIndex        =   16
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   13800
         TabIndex        =   15
         Top             =   1080
         Width           =   630
      End
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   855
      Left            =   15225
      TabIndex        =   37
      Top             =   9840
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Caption         =   "&Print"
      CapAlign        =   2
      BackStyle       =   3
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
      ImgAlign        =   4
      Image           =   "frmSalesInvoice.frx":736E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   855
      Left            =   16935
      TabIndex        =   38
      Top             =   9840
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   3
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
      ImgAlign        =   4
      Image           =   "frmSalesInvoice.frx":A5A7
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   855
      Left            =   16080
      TabIndex        =   39
      Top             =   9840
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   3
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
      ImgAlign        =   4
      Image           =   "frmSalesInvoice.frx":BF39
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
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
      Height          =   240
      Index           =   8
      Left            =   17280
      TabIndex        =   50
      Top             =   10800
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F11"
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
      Height          =   240
      Index           =   7
      Left            =   16320
      TabIndex        =   49
      Top             =   10800
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
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
      Height          =   240
      Index           =   6
      Left            =   15480
      TabIndex        =   48
      Top             =   10800
      Width           =   330
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
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
      Height          =   240
      Left            =   8640
      TabIndex        =   47
      Top             =   11520
      Width           =   210
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
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
      Height          =   240
      Left            =   9840
      TabIndex        =   46
      Top             =   11520
      Width           =   210
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
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
      Height          =   240
      Left            =   10920
      TabIndex        =   45
      Top             =   11520
      Width           =   210
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
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
      Height          =   240
      Left            =   11880
      TabIndex        =   44
      Top             =   11520
      Width           =   210
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
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
      Height          =   240
      Left            =   13080
      TabIndex        =   43
      Top             =   11520
      Width           =   330
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F11"
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
      Height          =   240
      Left            =   13920
      TabIndex        =   42
      Top             =   11520
      Width           =   330
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
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
      Height          =   240
      Left            =   14760
      TabIndex        =   41
      Top             =   11520
      Width           =   330
   End
   Begin VB.Label Labels 
      Caption         =   "Note:Press ESC to cancel transaction."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   4
      Left            =   15000
      TabIndex        =   40
      Top             =   9195
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   14520
      Picture         =   "frmSalesInvoice.frx":C813
      Top             =   9120
      Width           =   360
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   9840
      TabIndex        =   30
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   7800
      TabIndex        =   25
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   14520
      TabIndex        =   24
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   7320
      Width           =   975
   End
End
Attribute VB_Name = "frmSalesInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcRecord               As String
Dim srcRecCD                As Variant
Public rsData               As ADODB.Recordset
Public tempx                As Integer
 

Sub LoadEntries()
tempx = 0
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblProduct ORDER by productCode ASC", CN, adOpenStatic, adLockPessimistic
With rsData
    While Not .EOF
                Dim LS As ListItem
                Set LS = ListView1.ListItems.Add(, , rsData!ProductCode, 1, 1)
                LS.SubItems(1) = rsData!Description
                LS.SubItems(2) = rsData!Qty
                LS.SubItems(3) = Format(rsData!UnitPrice, "###,#0.00")
                LS.SubItems(4) = Format(rsData!SellingPrice, "###,#0.00")
                tempx = tempx + 1
           
            If !QtyRemain = 0 Then
               ListView1.ListItems(tempx).ForeColor = vbRed
               ListView1.ListItems(tempx).ListSubItems(1).ForeColor = vbRed
               ListView1.ListItems(tempx).ListSubItems(4).ForeColor = vbRed
            ElseIf !QtyRemain <= 5 Then
                ListView1.ListItems(tempx).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(1).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(4).ForeColor = &H4080&
            End If
        .MoveNext
    Wend
    End With
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub



Private Sub cmdExit_Click()
Picture234.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

 


Private Sub cmdSave_Click()
If is_empty(TxtDesc, True) = True Then Exit Sub
If is_empty(txtQty, True) = True Then Exit Sub
Dim zSI, zDesc, zqty, zprice, ztotal, zDate, zAdd As String
Dim tAmt As Double
Dim i, Qr, Qr2, Qr3 As Integer

If Text2.Text = "0.00" Then
         MsgBox "Please enter Amount..", vbInformation
         Exit Sub
End If


If ListView2.ListItems.Count < 1 Then
         MsgBox "There is no record..", vbInformation
         Exit Sub
End If

If Text5.Text = "0.00" Then
         MsgBox "Don't forget to pay pay..", vbInformation
         Exit Sub
End If

For i = 1 To ListView2.ListItems.Count
     zSI = ListView2.ListItems.Item(i).Text
     zDesc = ListView2.ListItems(i).ListSubItems(1).Text
     zqty = ListView2.ListItems(i).ListSubItems(2).Text
     zprice = ListView2.ListItems(i).ListSubItems(3).Text
     ztotal = ListView2.ListItems(i).ListSubItems(4).Text
     zDate = ListView2.ListItems(i).ListSubItems(5).Text

'-----------Minus Liters---------------
Set rsData = New ADODB.Recordset
rsData.Open "Select * From tblProduct Where ProductCode= '" & Text4.Text & "'", CN, adOpenStatic, adLockOptimistic
Qr3 = rsData.Fields("QtySold")
    Qr = rsData![QtyRemain]
    Qr2 = txtQty.Text
    rsData![QtySold] = Qr3 + 1
    rsData![QtyRemain] = Qr - Qr2
rsData.Update
Set rsData = Nothing

'-----------End----------------------
Set rsData = New ADODB.Recordset
rsData.Open "INSERT INTO tblSalesInvoice(InvoiceNO,ProductCode,Description,Qty,Price,TotalAmount,Remarks,DateCreated,CreatedBy)" & _
        "values ('" & zSI & "','" & Text4 & "','" & zDesc & "','" & zqty & "','" & zprice & _
        "','" & ztotal & "','" & txtRemarks.Text & " ','" & zDate & " ','" & CurrUser.UserNAME & "')", CN, adOpenStatic, adLockOptimistic
Next i
MsgBox "New record has been successfully saved.", vbInformation
If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
    Resetfields
    Text7.Text = "0.00"
  Else
     Unload Me
End If
End Sub

Sub Resetfields()
ListView2.ListItems.Clear
TxtDesc.Text = ""
txtQty.Text = ""
txtPrice.Text = "0.00"
txtAmount.Text = "0.00"
Text4.Text = "0.00"
Text2.Text = "0.00"
Text5.Text = "0.00"
Text4.Text = ""
Text6.Text = ""
GeneratePK
End Sub
 

Private Sub Form_Load()
lvButtons_H1.Enabled = False
GeneratePK
DTPicker1.Value = Date
End Sub

 Sub GeneratePK()
Dim iCode  As Long
    iCode = getIndex("tblSalesInvoice")
    Text1 = GenerateID(iCode, Format(Now, "yyyy-"), "00000")
End Sub

 
 
 

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmWelcome.LOAD_MY_URL
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsData = Nothing
 
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortLV ListView1
End Sub
 
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcRecCD = ListView1.SelectedItem.Index
srcRecord = ListView1.ListItems.Item(srcRecCD).Text
End Sub

Private Sub ListView2_Click()
If ListView2.ListItems.Count < 1 Then Exit Sub
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortLV ListView2
End Sub

Private Sub ListView1_DblClick()
If Trim(srcRecord) = vbNullString Then
     MsgBox "Please select a record from the list .Can't proceed to the operation!", vbExclamation
Else
    Text4.Text = ListView1.SelectedItem.Text
    TxtDesc = ListView1.SelectedItem.ListSubItems(1).Text
    Text6 = ListView1.SelectedItem.ListSubItems(2).Text
    txtPrice = ListView1.SelectedItem.ListSubItems(4).Text
    Picture234.Visible = False
    txtQty.SetFocus
'    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub ListView2_DblClick()
Dim x As Integer
Dim lngTot As Double
Dim lngIndex As Long

x = ListView2.SelectedItem.Index
ListView2.ListItems.Remove (x)

        For lngIndex = 1 To ListView2.ListItems.Count
            lngTot = lngTot + Me.ListView2.ListItems(lngIndex).SubItems(4)
        Next lngIndex

Me.Text5.Text = lngTot
End Sub

Private Sub lvButtons_H1_Click()
 If TxtDesc.Text = "" Then TxtDesc.SetFocus: Exit Sub
 If txtQty.Text = "0" Then txtQty.SetFocus: Exit Sub

If txtQty.Text = "" Then
    MsgBox "Please enter  Quantity", vbExclamation
    txtQty.SetFocus
    Exit Sub
End If

Dim xQty, xAmount, xPrice  As Double
xQty = txtQty.Text
xPrice = txtPrice.Text

txtAmount = xQty * xPrice
 
With ListView2
    .ListItems.Add , , Text1, , 1
    .ListItems(.ListItems.Count).ListSubItems.Add , , TxtDesc.Text
    .ListItems(.ListItems.Count).ListSubItems.Add , , txtQty.Text
    .ListItems(.ListItems.Count).ListSubItems.Add , , toCurr((txtPrice))
    .ListItems(.ListItems.Count).ListSubItems.Add , , toCurr((txtAmount))
    .ListItems(.ListItems.Count).ListSubItems.Add , , Format(DTPicker1.Value, "MMM-dd-yyyy")
  
'------Computation------------'
xPrice = Format$(CDbl(xPrice * xQty))
Text5 = Format(Text5 + Val(xPrice))
End With
'-------------End --------------------
End Sub

Private Sub lvButtons_H2_Click()
ListView1_DblClick
End Sub

Private Sub lvButtons_H4_Click()
LoadEntries
Text3.Text = ""
End Sub

Private Sub Text2_Change()
 Text7.Text = toMoney((toNumber(Text2.Text) - toNumber(Text5.Text)))
End Sub

 

Private Sub Text2_GotFocus()
HighL Text2
End Sub

Private Sub Text2_LostFocus()
HighL Text2
End Sub

Private Sub Text3_Change()
tempx = 0
ListView1.ListItems.Clear
Set rsData = New ADODB.Recordset
Set rsData = CN.Execute("Select * from tblProduct where Description like '" & Text3.Text & "%'")
With rsData
    While Not .EOF
                Dim x As ListItem
                Set x = ListView1.ListItems.Add(, , rsData!ProductCode, 1, 1)
                        x.SubItems(1) = rsData!Description
                        x.SubItems(2) = rsData!Qty
                        x.SubItems(3) = Format(rsData!UnitPrice, "###,#0.00")
                        x.SubItems(4) = Format(rsData!SellingPrice, "###,#0.00")
                   tempx = tempx + 1
           
            If !QtyRemain = 0 Then
               ListView1.ListItems(tempx).ForeColor = vbRed
               ListView1.ListItems(tempx).ListSubItems(1).ForeColor = vbRed
               ListView1.ListItems(tempx).ListSubItems(4).ForeColor = vbRed
            
            ElseIf !QtyRemain <= 5 Then
                ListView1.ListItems(tempx).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(1).ForeColor = &H4080&
                ListView1.ListItems(tempx).ListSubItems(4).ForeColor = &H4080&
 
            End If
           
        .MoveNext
     Wend
End With
End Sub

 

Private Sub txtAmount_Change()
txtAmount.Text = Format(txtAmount, "###,#0.00")
End Sub

Private Sub TxtDesc_Change()
Me.Text3.Text = Me.TxtDesc.Text
End Sub

Private Sub TxtDesc_Click()
Call LoadEntries
If Picture234.Visible = True Then
    Picture234.Visible = False
Else
    Picture234.Visible = False
With Picture234
   .Top = 2700
   .Left = 120
   .Visible = True
End With
End If
Me.Text3.Text = Me.TxtDesc.Text
End Sub
 
Private Sub txtPrice_Change()
txtPrice.Text = Format(txtPrice, "###,#0.00")
End Sub

Private Sub txtQty_Change()
If txtQty.Text = "" Then
    lvButtons_H1.Enabled = False
Exit Sub
End If

Set rsData = New ADODB.Recordset
rsData.Open "Select * From tblProduct Where ProductCOde= '" & Text4.Text & "'", CN, adOpenStatic, adLockOptimistic
With rsData
    If .EOF = True And .BOF = True Then
    Else
        If Val(!QtyRemain) = 0 Then
            MsgBox "" & !Description & " is out of Stock", vbCritical
             txtQty.SetFocus
             txtQty.SelStart = 0
             txtQty.SelLength = Len(txtQty.Text)
             lvButtons_H1.Enabled = False
        Else
            If Val(txtQty.Text) > !QtyRemain Then
                MsgBox "" & !Description & " quantity needed is greater than Stock", vbExclamation
                txtQty.SetFocus
                txtQty.SelStart = 0
                txtQty.SelLength = Len(txtQty.Text)
                lvButtons_H1.Enabled = False
            Else
            lvButtons_H1.Enabled = True
            End If
        End If
    End If
    
End With
Set rsData = Nothing
txtAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(toNumber(txtPrice.Text))))
 
End Sub

 

Private Sub txtQty_GotFocus()
   With txtQty
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    'KeyAscii = isNumber(KeyAscii)
    If KeyAscii = 13 Then
        lvButtons_H1_Click
    End If
End Sub


 
 
