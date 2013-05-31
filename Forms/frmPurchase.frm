VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPurchaseOrder 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Entry"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
   Icon            =   "frmPurchase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete Item"
      Height          =   315
      Left            =   4320
      TabIndex        =   65
      Top             =   8520
      Width           =   1080
   End
   Begin VB.TextBox txtcode 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   3000
      Width           =   570
   End
   Begin VB.PictureBox Picture234 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5715
      TabIndex        =   56
      Top             =   3960
      Visible         =   0   'False
      Width           =   5715
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
         Left            =   1680
         TabIndex        =   57
         Top             =   2880
         Width           =   2175
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   300
         Left            =   3960
         TabIndex        =   58
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
         Left            =   4920
         TabIndex        =   59
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   2535
         Left            =   120
         TabIndex        =   60
         Top             =   285
         Width           =   5565
         _ExtentX        =   9816
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmPurchase.frx":57E2
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   5293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Unit Price"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Selling Price"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
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
         TabIndex        =   62
         Top             =   2940
         Width           =   2175
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
         TabIndex        =   61
         Top             =   45
         Width           =   2175
      End
      Begin VB.Image cmdExit 
         Height          =   360
         Left            =   5400
         Picture         =   "frmPurchase.frx":5944
         ToolTipText     =   "Close"
         Top             =   -30
         Width           =   360
      End
   End
   Begin VB.TextBox txtcontactno 
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
      Left            =   6720
      TabIndex        =   53
      Tag             =   "Unit Price"
      Top             =   1920
      Width           =   2490
   End
   Begin VB.TextBox txtcontactperson 
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
      Left            =   6720
      TabIndex        =   51
      Tag             =   "Unit Price"
      Top             =   1560
      Width           =   2490
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   8280
      ScaleHeight     =   45
      ScaleWidth      =   2295
      TabIndex        =   40
      Top             =   7920
      Width           =   2295
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9180
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7500
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9180
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1425
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9180
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   8070
      Width           =   1425
   End
   Begin VB.TextBox txtGross 
      BackColor       =   &H00E6FFFF&
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
      Height          =   285
      Index           =   2
      Left            =   9180
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9180
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6900
      Width           =   1425
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   11055
      TabIndex        =   14
      Top             =   742
      Width           =   11055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   10695
      TabIndex        =   13
      Top             =   3000
      Width           =   10695
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
      TabIndex        =   4
      Top             =   1560
      Width           =   1890
   End
   Begin VB.TextBox txtaddress 
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
      Height          =   525
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   2280
      Width           =   4170
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
      Left            =   4320
      TabIndex        =   3
      Top             =   8040
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.TextBox txtremark 
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
      Height          =   1485
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   6960
      Width           =   3810
   End
   Begin VB.ComboBox txtsupplier 
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
      Tag             =   "Supplier"
      Top             =   1920
      Width           =   3090
   End
   Begin lvButton.lvButtons_H cmdrecieve 
      Height          =   330
      Left            =   5520
      TabIndex        =   15
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "&Recieved Item"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   330
      Left            =   9720
      TabIndex        =   16
      Top             =   8520
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1200
      TabIndex        =   28
      Top             =   2415
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   112984067
      CurrentDate     =   38207
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   120
      ScaleHeight     =   630
      ScaleWidth      =   10740
      TabIndex        =   21
      Top             =   3360
      Width           =   10740
      Begin VB.ComboBox txtunit 
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
         Left            =   3600
         TabIndex        =   63
         Tag             =   "Supplier"
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtitem 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   2490
      End
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   7680
         TabIndex        =   27
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9720
         TabIndex        =   25
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   2775
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   4950
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7620
         TabIndex        =   47
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Index           =   1
         Left            =   8640
         TabIndex        =   46
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3600
         TabIndex        =   45
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Items/Stocks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   44
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   240
         Index           =   9
         Left            =   4980
         TabIndex        =   43
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Qty"
         Height          =   240
         Index           =   10
         Left            =   2760
         TabIndex        =   42
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   6255
         TabIndex        =   41
         Top             =   0
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   48
      Top             =   3960
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIco"
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
      MouseIcon       =   "frmPurchase.frx":602E
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Code"
         Object.Width           =   3070
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Unit Price"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Gross"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Discount%"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Net Amout"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "totaldisc"
         Object.Width           =   0
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   330
      Left            =   6960
      TabIndex        =   55
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "&Print Preview"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdsave 
      Default         =   -1  'True
      Height          =   330
      Left            =   8400
      TabIndex        =   54
      Top             =   8520
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
         Weight          =   400
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
      ScaleWidth      =   733
      TabIndex        =   5
      Top             =   0
      Width           =   10995
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6720
         TabIndex        =   20
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   7200
         TabIndex        =   19
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
         TabIndex        =   7
         Top             =   510
         Width           =   3900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order"
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
         TabIndex        =   6
         Top             =   180
         Width           =   2190
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmPurchase.frx":6190
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      Height          =   225
      Index           =   3
      Left            =   5520
      TabIndex        =   52
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      Height          =   225
      Index           =   0
      Left            =   5520
      TabIndex        =   50
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vat(0.12)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7080
      TabIndex        =   39
      Top             =   7530
      Width           =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Base"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Index           =   0
      Left            =   7080
      TabIndex        =   38
      Top             =   7230
      Width           =   2040
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7080
      TabIndex        =   37
      Top             =   8100
      Width           =   2040
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7080
      TabIndex        =   36
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7080
      TabIndex        =   35
      Top             =   6930
      Width           =   2040
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   " Date Request"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Caption         =   " Purchase Order Detail"
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
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   10740
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000010&
      Caption         =   " Purchase Order Information"
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
      TabIndex        =   17
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Code"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   5520
      TabIndex        =   11
      Tag             =   "Description"
      Top             =   2400
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
      Left            =   240
      TabIndex        =   10
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5160
      X2              =   5160
      Y1              =   1560
      Y2              =   2880
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
      TabIndex        =   9
      Top             =   1920
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
      Left            =   7800
      TabIndex        =   8
      Top             =   1080
      Width           =   2805
   End
End
Attribute VB_Name = "frmPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcRecord               As String
Dim srcRecCD                As Variant
Public tempx                As Integer
Public State            As FormState
Public EditFlag         As Boolean
Public PK               As String
Dim NetTotal As Double, GrossTotal As Double, Discount As Double, Discount1 As Double
Dim i, ii, iii, n As Integer
Dim rsData1, rsData2, rsData4 As New ADODB.Recordset
Public rsData As ADODB.Recordset

 

Sub dDataCombo()


openRec rsData, "tblSupplier"
While rsData.EOF <> True
            txtsupplier.AddItem rsData![SupplierName]
         rsData.MoveNext
Wend


End Sub

Private Sub btnAdd_Click()
If btnAdd.Caption = "Add" Then

 If toNumber(txtUnitPrice.Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtUnitPrice.SetFocus
        Exit Sub
    End If
If txtitem.Text = "" Then txtitem.SetFocus: Exit Sub
 If txtQty.Text = "0" Then txtQty.SetFocus: Exit Sub

If txtQty.Text = "" Then
    MsgBox "Please enter  Quantity", vbExclamation
    txtQty.SetFocus
    Exit Sub
End If

Dim xQty, xAmount, xPrice  As Double
'xQty = txtQty.Text
'xPrice = txtPrice.Text
Discount = 0
 Discount = txtGross(1).Text * (txtDisc.Text / 100)
'txtAmount = xQty * xPrice
'
With ListView1
    .ListItems.Add , , txtcode.Text, , 1
    .ListItems(.ListItems.Count).ListSubItems.Add , , txtitem.Text
    .ListItems(.ListItems.Count).ListSubItems.Add , , txtQty.Text
    .ListItems(.ListItems.Count).ListSubItems.Add , , txtunit
    .ListItems(.ListItems.Count).ListSubItems.Add , , toCurr((txtUnitPrice))
    .ListItems(.ListItems.Count).ListSubItems.Add , , toCurr((txtGross(1)))
    .ListItems(.ListItems.Count).ListSubItems.Add , , txtDisc
    .ListItems(.ListItems.Count).ListSubItems.Add , , toCurr((txtNetAmount))
    .ListItems(.ListItems.Count).ListSubItems.Add , , toCurr((Discount))
   
'------Computation------------'
End With
'-------------End --------------------
txtitem.Text = ""
txtQty.Text = 0
txtunit.Text = ""
txtUnitPrice.Text = "0.00"
txtGross(1).Text = "0.00"
txtDisc.Text = 0
txtNetAmount.Text = "0.00"
txtitem.SetFocus
'-----Compute--------
NetTotal = 0
Discount1 = 0
i = 0
For i = 1 To ListView1.ListItems.Count
    GrossTotal = GrossTotal + ListView1.ListItems(i).SubItems(5)
    Discount1 = Discount1 + ListView1.ListItems(i).SubItems(8)
Next i
 txtDesc.Text = Format$(Discount1, "#,##0.00")
txtGross(2).Text = Format(GrossTotal, "#,##0.00")
NetTotal = toMoney(txtGross(2).Text) - toMoney(txtDesc.Text)
txtNet.Text = Format(NetTotal, "#,##0.00")


  txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
Else
Dim reply11 As Integer
    reply11 = MsgBox("Do you want to update this record?", vbQuestion + vbYesNo)
 
        If reply11 = vbYes Then

Set rsData2 = New ADODB.Recordset
Dim sql As String
Dim test, test1 As Integer
sql = "Update tblPurchaseOrderDetail set Description = '" & txtitem & "',Qty='" & txtQty & "',Unit='" & txtunit & "',UnitPrice ='" & txtUnitPrice & "', " & _
" Gross = '" & txtGross(1) & "',Discount= '" & txtDisc & " ',NetAmount='" & txtNetAmount & " ',DateCreated='" & Format(Now, "MMM-dd-yyyy") & "',CreatedBy ='" & CurrUser.UserNAME & "' " & _
" where ProductCode = '" & txtcode & "'"

Set rsData2 = CN.Execute(sql)
PuchaseOrderDetails

Discount = 0
 Discount = txtGross(1).Text * (txtDisc.Text / 100)
NetTotal = 0
Discount1 = 0
i = 0
ii = 0
For ii = 1 To ListView1.ListItems.Count
 test = ListView1.ListItems(ii).SubItems(5) * (ListView1.ListItems(ii).SubItems(6) / 100)
 test1 = test1 + test
 Next ii

 
For i = 1 To ListView1.ListItems.Count
    GrossTotal = GrossTotal + ListView1.ListItems(i).SubItems(5)
  '  Discount1 = Discount1 + ListView1.ListItems(i).SubItems(8)
   
Next i
 txtDesc.Text = Format$(test1, "#,##0.00")
 test1 = 0
  test = 0
  ii = 0
txtGross(2).Text = Format(GrossTotal, "#,##0.00")
NetTotal = toMoney(txtGross(2).Text) - toMoney(txtDesc.Text)
txtNet.Text = Format(NetTotal, "#,##0.00")


  txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
    test1 = 0
  test = 0
  ii = 0
  i = 0
  GrossTotal = 0
  NetTotal = 0
  
cmdSave_Click
 Else
 End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim ix As Integer
Dim strnames, sql1 As String
ix = 0
For ix = ListView1.ListItems.Count To 1 Step -1
If ListView1.ListItems(ix).Checked Then
strnames = ListView1.SelectedItem.SubItems(1)



                 sql1 = "Delete from tblPurchaseOrderDetail where " & _
                        " ProductCode= '" & ListView1.ListItems(ix).Text & "'"
    Set rsData = CN.Execute(sql1)
End If
Next ix


PuchaseOrderDetails
 MsgBox "Changes in record has been successfully saved.", vbInformation
End Sub
Public Sub PuchaseOrderDetails()
frmPurchaseOrder.ListView1.ListItems.Clear
Set rsData4 = New ADODB.Recordset
rsData4.Open "Select * from tblPurchaseOrderDetail where POid= '" & Text1.Text & "'", CN, adOpenStatic, adLockPessimistic
With rsData4
    While Not .EOF
                Dim LS As ListItem
                Set LS = frmPurchaseOrder.ListView1.ListItems.Add(, , rsData4!ProductCode, 1, 1)
                    LS.SubItems(1) = rsData4!Description
                    LS.SubItems(2) = rsData4!Qty
                    LS.SubItems(3) = rsData4!Unit
                    LS.SubItems(4) = Format(rsData4!UnitPrice, "###,#0.00")
                    LS.SubItems(5) = Format(rsData4!Gross, "###,#0.00")
                    LS.SubItems(6) = Format(rsData4!Discount, "###,#0.00")
                    LS.SubItems(7) = Format(rsData4!NetAmount, "###,#0.00")
 
              
                    LS.ListSubItems(1).Bold = True
                    LS.ListSubItems(5).Bold = True
                    LS.ListSubItems(7).Bold = True
            .MoveNext
         'Label1.Caption = "Selected Record: " & ListView1.SelectedItem.Index & "/" & ListView1.ListItems.Count
    Wend
    End With
End Sub

Private Sub cmdExit_Click()
Picture234.Visible = False
End Sub

Private Sub cmdPrint_Click()

Set rsData = New ADODB.Recordset
    rsData.Open "SELECT * FROM QSupplierReport WHERE POCode = '" & Me.Text1.Text & "'", CN, adOpenStatic, adLockPessimistic
    Set rptPurchaseOrder.DataSource = rsData
    rptPurchaseOrder.Sections(2).Controls("lblPOCode").Caption = rsData.Fields("POCode")
    rptPurchaseOrder.Sections(2).Controls("lblSupplier").Caption = rsData.Fields("Supplier")
    rptPurchaseOrder.Sections(2).Controls("lblAddress").Caption = rsData.Fields("Address")
    rptPurchaseOrder.Sections(2).Controls("lblContactNo").Caption = rsData.Fields("ContactNo")
    rptPurchaseOrder.Sections(2).Controls("lblContactPerson").Caption = rsData.Fields("ContactPerson")
    rptPurchaseOrder.Sections(2).Controls("lblDateRequest").Caption = rsData.Fields("DateRequisition")
    
    rptPurchaseOrder.Sections(5).Controls("lblTotalGross").Caption = Format(rsData.Fields("TotalGross"), "###,#0.00")
    rptPurchaseOrder.Sections(5).Controls("lblTotalDiscount").Caption = Format(rsData.Fields("TotalDiscount"), "###,#0.00")
    rptPurchaseOrder.Sections(5).Controls("lblTaxBase").Caption = Format(rsData.Fields("TaxBase"), "###,#0.00")
    rptPurchaseOrder.Sections(5).Controls("lblVat").Caption = Format(rsData.Fields("Vat"), "###,#0.00")
    rptPurchaseOrder.Sections(5).Controls("lblTotalNet").Caption = Format(rsData.Fields("TotalNet"), "###,#0.00")
    
    rptPurchaseOrder.Show vbModal
    
    
End Sub

Private Sub cmdSave_Click()

'Dim reply1 As Integer
' If EditFlag = True Then
'        reply1 = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
' Else
'
' End If
'        If reply1 = vbYes Then

On Error GoTo error:
If is_empty(txtsupplier, True) = True Then Exit Sub
If txtGross(2) = "" Then Exit Sub
If txtDesc.Text = "" Then Exit Sub
If txtTaxBase.Text = "" Then Exit Sub
If txtVat.Text = "" Then Exit Sub
If txtNet.Text = "" Then Exit Sub


  
    With frmPurchaseOrders.rsData
            If EditFlag = True Then .AddNew
            .Fields("POCode") = Text1.Text
            .Fields("Supplier") = txtsupplier.Text
            .Fields("ContactPerson") = txtcontactperson.Text
            .Fields("Contactno") = txtcontactno.Text
            .Fields("Address") = txtaddress.Text
            .Fields("Gross") = txtGross(2).Text
            .Fields("Discount") = txtDesc.Text
            .Fields("Taxbase") = txtTaxBase.Text
            .Fields("Vat") = txtVat.Text
            .Fields("TotalNet") = txtNet.Text
            .Fields("Remarks") = txtremark.Text
            .Fields("Active") = chkActive.Value
            .Fields("DateRequisition") = Format(dtpDate.Value, "MMM-dd-yyyy")
           .Fields("DateCreated") = Format(Now, "MMM-dd-yyyy")
            .Fields("CreatedBy") = CurrUser.UserNAME
        .Update
    End With
    '
    If EditFlag = True Then
    PurchaseDetail
    Else
    PurchaseDetailUpdate
    End If
     If EditFlag = True Then
        Call FindRec(frmPurchaseOrders.rsData, "POCode", True, Text1.Text, 0)
        MsgBox "New record has been successfully added.", vbInformation
        Dim REPLY As Integer
        REPLY = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
        If REPLY = vbYes Then
            Resetfields
        Else
            frmPurchaseOrders.LoadEntries
            Unload Me
        End If
    Else
        Call FindRec(frmPurchaseOrders.rsData, "POCode", True, Text1.Text, 0)
        MsgBox "Changes in record has been successfully saved.", vbInformation
        frmPurchaseOrders.LoadEntries
        Unload Me
    End If
    
    
    
Exit Sub
error:
    MsgBox err.Description, vbExclamation
    Unload Me
'----------------iii
'Else
'
'    End If

   


End Sub
Sub PurchaseDetail()
Dim a3, a5, a6, a7, a8 As Double
Dim a1, a2, a4 As String

i = 0

For i = 1 To ListView1.ListItems.Count
     a1 = ListView1.ListItems.Item(i).Text
     a2 = ListView1.ListItems(i).ListSubItems(1).Text
     a3 = ListView1.ListItems(i).ListSubItems(2).Text
     a4 = ListView1.ListItems(i).ListSubItems(3).Text
     a5 = ListView1.ListItems(i).ListSubItems(4).Text
     a6 = ListView1.ListItems(i).ListSubItems(5).Text
     a7 = ListView1.ListItems(i).ListSubItems(6).Text
     a8 = ListView1.ListItems(i).ListSubItems(7).Text

Set rsData = New ADODB.Recordset
rsData.Open "INSERT INTO tblPurchaseOrderDetail(ProductCode,Description,Qty,Unit,UnitPrice,Gross,Discount,NetAmount,POid,DateCreated,CreatedBy)" & _
        "values ('" & a1 & "','" & a2 & "','" & toNumber(a3) & "','" & a4 & "','" & toMoney(a5) & _
        "','" & toMoney(a6) & "','" & toMoney(a7) & " ','" & toMoney(a8) & " ','" & Text1.Text & "','" & Format(Now, "MMM-dd-yyyy") & "','" & CurrUser.UserNAME & "')", CN, adOpenStatic, adLockOptimistic
Next i

End Sub
Sub PurchaseDetailUpdate()
Dim a3, a5, a6, a7, a8 As Double
Dim a1, a2, a4 As String

iii = 0

For iii = 1 To ListView1.ListItems.Count
     a1 = ListView1.ListItems.Item(iii).Text
     a2 = ListView1.ListItems(iii).ListSubItems(1).Text
     a3 = ListView1.ListItems(iii).ListSubItems(2).Text
     a4 = ListView1.ListItems(iii).ListSubItems(3).Text
     a5 = ListView1.ListItems(iii).ListSubItems(4).Text
     a6 = ListView1.ListItems(iii).ListSubItems(5).Text
     a7 = ListView1.ListItems(iii).ListSubItems(6).Text
     a8 = ListView1.ListItems(iii).ListSubItems(7).Text

Set rsData2 = New ADODB.Recordset
Dim sql As String
sql = "Update tblPurchaseOrderDetail set Description = '" & a2 & "',Qty='" & toNumber(a3) & "',Unit='" & a4 & "',UnitPrice ='" & toMoney(a5) & "', " & _
" Gross = '" & toMoney(a6) & "',Discount= '" & toMoney(a7) & " ',NetAmount='" & toMoney(a8) & " ',DateCreated='" & Format(Now, "MMM-dd-yyyy") & "',CreatedBy ='" & CurrUser.UserNAME & "' " & _
" where ProductCode = '" & a1 & "'"

Set rsData2 = CN.Execute(sql)
Next iii
'Set rsData = CN.Execute("Select * from tblSupplier where suppliername like '" & txtsupplier.Text & "%'")
End Sub

Sub Resetfields()
ListView1.ListItems.Clear
txtitem.Text = ""
txtQty.Text = 0
txtunit.Text = ""
txtUnitPrice.Text = "0.00"
txtGross(1).Text = "0.00"
txtDisc.Text = 0
txtNetAmount.Text = "0.00"
GeneratePK
txtsupplier.Text = ""
txtcontactno.Text = ""
txtcontactperson.Text = ""
txtaddress.Text = ""
'Text2.SetFocus
frmPurchaseOrders.LoadEntries
txtTaxBase.Text = ""
txtVat.Text = ""
txtNet.Text = "0.00"
txtGross(2).Text = "0.00"
txtDesc.Text = "0.00"
End Sub
Sub GeneratePK()
Dim iCode  As Long
    iCode = getIndex("tblPurchaseOrder")
    Text1.Text = GenerateID(iCode, "PO-", "00000")
End Sub
 
Private Sub Form_Activate()
Form_Load
End Sub

Private Sub Form_Initialize()
Form_Load
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

frmPurchaseOrders.LoadEntries
frmWelcome.LOAD_MY_URL

Set rsData = Nothing
Exit Sub
err:
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub
 
Private Sub Text2_LostFocus()
ToUpper Text2
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

 
 
Private Sub ListView1_Click()




ListView1_DblClick
End Sub

Private Sub ListView1_DblClick()
  If EditFlag = False Then
  btnAdd.Caption = "Update"
On Error Resume Next
 txtcode = ListView1.SelectedItem.Text
txtitem = ListView1.SelectedItem.ListSubItems(1).Text
txtQty = ListView1.SelectedItem.ListSubItems(2).Text
txtunit = ListView1.SelectedItem.ListSubItems(3).Text
txtUnitPrice = ListView1.SelectedItem.ListSubItems(4).Text
txtGross(1) = ListView1.SelectedItem.ListSubItems(5).Text
txtDisc = ListView1.SelectedItem.ListSubItems(6).Text
txtNetAmount = ListView1.SelectedItem.ListSubItems(7).Text
Else

End If
   
End Sub

Private Sub ListView2_DblClick()
If Trim(srcRecord) = vbNullString Then
     MsgBox "Please select a record from the list .Can't proceed to the operation!", vbExclamation
Else
    txtcode = ListView2.SelectedItem.Text
    txtitem = ListView2.SelectedItem.ListSubItems(1).Text
    'txtQty = ListView2.SelectedItem.ListSubItems(2).Text
    txtUnitPrice = ListView2.SelectedItem.ListSubItems(4).Text
    txtunit = ListView2.SelectedItem.ListSubItems(5).Text
    Picture234.Visible = False
    txtQty.SetFocus
End If
End Sub

Private Sub lvButtons_H2_Click()

End Sub

Private Sub lvButtons_H1_Click()

End Sub

Private Sub Text3_Change()
If Text3.Text = "[" Or Text3.Text = "]" Or Text3.Text = "'" Then
Exit Sub
End If
tempx = 0
ListView2.ListItems.Clear
Set rsData = New ADODB.Recordset
Set rsData = CN.Execute("Select * from tblProduct where Description like '" & Text3.Text & "%'")
With rsData
    While Not .EOF
                Dim x As ListItem
                Set x = ListView2.ListItems.Add(, , rsData!ProductCode, 1, 1)
                        x.SubItems(1) = rsData!Description
                        x.SubItems(2) = rsData!Qty
                        x.SubItems(3) = Format(rsData!UnitPrice, "###,#0.00")
                        x.SubItems(4) = Format(rsData!SellingPrice, "###,#0.00")
                        x.SubItems(5) = rsData!UnitName
                   tempx = tempx + 1
           
            If !QtyRemain = 0 Then
               ListView2.ListItems(tempx).ForeColor = vbRed
               ListView2.ListItems(tempx).ListSubItems(1).ForeColor = vbRed
               ListView2.ListItems(tempx).ListSubItems(4).ForeColor = vbRed
            
            ElseIf !QtyRemain <= 5 Then
                ListView2.ListItems(tempx).ForeColor = &H4080&
                ListView2.ListItems(tempx).ListSubItems(1).ForeColor = &H4080&
                ListView2.ListItems(tempx).ListSubItems(4).ForeColor = &H4080&
 
            End If
           
        .MoveNext
     Wend
End With
End Sub

Private Sub txtDisc_Change()
 txtQty_Change
End Sub

Private Sub txtDisc_Click()
 txtQty_Change
End Sub

Private Sub txtDisc_GotFocus()
If txtDisc = "0" Then
    txtDisc.Text = Empty
End If
End Sub

Private Sub txtDisc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    txtQty.Text = ""
End If
End Sub

Private Sub txtDisc_LostFocus()
If Trim(txtDisc) = Empty Then
txtDisc.Text = "0"
End If
End Sub

Private Sub txtDisc_Validate(Cancel As Boolean)
 txtDisc.Text = toNumber(txtDisc.Text)
End Sub



Private Sub txtGross_Validate(Index As Integer, Cancel As Boolean)
 txtGross(1).Text = toMoney(toNumber(txtGross(1).Text))
End Sub

Private Sub txtitem_Click()
Call LoadEntries
If Picture234.Visible = True Then
    Picture234.Visible = False
Else
    Picture234.Visible = False
With Picture234
   .Top = 3960
   .Left = 120
   .Visible = True
End With
End If
'Text3.SetFocus

End Sub

Private Sub txtNetAmount_Validate(Cancel As Boolean)
txtNetAmount.Text = toMoney(toNumber(txtNetAmount.Text))
End Sub

Private Sub txtQty_Change()
If toNumber(txtQty.Text) < 1 Then
        btnAdd.Enabled = False
    Else
        btnAdd.Enabled = True
    End If
    
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)) - ((toNumber(txtDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text))))
End Sub

Private Sub txtQty_GotFocus()
If txtQty = "0" Then
    txtQty.Text = Empty
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    txtQty.Text = ""
End If
    
End Sub

Private Sub txtQty_LostFocus()
If Trim(txtQty) = Empty Then
txtQty.Text = "0"
End If
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtsupplier_Change()
Set rsData = CN.Execute("Select * from tblSupplier where suppliername like '" & txtsupplier.Text & "%'")
With rsData
On Error Resume Next
            txtcontactno.Text = rsData.Fields("ContactNos")
            txtcontactperson.Text = rsData.Fields("ContactPerson")
            txtaddress.Text = rsData.Fields("Address")
End With
End Sub

Private Sub txtsupplier_Click()
txtsupplier_Change
End Sub

Sub LoadEntries()
tempx = 0
ListView2.ListItems.Clear
Set rsData = New ADODB.Recordset
rsData.Open "Select * from tblProduct ORDER by productCode ASC", CN, adOpenStatic, adLockPessimistic
With rsData
    While Not .EOF
                Dim LS As ListItem
                Set LS = ListView2.ListItems.Add(, , rsData!ProductCode, 1, 1)
                LS.SubItems(1) = rsData!Description
                LS.SubItems(2) = rsData!Qty
                LS.SubItems(3) = Format(rsData!UnitPrice, "###,#0.00")
                LS.SubItems(4) = Format(rsData!SellingPrice, "###,#0.00")
                LS.SubItems(5) = rsData!UnitName
                tempx = tempx + 1
           
            If !QtyRemain = 0 Then
               ListView2.ListItems(tempx).ForeColor = vbRed
               ListView2.ListItems(tempx).ListSubItems(1).ForeColor = vbRed
               ListView2.ListItems(tempx).ListSubItems(4).ForeColor = vbRed
            ElseIf !QtyRemain <= 5 Then
                ListView2.ListItems(tempx).ForeColor = &H4080&
                ListView2.ListItems(tempx).ListSubItems(1).ForeColor = &H4080&
                ListView2.ListItems(tempx).ListSubItems(4).ForeColor = &H4080&
            End If
        .MoveNext
    Wend
    End With
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortLV ListView2
End Sub
 
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcRecCD = ListView2.SelectedItem.Index
srcRecord = ListView2.ListItems.Item(srcRecCD).Text
End Sub

Private Sub ListView2_Click()
If ListView2.ListItems.Count < 1 Then Exit Sub
End Sub


Private Sub txtUnitPrice_Validate(Cancel As Boolean)
 txtUnitPrice.Text = toMoney(toNumber(txtUnitPrice.Text))
End Sub
