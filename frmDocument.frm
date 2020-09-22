VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomer 
   Appearance      =   0  'Flat
   Caption         =   "Customer Information"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9060
      TabIndex        =   39
      Top             =   0
      Width           =   9060
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   45
      ScaleHeight     =   8055
      ScaleWidth      =   8985
      TabIndex        =   0
      Top             =   660
      Width           =   8985
      Begin VB.Frame Frame2 
         Caption         =   "Customer Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00842520&
         Height          =   3255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   5775
         Begin VB.ComboBox varContactTitle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3480
            TabIndex        =   13
            Top             =   2640
            Width           =   2055
         End
         Begin VB.TextBox varClientName 
            Height          =   375
            Left            =   120
            MaxLength       =   60
            TabIndex        =   3
            Top             =   480
            Width           =   5385
         End
         Begin VB.TextBox varAddress 
            Height          =   1080
            Left            =   120
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1920
            Width           =   3255
         End
         Begin VB.TextBox varContactPerson 
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox varCell 
            Height          =   375
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1920
            Width           =   2055
         End
         Begin VB.TextBox varPhone 
            Height          =   375
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*Client Name/Company Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Title:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3480
            TabIndex        =   12
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person "
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   3480
            TabIndex        =   5
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cell Phone"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   3480
            TabIndex        =   9
            Top             =   1680
            Width           =   855
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Credits and Limits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00842520&
         Height          =   3255
         Left            =   5820
         TabIndex        =   14
         Top             =   120
         Width           =   3015
         Begin VB.TextBox varRemarks 
            BackColor       =   &H00FFFFFF&
            Height          =   1950
            Left            =   105
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1200
            Width           =   2835
         End
         Begin VB.TextBox varNumeric 
            Height          =   375
            Index           =   0
            Left            =   75
            MaxLength       =   6
            TabIndex        =   18
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox varNumeric 
            Height          =   375
            Index           =   1
            Left            =   1065
            MaxLength       =   6
            TabIndex        =   19
            Text            =   "0"
            Top             =   495
            Width           =   930
         End
         Begin VB.TextBox varNumeric 
            Height          =   375
            Index           =   2
            Left            =   2025
            MaxLength       =   6
            TabIndex        =   20
            Text            =   "0"
            Top             =   495
            Width           =   930
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   915
            Width           =   855
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Days:"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   9
            Left            =   1065
            TabIndex        =   17
            Top             =   255
            Width           =   885
         End
         Begin VB.Label lbl_crlimit 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Limit:"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   15
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Discounts:"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   3
            Left            =   2025
            TabIndex        =   16
            Top             =   255
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Default         =   -1  'True
         Height          =   375
         Left            =   7845
         TabIndex        =   24
         Top             =   3420
         Width           =   855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6885
         TabIndex        =   23
         Top             =   3420
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Register Vehicles/Autos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00842520&
         Height          =   4035
         Left            =   0
         TabIndex        =   25
         Top             =   3840
         Width           =   8865
         Begin VB.TextBox varVRemarks 
            Height          =   375
            Left            =   120
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1080
            Width           =   7500
         End
         Begin VB.ComboBox varMake 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1800
            TabIndex        =   31
            Top             =   480
            Width           =   1680
         End
         Begin VB.CommandButton cmdRemoveList 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8040
            Picture         =   "frmDocument.frx":01CA
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Click To Register Vechicles and Autos of Client"
            Top             =   1080
            Width           =   330
         End
         Begin VB.TextBox varPlateNum 
            Height          =   375
            Left            =   120
            MaxLength       =   10
            TabIndex        =   30
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox varColor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6120
            TabIndex        =   33
            Top             =   480
            Width           =   1785
         End
         Begin VB.ComboBox varModel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3720
            TabIndex        =   32
            Top             =   480
            Width           =   2040
         End
         Begin VB.CommandButton cmdAddToList 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7680
            Picture         =   "frmDocument.frx":0394
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Click To Register Vechicles and Autos of Client"
            Top             =   1080
            Width           =   330
         End
         Begin MSComctlLib.ListView lvVechicles 
            Height          =   2370
            Left            =   120
            TabIndex        =   38
            Top             =   1560
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   4180
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Plate #"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Make"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Color"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Remarks /Memos"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   12
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   11
            Left            =   3600
            TabIndex        =   29
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Make"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   2040
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Plate No:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lbl_cust 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Color:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   6720
            TabIndex        =   28
            Top             =   240
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub
