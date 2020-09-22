VERSION 5.00
Begin VB.Form frmProduct 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Products"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6480
      TabIndex        =   18
      Top             =   0
      Width           =   6480
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "0"
      Top             =   3720
      Width           =   1020
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   3090
      Width           =   6045
   End
   Begin VB.TextBox txtSerialnum 
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
      Left            =   120
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2385
      Width           =   6000
   End
   Begin VB.TextBox txtProdname 
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
      Left            =   120
      MaxLength       =   155
      TabIndex        =   1
      Top             =   825
      Width           =   6000
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1410
      Width           =   6000
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "0"
      Top             =   3720
      Width           =   1020
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "0"
      Top             =   3720
      Width           =   1020
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "0"
      Top             =   3720
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Selling Price"
      Height          =   210
      Left            =   1320
      TabIndex        =   17
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Category If Any.."
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2835
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "*Product Name"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   585
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number ** if any"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2175
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   1215
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Quantity In Stock"
      Height          =   210
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Reorder Level"
      Height          =   210
      Left            =   3840
      TabIndex        =   15
      Top             =   3480
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "*Unit Price"
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   3495
      Width           =   735
   End
End
Attribute VB_Name = "frmProduct"
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

