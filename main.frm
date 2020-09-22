VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formMain 
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11715
      TabIndex        =   8
      Top             =   0
      Width           =   11715
      Begin VB.Label lbl_view 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Probably  Tool Bar will go Here .......  Right Click On The ICONS of WebBrowser Try Different Links (Start with POS)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   240
         MouseIcon       =   "main.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   150
         Width           =   8235
      End
   End
   Begin VB.PictureBox pageTaskPanel 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00B16303&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7965
      Left            =   0
      ScaleHeight     =   7965
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   615
      Width           =   2895
      Begin VB.PictureBox panelTask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   45
         ScaleHeight     =   6345
         ScaleWidth      =   2745
         TabIndex        =   1
         Top             =   75
         Width           =   2775
         Begin VB.TextBox txtSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            MaxLength       =   35
            TabIndex        =   3
            Top             =   735
            Width           =   2175
         End
         Begin VB.ComboBox cmbSearchList 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "main.frx":044A
            Left            =   45
            List            =   "main.frx":0469
            Style           =   2  'Dropdown List
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   315
            Width           =   2610
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   270
            TabIndex        =   11
            Top             =   2040
            Width           =   1800
         End
         Begin VB.Label lbl_view 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   " OK"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Index           =   3
            Left            =   2235
            MouseIcon       =   "main.frx":04E2
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   750
            Width           =   450
         End
         Begin VB.Image img01 
            Height          =   240
            Index           =   1
            Left            =   120
            Picture         =   "main.frx":092C
            Top             =   1470
            Width           =   240
         End
         Begin VB.Label lbl_view 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "      Display ListView"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   420
            MouseIcon       =   "main.frx":0AF6
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1500
            Width           =   1860
         End
         Begin VB.Image img01 
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "main.frx":0F40
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label lbl_view 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "      No Take Me home"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   405
            MouseIcon       =   "main.frx":110A
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1200
            Width           =   1875
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wBro 
      Height          =   3375
      Left            =   5400
      TabIndex        =   4
      Top             =   615
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   5953
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
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
   Begin MSComctlLib.ListView lvGrid 
      Height          =   3975
      Left            =   2940
      TabIndex        =   5
      Top             =   615
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "My ColumnHeader1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "My ColumnHeader2"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu m1 
      Caption         =   "Some Menus"
      Begin VB.Menu m12 
         Caption         =   "{Menu1}"
      End
      Begin VB.Menu m13 
         Caption         =   "{Menu1}"
      End
   End
   Begin VB.Menu m2 
      Caption         =   "MENULINE"
      Visible         =   0   'False
      Begin VB.Menu m21 
         Caption         =   "{}"
      End
      Begin VB.Menu m22 
         Caption         =   "{}"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu m23 
         Caption         =   "{}"
      End
      Begin VB.Menu m24 
         Caption         =   "{}"
      End
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents doc As HTMLDocument
Attribute doc.VB_VarHelpID = -1
Dim mWindow As HTMLWindow2
Public MyWindow As GridType

Private Function doc_onclick() As Boolean
    Dim strTag As String, strType As String
    Dim objSrc As HTMLInputButtonElement
 Set objSrc = mWindow.event.srcElement
    If objSrc.tagName = "A" Then
          strTag = objSrc.Id
         Debug.Print strTag
         
          Select Case strTag
          Case "idx_addcustomer"
                 frmProduct.Show
          Case "idx_addproduct"
                frmProduct.Show
          Case "idx_addvendor"
        
         Case "idx_payslip"
            MyWindow = ListView
         Form_Resize
                
        Case "idx_payslip", "idx_mysetting", "idx_salesregistry"
                        frmProduct.Show
        Case "idx_pos"
                frmProduct.Show vbModal
        Case Else
            MsgBox "Try Click other Icons Too"
          End Select
           
        doc_onclick = True
        
    End If
    
     
End Function


Private Function doc_oncontextmenu() As Boolean
Dim strTag As String, strType As String
Dim mcap1 As String
Dim mcap2 As String
Dim mcap3 As String
doc_oncontextmenu = False
   
 Set objSrc = mWindow.event.srcElement
    If objSrc.tagName = "A" Then
          strTag = objSrc.Id
         
         Select Case strTag
        
        Case "idx_pos"
            mcap1 = "&Goto My Pos": mcap2 = "": mcap3 = "And So on"
         Case "idx_receipts"
             mcap1 = "Add New Products": mcap2 = "Review Product History": mcap3 = "And So on and on and on"
         Case "idx_salesregistry"
             mcap1 = "Add New Vendor": mcap2 = "Review Customers History": mcap3 = "And So on"
         Case "idx_payslip"
             mcap1 = "Generate Paysilp": mcap2 = "Search Employee Bla bla": mcap3 = "And So on"
         Case "idx_pettycash"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
         Case "idx_payslip"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_payments"
            mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_purchasereturn"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_purchaseregistry"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_addcustomer"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_expensevoucher"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_addproduct"
             mcap1 = "Add New Customers": mcap2 = "Review Customers History": mcap3 = "And So on"
        Case "idx_addvendor"
             mcap1 = "Add New Customers": mcap2 = "add ventors ": mcap3 = "And So on"
        Case "idx_addemployee"
             mcap1 = "Add New Employee": mcap2 = "...": mcap3 = "And So on"
        Case "idx_mysetting"
             mcap1 = "lOGG oFF": mcap2 = "SHOW OFF": mcap3 = "And So on"
        Case Else
            Exit Function
        End Select
                

                
           m21.Caption = mcap1
           m22.Caption = mcap2
           m23.Caption = mcap3
           m24.Caption = " Need Help Got to VbAccelator.com "
           PopupMenu m2
    End If
     
    
End Function




Private Sub Form_Load()
lbl.Caption = "Few of the Images are from MAMBO.. and others are from Some Open Source Softwares"
'This is very classic however i saw plenty of screen welcome screen here and there in Planet source
' i though it might be helpful

'Credits:VBacclerators

'you can see how to add resources on or view
'Credits:::MINDS /Cutting Edge Classic'
'Dino Esposito'
'same topic on Res Protocol
'instead of making a different res files I have used same exe which
'You can put xsl,js css etc in there
'so webbrowser can act like grid too.
'that i will later however
'here is the link for res://protocol in VBaccelrator
'TOPIC::::Storing and Showing HTML Resources in a VB Application
'LINK::::;http://www.vbaccelerator.com/home/vb/code/libraries/Resources/Storing_HTML_Resources_in_VB_Applications/article.asp
'try downloading and check the exe you have made '

'another helpful link Dino Esposito :http://www.microsoft.com/mind/0199/cutting/cutting0199.asp



     If MyWindow = GridType.HomeList Then
         wBro.Navigate ("res://" & App.Path & "\" & App.EXEName & ".exe/home.html")
    Else
        lvGrid.Visible = True
        wBro.Visible = False
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If MyWindow = GridType.HomeList Then
        
    wBro.Left = pageTaskPanel.ScaleWidth
    wBro.Width = Me.ScaleWidth - pageTaskPanel.ScaleWidth
    wBro.Height = Me.ScaleHeight
    lvGrid.Visible = False
    wBro.Visible = True
    
Else
    lvGrid.Visible = True
    wBro.Visible = False
    lvGrid.Left = pageTaskPanel.Width
    
    lvGrid.Width = Me.ScaleWidth - pageTaskPanel.ScaleWidth
    lvGrid.Height = Me.ScaleHeight
End If
End Sub


Private Sub lbl_view_Click(Index As Integer)
    If Index = 1 Then
        loadmydummydata
         MyWindow = ListView
         
         Form_Resize
         
    Else
        MyWindow = HomeList
         Form_Resize
    End If
End Sub


Private Sub lbl_view_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbl_view(Index).BackColor = &H75CDEE
End Sub



Private Sub panelTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For I = 0 To lbl_view.UBound
    lbl_view(I).BackColor = vbButtonFace
Next

End Sub

Private Sub wBro_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error GoTo adder:
     Set doc = wBro.Document
     Set mWindow = doc.parentWindow
Exit Sub
adder:
    wBro.Navigate ("About:blank")
    wBro.Document.Write ("Sorry For This ....YOU GOT TO MAKE EXE FIRST... Then You Will See Some Sort of Display HERE.. PSC Doesn't Allow To upload Exe Thanks")
End Sub




Sub loadmydummydata()
Dim lst As ListItem
lvGrid.ListItems.Clear

lvGrid.ColumnHeaders(1).Width = lvGrid.Width \ 2
lvGrid.ColumnHeaders(2).Width = lvGrid.Width \ 2

For I = 1 To 200
Set lst = lvGrid.ListItems.Add(, , "Data--Line" & I)
    lst.SubItems(1) = "You Cost Me Php." & I
    

Next
End Sub
