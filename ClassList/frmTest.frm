VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "clsClassList Edit Test"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQty 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   10305
      TabIndex        =   2
      Top             =   0
      Width           =   10335
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTest.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   10095
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11033
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private x As New clsList, Px As Single, Py As Single

Private Sub Form_Load()
    Set x.List = ListView1
    Set x.TextBox = txtQty
    
    x.addcolumn "Code", "code", 1200, False, True
    x.addcolumn "Description", "desc", 3000, True, False
    x.addcolumn "Qty", "qty", 800, False, True, lvwColumnRight
    x.addcolumn "Cost $", "cost", 1200, False, True, lvwColumnRight
    x.addcolumn "Sell $", "sell", 1200, False, False, lvwColumnRight
    
    x.AddItem "AB1234", "48X CDROM", "2", "10.12", "0"
    x.AddItem "AB1234", "80gb 7200RPM HDD", "2", "10.12", "0"
    
    x.Resize
    
    CalculateTotals
End Sub

Sub CalculateTotals()
    Dim Row As ListItem
    
    For Each Row In ListView1.ListItems
        Row.SubItems(4) = Format(Val(Row.SubItems(2)) * Val(Row.SubItems(3)), "#0.00")
    Next
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    CalculateTotals
End Sub
