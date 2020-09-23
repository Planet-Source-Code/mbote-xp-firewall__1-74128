VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   840
   ClientTop       =   2775
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   10455
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8040
      Top             =   4080
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8760
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Source Ip"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Source Mask"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Src. Port"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dest. Ip"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dest. Mask"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dest. Port"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Protocol"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "InOut"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Action"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9600
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":015A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02B4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0706
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B58
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FAA
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1724
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   953
      ButtonWidth     =   1429
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Allow All"
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Block All"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Block Ping"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Rule"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Del Rule"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Dim Executing As Boolean 'one instance app. only
Private Sub Form_Load()
  On Error GoTo Form1Load_Error

  Executing = False
  If App.PrevInstance Then
    Executing = True
    MsgBox "Ya está en ejecución", 16
    Unload Me
    End
  End If

  LoadRulesFirewall
  Toolbar1_ButtonClick Toolbar1.Buttons(1)  'activate the firewall
  
  Exit Sub
Form1Load_Error:
  MsgBox "MbInfo: Error en Form1_Load: " & Err.Description, vbCritical
        
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo Form1QUnload_Error

  If Executing = True Then 'one instance app. only
    Unload Me
    Exit Sub
  End If
  
  Winsock1.Close
  Unload Me
  
  Exit Sub
Form1QUnload_Error:
  MsgBox "MbInfo: Error en Form1Query_Unload: " & Err.Description, vbCritical
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo Form1Unload_Error

  Set Form1 = Nothing

  Exit Sub
Form1Unload_Error:
  MsgBox "MbInfo: Error en Form1_Unload: " & Err.Description, vbCritical

End Sub
Private Sub Form_Resize()
  On Error GoTo Form1Resize_Error

  If Me.WindowState = vbMinimized Then Exit Sub
  If Me.Width < 5000 Then Me.Width = 5000
  If Me.Height < 5000 Then Me.Height = 5000
  
  ListView1.Width = Form1.Width - 100
  ListView1.Height = Form1.Height - 1000
    
  Exit Sub
Form1Resize_Error:
  MsgBox "MbInfo: Error en Form1_Resize: " & Err.Description, vbCritical

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button) 'firewall - 10 Buttons
  On Error GoTo Toolbar1_ButtonClick_Error

  Select Case Button.Index
    Case 1 'allow all
      Toolbar1.Buttons(3).Enabled = False
      StartFilters
    Case 2 'stop
      Toolbar1.Buttons(1).Enabled = True
      Toolbar1.Buttons(3).Enabled = True
      StopFilters
    Case 3 'block all
      Toolbar1.Buttons(1).Enabled = False
      BlockAll
    Case 4 'sep1
      '
    Case 5 'block ping
      'not implemented
    Case 6 'sep2
      '
    Case 7 'add rule
      'not implemented - editing the file
    Case 8 'del rule
      'not implemented - editing the file
    Case 9 'sep3
      '
    Case 10 'exit
      Unload Me
  End Select
  
  Exit Sub
Toolbar1_ButtonClick_Error:
  MsgBox "MbInfo: Error en Toolbar1_ButtonClick: " & Err.Description, vbCritical

End Sub
Private Sub Timer1_Timer() '1000 ms.
  On Error GoTo Timer1_Error

  Timer1.Enabled = False
  
  'firewall reload
  If FileDateTime(App.Path & "\FirewallRules.txt") <> RulesFirewallDate Then
    LoadRulesFirewall
    ResetFilters
  End If
  
  DoEvents
  Timer1.Enabled = True
  
  Exit Sub
Timer1_Error:
  MsgBox "MbInfo: Error en Timer1: " & Err.Description, vbCritical

End Sub
