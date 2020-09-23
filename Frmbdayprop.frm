VERSION 5.00
Begin VB.Form FrmProperties 
   Caption         =   "Properties"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmbdayprop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1050
      TabIndex        =   5
      Top             =   5925
      Width           =   1275
   End
   Begin VB.TextBox TxtCaption 
      Height          =   330
      Left            =   3885
      TabIndex        =   4
      ToolTipText     =   "Enter the Caption to appear on the Main form."
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4725
      TabIndex        =   2
      Top             =   5925
      Width           =   1275
   End
   Begin VB.ListBox LstNoDays 
      Height          =   1200
      ItemData        =   "Frmbdayprop.frx":030A
      Left            =   5565
      List            =   "Frmbdayprop.frx":036B
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      ToolTipText     =   "Just select the dates whenever you want Birthday Reminder should allow you to enter Records."
      Top             =   105
      Width           =   645
   End
   Begin VB.Label Label9 
      Caption         =   "To Delete any Record: F7"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   13
      Top             =   5280
      Width           =   5370
   End
   Begin VB.Label Label8 
      Caption         =   "To Save the Changes made to Birthday Records: F6"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   12
      Top             =   4835
      Width           =   5370
   End
   Begin VB.Label Label5 
      Caption         =   "To Refresh Birthday Reminder: F5"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   11
      Top             =   4393
      Width           =   5370
   End
   Begin VB.Label Label7 
      Caption         =   "To make the Settings for Birthday Reminder: F1"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   10
      Top             =   2625
      Width           =   5265
   End
   Begin VB.Label Label6 
      Caption         =   "To See the Birthday Records: F4"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   9
      Top             =   3951
      Width           =   5265
   End
   Begin VB.Label Label4 
      Caption         =   "To Add Birthday Records for Reminding: F2"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   8
      Top             =   3067
      Width           =   5370
   End
   Begin VB.Label Label3 
      Caption         =   "To make Changes to Birthday Records: F3"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   735
      TabIndex        =   7
      Top             =   3509
      Width           =   5265
   End
   Begin VB.Label Label2 
      Caption         =   "Keyboard Info"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2310
      TabIndex        =   6
      Top             =   2205
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Caption to appear for your Birthday Reminder like Mithuns Birthday Reminder"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   105
      TabIndex        =   3
      Top             =   1365
      Width           =   3585
   End
   Begin VB.Label LblNoDays 
      Caption         =   $"Frmbdayprop.frx":03E2
      ForeColor       =   &H000000FF&
      Height          =   1170
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5160
   End
End
Attribute VB_Name = "FrmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
Dim SettRs As New ADODB.Recordset
Dim ssql As String
Dim databasestring As String
Dim days(31) As Integer
Dim i As Integer

Private Sub CmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdOK_Click()
Dim rs As New ADODB.Recordset
Dim c As Integer
For c = 1 To i
    ssql = "update dates set chk = true where dat = " & days(c)
    cn.Execute (ssql)
Next c
    
ssql = "Delete from settings"
cn.Execute (ssql)
    
ssql = "Insert into settings values ('" & TxtCaption.Text & "')"
cn.Execute (ssql)
FrmBday.Caption = TxtCaption.Text
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOK_Click
End If
End Sub

Private Sub Form_Load()
databasestring = "Provider=microsoft.jet.oledb.4.0;Data Source = " & App.Path & "\birthday.mdb;persist security info=False"
cn.ConnectionString = databasestring
cn.Open
ssql = "Select caption from settings"
Set SettRs = cn.Execute(ssql)
TxtCaption.Text = SettRs!Caption
i = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
Set cn = Nothing
Set SettRs = Nothing
End Sub

Private Sub LstNoDays_Click()
i = i + 1
days(i) = LstNoDays.Text
ssql = "Update dates set chk = false"
cn.Execute (ssql)
End Sub

Private Sub LstNoDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub TxtCaption_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
    Beep
End If
End Sub

