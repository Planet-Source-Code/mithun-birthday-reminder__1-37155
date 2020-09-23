VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBday 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBday.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "To add a new record (F2)"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "Se&ttings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "To make settings for Birthday Reminder (F1)"
      Top             =   2595
      Width           =   1695
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "To Refresh Data (F5)"
      Top             =   3030
      Width           =   1695
   End
   Begin VB.ComboBox CmbName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1785
      TabIndex        =   13
      Top             =   210
      Width           =   2115
   End
   Begin VB.CommandButton CmdBirthdays 
      Caption         =   "Show &Birth Dates"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "To view the birthday records (F4)"
      Top             =   2595
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPBdate 
      Height          =   330
      Left            =   1785
      TabIndex        =   1
      Top             =   705
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      _Version        =   393216
      Format          =   19529729
      CurrentDate     =   36981
   End
   Begin VB.TextBox TxtPhoneNo 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1785
      TabIndex        =   3
      Top             =   1680
      Width           =   2115
   End
   Begin VB.TextBox TxtEMailId 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1785
      TabIndex        =   2
      Top             =   1185
      Width           =   2115
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   0
      Top             =   210
      Width           =   2115
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   12
      Top             =   1785
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail Id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   11
      Top             =   1260
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Birth Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   10
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   9
      Top             =   315
      Width           =   525
   End
End
Attribute VB_Name = "FrmBday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim databasestring As String
Dim ssql As String
Dim Birth_No As Integer
Dim dates(31) As Integer

Private Sub CmbName_Click()
rs.MoveFirst
Do Until rs.EOF
    If rs!bname = CmbName.Text Then
        Birth_No = rs!bno
        Fill
        GoTo check:
    Else
        rs.MoveNext
    End If
Loop
check:
If LTrim(RTrim(TxtName.Text)) = "" Then
Else
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
End If
End Sub
Private Sub CmdAbout_Click()
FrmProperties.Show vbModal
End Sub

Private Sub CmdAdd_Click()
Dim tmprs As ADODB.Recordset
Dim no As Integer
Dim RsBno As New ADODB.Recordset
Dim ans As Integer

If CmdAdd.Caption = "&Add" Then
    If LTrim(RTrim(TxtName.Text)) = "" Then
        MsgBox "Please fill in all the details", vbCritical, "Error"
        TxtName.SetFocus
        TxtName.Text = ""
    Else
        ssql = "Select max(bno) from bday"
        RsBno.CursorLocation = adUseClient
        RsBno.Open ssql, cn, adOpenStatic
        If IsNull(RsBno.Fields(0).Value) Then
            no = 1
        Else
            no = RsBno.Fields(0).Value + 1
        End If
        Set RsBno = Nothing
        ssql = "Insert into bday values (" & no & ",'" & TxtName.Text & "','" & FormatDateTime(DTPBdate.Value, vbShortDate) & "'"
        If TxtEMailId.Text = "" Then
            ssql = ssql & ",' '"
        Else
            ssql = ssql & ",'" & TxtEMailId.Text & "'"
        End If
        If TxtPhoneNo.Text = "" Then
            ssql = ssql & "," & 0
        Else
            ssql = ssql & ",'" & TxtPhoneNo & "'"
        End If
        ssql = ssql & ")"
        cn.Execute (ssql)
        MsgBox "Record Inserted Succesfully for Reminding", vbInformation, "Insertion Message"
        clear
        TxtName.SetFocus
    End If
Else
    ssql = "Select bno from bday where bname = '" & TxtName.Text & "'"
    Set tmprs = cn.Execute(ssql)
    ans = MsgBox("Are you sure you want to delete the record.", vbYesNo, "Deletion Confirmation")
    If ans = 6 Then
        ssql = "Delete from bday where bno = " & tmprs.Fields(0).Value
        cn.Execute (ssql)
    End If
    CmdRefresh_Click
End If
End Sub

Private Sub CmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CmdAdd.Caption = "&Add" Then
    CmdAdd.ToolTipText = "To add new records for reminding (F2)"
Else
    CmdAdd.ToolTipText = "To delete the record (F7)"
End If
End Sub

Private Sub CmdBirthdays_Click()
FrmBirthRecords.Show vbModal
End Sub

Private Sub CmdClose_Click()
CmdAdd.Enabled = True
FrmBday.Height = 3700
End Sub


Private Sub CmdEdit_Click()
If CmdEdit.Caption = "&Edit" Then
    ssql = "Select * from bday"
    Set rs = cn.Execute(ssql)
    If rs.EOF And rs.BOF Then
        MsgBox "No records present", vbOKOnly, "Record Navigation"
        CmbName.Visible = False
        TxtName.Visible = True
        TxtName.SetFocus
        Exit Sub
    Else
        CmbName.Visible = True
        TxtName.Visible = False
        CmdEdit.Caption = "&Save"
        CmdAdd.Caption = "&Delete"
        CmdAdd.Enabled = False
        CmdEdit.Enabled = False
        Do Until rs.EOF
            CmbName.AddItem (rs.Fields(1).Value)
            rs.MoveNext
        Loop
    End If
    CmbName.SetFocus
Else
    If LTrim(RTrim(CmbName.Text)) = "" Then
        CmdAdd.Enabled = True
        CmdEdit.Caption = "&Edit"
        CmdAdd.Caption = "&Add"
        CmbName.Visible = False
        TxtName.Visible = True
        clear
        TxtName.SetFocus
        Exit Sub
    Else
        Dim emid As String
        Dim phno As String
        Dim bnum As ADODB.Recordset
        If TxtEMailId.Text = "" Then
            emid = "None"
        Else
            emid = TxtEMailId.Text
        End If
        If TxtPhoneNo.Text = "" Then
            phno = "0"
        Else
            phno = TxtPhoneNo.Text
        End If
        ssql = "Update bday set bname = '" & TxtName.Text & "', bdate = '" & DTPBdate.Value & "', E_MailId = '" & emid & "', phone_no = " & phno & " where bno = " & Birth_No
        cn.Execute (ssql)
        ssql = ""
        MsgBox "Record updated Successfully.", vbOKOnly, "Record Updation"
        CmdRefresh_Click
        clear
        TxtName.SetFocus
    End If
End If
End Sub

Private Sub CmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CmdEdit.Caption = "&Edit" Then
    CmdEdit.ToolTipText = "To make any changes or delete the record (F3)"
Else
    CmdEdit.ToolTipText = "To save the changes in records (F6)"
End If
End Sub

Private Sub CmdRefresh_Click()
CmbName.Visible = False
CmbName.clear
TxtName.Visible = True
TxtName.Text = ""
CmdAdd.Enabled = True
CmdEdit.Enabled = True
CmdAdd.Caption = "&Add"
CmdEdit.Caption = "&Edit"
DTPBdate.Value = Now()
TxtEMailId = ""
TxtPhoneNo = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then 'F1
    CmdAbout_Click
End If
If KeyCode = 113 Then
    CmdAdd_Click 'F2
End If
If KeyCode = 114 Then
    CmdEdit_Click 'F3
End If
If KeyCode = 115 Then
    CmdBirthdays_Click 'F4
End If
If KeyCode = 116 Then
    CmdRefresh_Click 'F5
End If
If KeyCode = 117 And CmdEdit.Caption = "&Edit" Then
    MsgBox "To save the record first click on edit or press F3 to make changes", vbCritical + vbOKOnly, "Error"
End If
If KeyCode = 117 And CmdEdit.Caption = "&Save" Then
    CmdEdit_Click 'F6
End If
If KeyCode = 118 And CmdAdd.Caption = "&Add" Then
    MsgBox "To delete the record first click on Edit or press F3 to select the record.", vbCritical + vbOKOnly, "Error"
End If
If KeyCode = 118 And CmdAdd.Caption = "&Delete" Then
    CmdAdd_Click 'F7
End If
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim SettRs As New ADODB.Recordset
Dim SDate As Date
Dim c As Integer
clear
CmbName.Visible = False
FrmBday.Height = 4150

'*********** Database Connectivity *************
databasestring = "Provider=microsoft.jet.oledb.4.0;Data Source = " & App.Path & "\birthday.mdb;persist security info=False"
cn.ConnectionString = databasestring
cn.Open
ssql = "Select * from bday"
Set rs = cn.Execute(ssql)
ssql = "select caption from settings"
Set SettRs = cn.Execute(ssql)
Me.Caption = SettRs!Caption
'***********************************************

SDate = Now()
Do Until rs.EOF = True
    If Left(SDate, 5) = Left(rs.Fields(2).Value, 5) Then
        MsgBox "Today is " & rs.Fields(1).Value & "'s Birthday. E-Mail Id : " & rs.Fields(3).Value & ", Phone No. : " & rs.Fields(4).Value, vbOKOnly + vbInformation, "Birthday Reminder"
    End If
    If Left(DateAdd("H", 24, SDate), 5) = Left(rs.Fields(2).Value, 5) Then
        MsgBox "Tomorrow is " & rs.Fields(1).Value & "'s Birthday. E-Mail Id : " & rs.Fields(3).Value & ", Phone No. : " & rs.Fields(4).Value, vbOKOnly + vbCritical, "Birthday Reminder"
    End If
    If Left(DateAdd("H", 168, SDate), 5) = Left(rs.Fields(2).Value, 5) Then
        MsgBox "Its " & rs.Fields(1).Value & "'s birthday on " & FormatDateTime(DateAdd("H", 168, SDate), vbShortDate) & ". E-Mail Id : " & rs.Fields(3).Value & " Phone No. : " & rs.Fields(4).Value, vbInformation, "Birthday Reminder"
    End If
    rs.MoveNext
Loop

Dim ans As Integer
ssql = "select dat from dates where chk = True"
Set SettRs = cn.Execute(ssql)
Do Until SettRs.EOF
    If Day(SDate) = SettRs!dat Or Weekday(SDate) = 7 Or Weekday(SDate) = 1 Then
        ans = MsgBox("Do you want to Input New Birthdates?", vbYesNo, "New Birthdays")
        If ans = 7 Then
            Unload Me
            Exit Sub
        End If
        Exit Sub
    End If
SettRs.MoveNext
Loop
If Weekday(SDate) = 7 Or Weekday(SDate) = 1 Then
    ans = MsgBox("Do you want to Input New Birthdates?", vbYesNo, "New Birthdays")
    If ans = 7 Then
        Unload Me
        Exit Sub
    End If
    Exit Sub
End If
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
Set cn = Nothing
Set rs = Nothing
End Sub

Public Sub clear()
TxtName.Text = ""
CmbName.clear
DTPBdate.Value = Now()
TxtEMailId.Text = ""
TxtPhoneNo.Text = ""
End Sub

Public Sub Fill()
CmbName.Visible = False
TxtName.Visible = True
TxtName.Text = rs!bname
DTPBdate.Value = rs!bdate
If IsNull(rs!e_mailid) Then
    TxtEMailId.Text = ""
Else
    TxtEMailId.Text = rs!e_mailid
End If
If IsNull(rs!phone_no) Then
    TxtPhoneNo.Text = 0
Else
    TxtPhoneNo.Text = rs!phone_no
End If
End Sub
