VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C语言函数查询 - 计算机协会技术部"
   ClientHeight    =   5970
   ClientLeft      =   5625
   ClientTop       =   3360
   ClientWidth     =   8040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8040
   Begin VB.Frame Frame2 
      Caption         =   "函数"
      Height          =   5655
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.TextBox sl 
         Height          =   3255
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox yf 
         Height          =   1095
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox gn 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox hsm 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "示 例"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "用 法"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "功 能"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "函数名"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtDescribe 
         Height          =   270
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "回车搜索"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtFuncName 
         Height          =   270
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "功  能"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "函数名"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As New SunSoft.AdodbHelper

Dim res As ADODB.Recordset

Function CNull(ByVal sTxt As Variant) As String   'ok at 11-10-08
  If IsNull(sTxt) = True Then
    CNull = ""
  Else
    CNull = sTxt
  End If
End Function

Function ReWind(ByVal inPutX As String)
  ReWind = Replace(inPutX, "'", "''")
End Function

Private Sub Form_Load()
  If Dir(App.Path & "\clanguage.mdb") = "" Then
    Call OutputFileS(101, App.Path & "\clanguage.mdb")
  End If
  db.SetConnToFile App.Path & "\clanguage.mdb"
  Call loadtolist
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Kill App.Path & "\clanguage.mdb"
  End
End Sub

Private Sub List1_Click()
  Dim st As String
  
  If List1.ListCount = 0 Then Exit Sub
  
  st = List1.List(List1.ListIndex)
  Set res = db.ExecParamQuery("select * from functions where hsm = ?", st)

  If res.RecordCount = 0 Then
    db.ReleaseRecordset res
    Exit Sub
  End If
  
  hsm.Text = res.Fields("hsm")
  gn.Text = res.Fields("gn")
  yf.Text = res.Fields("yf")
  sl.Text = res.Fields("sl")

  db.ReleaseRecordset res
End Sub

Function loadtolist()
  Set res = db.ExecQuery("select distinct hsm from functions")
  List1.Clear
  If res.RecordCount = 0 Then
    db.ReleaseRecordset res
    Exit Function
  End If
  
  Do While Not res.EOF = True
    List1.AddItem res.Fields("hsm")
    res.MoveNext
  Loop
  
  db.ReleaseRecordset res
End Function



Private Sub txtFuncName_KeyUp(KeyCode As Integer, Shift As Integer)
  List1.Clear

  If txtFuncName.Text = "" Then
    Call loadtolist
    Exit Sub
  End If

  Set res = db.ExecParamQuery("select * from functions where hsm Like ?", ReWind(txtFuncName.Text) & "%")
  If res.RecordCount = 0 Then
    db.ReleaseRecordset res
    Call clsme
    Exit Sub
  End If
  
  Do While Not res.EOF
    List1.AddItem res.Fields("hsm")
    res.MoveNext
  Loop
  
  db.ReleaseRecordset res
  List1.ListIndex = 0
End Sub


Private Sub txtDescribe_KeyPress(KeyAscii As Integer)
  Dim pattern As String

  If txtDescribe.Text = "" Then
    If KeyAscii = 13 Then
      List1.Clear
      Call loadtolist
    End If
    Call clsme
    Exit Sub
  End If
  
  If KeyAscii = 13 Then
    pattern = "%" & ReWind(txtDescribe.Text) & "%"
    List1.Clear
    Set res = db.ExecParamQuery("select * from functions where gn Like ?", pattern)
    If res.RecordCount = 0 Then
      db.ReleaseRecordset res
      Call clsme
      Exit Sub
    End If
    
    Do While Not res.EOF
      List1.AddItem res.Fields("hsm")
      res.MoveNext
    Loop
    
    db.ReleaseRecordset res
    List1.ListIndex = 0
  End If

End Sub

Function clsme()
  hsm.Text = ""
  gn.Text = ""
  yf.Text = ""
  sl.Text = ""
End Function

Function OutputFileS(ByVal sId As Long, ByVal sFile As String)
  Dim sTemp() As Byte
  sTemp = LoadResData(sId, "CUSTOM")
  Open sFile For Binary As #1
    Put #1, , sTemp
  Close #1
End Function
