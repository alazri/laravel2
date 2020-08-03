VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3948
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4068
   LinkTopic       =   "Form1"
   ScaleHeight     =   3948
   ScaleWidth      =   4068
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "form2"
      Height          =   252
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "close cd room"
      Height          =   492
      Left            =   2640
      TabIndex        =   4
      Top             =   2880
      Width           =   1332
   End
   Begin VB.CommandButton OpenCDDriveDoor1 
      Caption         =   "open cd room"
      Height          =   492
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1932
   End
   Begin VB.CommandButton change_row_background 
      Caption         =   "change row 2 back ground"
      Height          =   372
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   2052
   End
   Begin VB.CommandButton fill_Table 
      Caption         =   "fill table"
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Width           =   1092
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "dbusage.frx":0000
      Height          =   1692
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3132
      _ExtentX        =   5525
      _ExtentY        =   2985
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long




Private Sub change_row_background_Click()
Dim x As Integer
x = Val(InputBox("«œŒ· —ﬁ„« Ì —«ÊÕ »Ì‰ 0 Ê  " & MSFlexGrid1.Rows - 1, " ÕœÌœ »«··Ê‰ «·√Õ„—"))
If x < 0 Or x > (MSFlexGrid1.Rows - 1) Then
    MsgBox "ﬁÌ„… ’› Œ«ÿ∆ !!!"
    Exit Sub
End If

MSFlexGrid1.Row = x
For i = 0 To 1
    MSFlexGrid1.Col = i
    MSFlexGrid1.CellBackColor = QBColor(12)
Next i

End Sub

Private Sub D1_Validate(Action As Integer, Save As Integer)
'√Ê· ŒÿÊ… ÂÌ ⁄„· « ’«· »Ì‰ ﬁ«⁄œ… «·»Ì«‰«  «·„ƒﬁ … T2 Ê«·ÃœÊ· :

'Set T2 = D1.OpenRecordset("login", dbOpenTable)
End Sub



Private Sub Command1_Click()
OpenCDDriveDoor (False)

End Sub

Private Sub Command2_Click()
frmClock.Show

End Sub

Private Sub fill_Table_Click()

'”Ê› ‰ﬁÊ„ «·¬‰ »⁄„· «Ã—«¡ »«”„ fill_Table „Â„ Â —”„ «·ÃœÊ· ÊÊ÷⁄ »Ì«‰«  T2 ›Ì œ«Œ·Â .
'›Ì Â–« «·«Ã—«¡ Ê„Â„ ‰« «·√Ê·Ï ﬁ»· «œŒ«· «·»Ì«‰«  ÂÌ  ’„Ì„ «·‘ﬂ· «·⁄«„ ··ÃœÊ· ° ›≈–« «—œ‰« «·ﬂ «»… ›Ì „—»⁄ 0 Ê 0 „À·« ° ›≈‰‰« ‰ﬁÊ„ »«· «·Ì :
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0

MSFlexGrid1.Text = "«·ﬁÌ„…"
'›Ì »œ«Ì…  ‘ﬂÌ· «·ÃœÊ· ° ›»⁄œ „”Õ «·ÃœÊ· »«·ﬂ«„· ° ”Ê› ‰Ã⁄· ⁄œœ «·’›Ê› = 1 ÊÂÊ ’› «·⁄‰Ê«‰ ›ﬁÿ ° Ê⁄œœ «·√⁄„œ… 3 ( «·—ﬁ„ - «·«”„ - «·Â« › ) :
MSFlexGrid1.Clear
MSFlexGrid1.Cols = 3
MSFlexGrid1.Rows = 1
'‰‰ ﬁ· »Ì‰ «·√⁄„œ… ›Ì «·’› «·√Ê· ·Ê÷⁄ —ƒÊ” «·Ãœ«Ê· :
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.Text = "«·—ﬁ„"

MSFlexGrid1.Col = 1
MSFlexGrid1.Text = "«·«”„"

MSFlexGrid1.Col = 2
MSFlexGrid1.Text = "«·Â« ›"


'«·ŒÿÊ… «·√Ê·Ï ÂÌ „⁄—›… ⁄œœ «·”Ã·«  ° ÊÂÊ ⁄œœ «·’›Ê› ›Ì «·ÃœÊ· + 1 ( „‰ √Ã· ’› «·⁄‰Ê«‰ ) .
If T2.RecordCount < 1 Then Exit Sub

    T2.MoveLast
    T2.MoveFirst
    N = T2.RecordCount
    
   MSFlexGrid1.Rows = N + 1
' Ê«·¬‰ ”Ê› ‰ﬁÊ„ »«·œÊ—«‰ ⁄·Ï ﬂ· ”Ã·«  ﬁ«⁄œ… «·»Ì«‰«  ° Ê„⁄ ﬂ· ”Ã· ”Ê› ‰‰ﬁ· ≈·Ï «·’› «·Œ«’ »Â ° Ê„‰ À„ ‰ ‰ﬁ· »Ì‰ «·√⁄„œ… «·À·«À… ·Ê÷⁄ ﬂ· »Ì«‰ ›Ì „Ê÷⁄Â «·’ÕÌÕ ° Ê√ŒÌ—« ‰‰ ﬁ· ≈·Ï «·”Ã· «· «·Ì .

For i = 1 To N
        MSFlexGrid1.Row = i
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = T2!nu

        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = T2!Fn

        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = T2!Te

        T2.MoveNext

        Next i
'√ŒÌ—« ”Ê› ‰ﬁÊ„ » ⁄œÌ· ⁄—÷ «·⁄„ÊœÌ‰ «·√Ê· Ê«·À«‰Ì ·Ì” ÿÌ⁄« «” Ì⁄«» „“Ìœ „‰ «·»Ì«‰«  :
MSFlexGrid1.ColWidth(0) = 500
MSFlexGrid1.ColWidth(1) = 1500
End Sub



Private Sub Form_Load()
'Set T2 = D1.OpenRecordset("login", dbOpenTable)
Dim Conn As New Connection
Dim Rs As New Recordset
Dim strSQL As String
strSQL = "Select id,user_name From login"
Conn.Open "alhyah_trading2"
Rs.Open strSQL, Conn

frmClock.Show
Me.Hide




End Sub

Private Sub MSFlexGrid1_DblClick()

    MsgBox MSFlexGrid1.Text

End Sub
Private Sub OpenCDDriveDoor1_Click()
OpenCDDriveDoor (True)
End Sub



Public Sub OpenCDDriveDoor(ByVal State As Boolean)
If State = True Then
Call mciSendString("Set CDAudio Door Open", 0&, 0&, 0&)
Else
Call mciSendString("Set CDAudio Door Closed", 0&, 0&, 0&)
End If
End Sub
