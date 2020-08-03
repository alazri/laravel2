VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4872
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4884
   LinkTopic       =   "Form1"
   ScaleHeight     =   4872
   ScaleWidth      =   4884
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   3612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   492
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Conn As New Connection
Dim Rs As New Recordset
Dim strSQL As String
strSQL = "Select id,user_name From login"
Conn.Open "alhyah_trading2"
Rs.Open strSQL, Conn


If Rs.EOF = True Then
 MsgBox ("No fields ")
End If
If Rs.EOF = False Then
MsgBox ("Yes")
End If

Do Until Rs.EOF
List1.AddItem Rs.Fields("id")
List1.AddItem Rs.Fields("user_name")

Rs.MoveNext
Loop




End Sub
