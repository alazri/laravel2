VERSION 5.00
Begin VB.Form frmClock 
   Caption         =   "CLOCK DONE BY: HAMOOD ALAZRI"
   ClientHeight    =   6048
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5976
   LinkTopic       =   "Form2"
   ScaleHeight     =   6048
   ScaleWidth      =   5976
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "flash form"
      Height          =   372
      Left            =   2640
      TabIndex        =   15
      Top             =   5640
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   1920
   End
   Begin VB.Timer tmrQuartz 
      Interval        =   1000
      Left            =   3600
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MINIMIZE"
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Top             =   5640
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MAXIMIZE"
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   1212
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4680
      TabIndex        =   14
      Top             =   2400
      Width           =   732
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2400
      TabIndex        =   13
      Top             =   4680
      Width           =   732
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   732
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   3360
      TabIndex        =   11
      Top             =   360
      Width           =   732
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4440
      TabIndex        =   10
      Top             =   1080
      Width           =   732
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4560
      TabIndex        =   9
      Top             =   3600
      Width           =   732
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   3600
      TabIndex        =   8
      Top             =   4560
      Width           =   732
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   960
      TabIndex        =   7
      Top             =   4440
      Width           =   732
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   732
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   20
      DrawMode        =   6  'Mask Pen Not
      Height          =   8172
      Left            =   840
      Shape           =   3  'Circle
      Top             =   -1320
      Width           =   3852
   End
   Begin VB.Line LineSecond 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   5
      X1              =   2640
      X2              =   1920
      Y1              =   2760
      Y2              =   3720
   End
   Begin VB.Line LineMinute 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   2640
      X2              =   3000
      Y1              =   2760
      Y2              =   3960
   End
   Begin VB.Line LineHour 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      X1              =   2640
      X2              =   2760
      Y1              =   2760
      Y2              =   1920
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4320
      TabIndex        =   2
      Top             =   5640
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   372
      Left            =   4680
      Top             =   5640
      Width           =   852
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0FF&
      Height          =   5532
      Left            =   0
      Top             =   0
      Width           =   5532
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.14159
Option Explicit
  

   Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
   Private mb_Flashing As Boolean
Private Sub Command1_Click()
frmClock.Width = 7670
End Sub
Private Sub Command2_Click()
frmClock.Width = 5500
End Sub

Private Sub Command3_Click()
   
   
   
         
         mb_Flashing = Not mb_Flashing
         Timer1.Enabled = mb_Flashing
         
         If mb_Flashing = False Then
             Call FlashWindow(Me.hwnd, 0)
         End If
   
   End Sub
   
Private Sub Form_Load()
Call tmrQuartz_Timer
Dim vWindowPos As Long
 vWindowPos = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 1 Or 2)

End Sub
Private Sub Timer1_Timer()
Call FlashWindow(Me.hwnd, 1)
If Image1.Visible = True Then
Image1.Visible = False
Else
Image1.Visible = True
End If
End Sub
Private Sub tmrQuartz_Timer()
Dim Hours As Single, Minutes As Single, Seconds As Single
Dim TrueHours As Single
lblTime.Caption = Time
'Beep
Hours = Hour(Time)
Minutes = Minute(Time)
Seconds = Second(Time)
TrueHours = Hours + Minutes / 60
' I made all the X1 and Y1 equal in the form
LineHour.X2 = 1000 * Cos(PI / 180 * (30 * TrueHours - 90)) + LineHour.X1
LineHour.Y2 = 1000 * Sin(PI / 180 * (30 * TrueHours - 90)) + LineHour.Y1
LineMinute.X2 = 1500 * Cos(PI / 180 * (6 * Minutes - 90)) + LineHour.X1
LineMinute.Y2 = 1500 * Sin(PI / 180 * (6 * Minutes - 90)) + LineHour.Y1
LineSecond.X2 = 1600 * Cos(PI / 180 * (6 * Seconds - 90)) + LineHour.X1
LineSecond.Y2 = 1600 * Sin(PI / 180 * (6 * Seconds - 90)) + LineHour.Y1
End Sub


 Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

