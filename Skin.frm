VERSION 5.00
Begin VB.Form Skin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "GlassBall"
   ClientHeight    =   7635
   ClientLeft      =   1755
   ClientTop       =   1350
   ClientWidth     =   10305
   Icon            =   "Skin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   687
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   435
      Picture         =   "Skin.frx":DDB5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   9705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Aero Vista For VB6 Good"
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Aero Vista For VB6 Good"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6375
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   375
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Aero Vista For VB6 Good"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   345
      Width           =   2295
   End
End
Attribute VB_Name = "Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY As Long = &H1
Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Dim FileRes As Integer
Dim Buffer() As Byte
Sub Center(FormName As Form)
Move (Screen.Width - FormName.Width) \ 2, (Screen.Height - FormName.Height) \ 2
End Sub
Sub MoveForm(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub
Private Sub Form_Load()
On Error GoTo Down:
    Dim tmpSt As Long
    Open TheSystemDir() & "\Vista.png" For Input As 1#
    Close 1#
    GoTo Nex:
Down:
    Buffer = LoadResData("Vista", "DOWNLOAD")
    FileRes = FreeFile
    Open TheSystemDir() & "\Vista.png" For Binary Access Write As #FileRes
    Put #FileRes, , Buffer
    Close #FileRes
Nex:
    Call Center(Skin)
    tmpSt = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    tmpSt = tmpSt Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, tmpSt
    SetLayeredWindowAttributes Me.hWnd, RGB(255, 0, 0), 0, LWA_COLORKEY
    Load Aero
    Aero.Show
End Sub
Private Sub Image1_Click()
Unload Aero
End Sub
