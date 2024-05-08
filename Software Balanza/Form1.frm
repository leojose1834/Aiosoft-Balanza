VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{C3967F87-FD47-4E87-B007-06264CBD1A36}#2.0#0"; "systray.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aiosoft Balanza."
   ClientHeight    =   4515
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   2445
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbbpuerto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   810
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cerrar programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ocultar ventana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin IconSystray.sysTray sysTray1 
      Left            =   1800
      Top             =   4680
      _ExtentX        =   529
      _ExtentY        =   529
      IconPicture     =   "Form1.frx":058A
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detener"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Timer Timer3 
      Interval        =   5
      Left            =   720
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   4680
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   4680
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1200
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reiniciar Balanza"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      MaxLength       =   8
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nro Puerto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Const GW_HWNDNEXT = 2
Const INPUT_KEYBOARD = 1

Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type

Private strMvCharacters As String
Private lngMvHwnd  As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Dim ValorAnterior As String
Dim PesoSinFormato As String
Dim peso1 As Double
Private Sub Command1_Click()

Dim fso, txtfile
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtfile = fso.CreateTextFile(App.Path + "\puerto.txt", True)
txtfile.Write (cbbpuerto.Text)  ' Escribe una línea.
txtfile.Close

Label1.Visible = False
ProgressBar1.Visible = False
Timer1.Interval = 1000
ProgressBar1.Max = 6
On Error GoTo openerror
MSComm1.CommPort = cbbpuerto.Text

On Error GoTo openerror
MSComm1.PortOpen = True
ProgressBar1.Visible = True
Timer1.Enabled = True

peso1 = 0
Text1.Text = ""
Timer2.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Label1.Caption = "INICIADO..."
Label1.ForeColor = &HC000&
Label1.Left = 240
Label1.Top = 120
Label1.Visible = True
Exit Sub

openerror:
MsgBox "No se pudo establecer comunicación. Verifique la conexión y configuración del puerto.", vbCritical, "Aiosoft Balanza"
Command1.Enabled = True
Command2.Enabled = False
Label1.Caption = "NO CONECTADO."
Label1.ForeColor = &HFF&
Label1.Left = 240
Label1.Top = 120
Label1.Visible = True

End Sub

Private Sub Command2_Click()
MSComm1.PortOpen = False
Command1.Enabled = True
Timer2.Enabled = False
Command2.Enabled = False
Label1.Caption = "NO CONECTADO."
Label1.ForeColor = &HFF&
Label1.Left = 240
Label1.Top = 120
Label1.Visible = True
End Sub

Private Sub Command3_Click()
Form1.Hide
End Sub

Private Sub Command4_Click()
sysTray1.RemoverSystray
End
End Sub

Private Sub Form_Load()
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(App.Path + "\puerto.txt")
    Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
    nropto = ts.ReadLine
    ts.Close

For x = 1 To 20
    cbbpuerto.AddItem x
Next x
cbbpuerto.Text = nropto

sysTray1.PonerSystray
sysTray1.ToolTipText = " Aiosoft Balanza "

If App.PrevInstance Then Unload Me
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

ProgressBar1.Visible = False
Timer1.Interval = 1000
ProgressBar1.Max = 6

On Error GoTo openerror
MSComm1.CommPort = cbbpuerto.Text

On Error GoTo openerror
MSComm1.PortOpen = True
ProgressBar1.Visible = True
Timer1.Enabled = True

peso1 = 0
Text1.Text = ""
Timer2.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Label1.Caption = "INICIADO..."
Label1.ForeColor = &HC000&
Label1.Left = 240
Label1.Top = 120
Label1.Visible = True

Exit Sub

openerror:
MsgBox "No se pudo establecer comunicación. Verifique la conexión y configuración del puerto.", vbCritical, "Aiosoft Balanza"
Command1.Enabled = True
Command2.Enabled = False
Label1.Caption = "NO CONECTADO."
Label1.ForeColor = &HFF&
Label1.Left = 240
Label1.Top = 120
Label1.Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub sysTray1_MouseDown(Button As Integer)

Form1.Show

End Sub



Private Sub Text1_Change()
Dim PesoFinal As String
Dim largoPesoFinal As Integer
Dim Caracter As String
Dim x, y As Integer
Dim lngLvReturnValue As Long

lngLvReturnValue = SetForegroundWindow(lngMvHwnd)

largoPesoFinal = Len(Text1.Text)
PesoFinal = Text1.Text
For y = 1 To 6
    SendKey 8
Next y
For x = 1 To largoPesoFinal
    Caracter = Mid(PesoFinal, x, 1)
    SendKey Asc(Caracter)
Next x
End Sub

Private Sub Timer1_Timer()
Static intTime ' Declara la variable estática.
   ' La primera vez, la variable estará vacía.
   ' Establece la variable a 1 si está vacía.
   If IsEmpty(intTime) Then intTime = 1
   
   ProgressBar1.Value = intTime ' Update the ProgressBar.
       
   If intTime = ProgressBar1.Max Then
      Timer1.Enabled = False
      ProgressBar1.Visible = False
      intTime = 1
      ProgressBar1.Value = ProgressBar1.Min
      Form1.Hide
   Else
      intTime = intTime + 1
   End If

End Sub

Private Sub Timer2_Timer()

   
Static intTime2 ' Declara la variable estática.
 'La primera vez, la variable estará vacía.
' Establece la variable a 1 si está vacía.
If IsEmpty(intTime2) Then intTime2 = 0
If intTime2 = 1 Then
    'borra todo el contenido del buffer
    MSComm1.InBufferCount = 0
End If
If intTime2 >= 2 Then
    
        On Error GoTo errorconexion
        PesoSinFormato = MSComm1.Input
        largocifra = InStr(1, PesoSinFormato, ".", vbTextCompare)
        largocifra2 = largocifra - 1
        If Val(largocifra2) >= 1 Then
            Pesoengramos = Mid(PesoSinFormato, 1, largocifra2)
            If Val(Pesoengramos) <= 20 Then
                Text1.Text = ""
            Else
                If Pesoengramos > 3 Then
                    Text1.Text = Format(Val(Pesoengramos) / 1000, "#0.000")
                    
                End If
            End If
        Else
            'Text1.Text = ""
        End If
        
    
End If
intTime2 = intTime2 + 1
Exit Sub
errorconexion:
MsgBox "Se ha perdido la comunicación con la balanza. Reconecte el dispositivo y presione el botón reiniciar balanza.", vbCritical, "Aiosoft Balanza"
Command1.Enabled = True
Command2.Enabled = False
Timer2.Enabled = False
MSComm1.PortOpen = False
Label1.Caption = "NO CONECTADO."
Label1.ForeColor = &HFF&
Label1.Left = 240
Label1.Top = 120
Label1.Visible = True
End Sub

Private Sub Timer3_Timer()
 On Error GoTo ErrorHandler
    Dim ctlLvControl As Control
    Dim blnLvActiveWindowIsWithinThisApplication  As Boolean
    If GetForegroundWindow() <> Me.hwnd Then
        blnLvActiveWindowIsWithinThisApplication = False
        For Each ctlLvControl In Me.Controls
            If ctlLvControl.hwnd <> GetForegroundWindow() Then
            Else
                blnLvActiveWindowIsWithinThisApplication = True
                Exit For
            End If
        Next
        If Not blnLvActiveWindowIsWithinThisApplication Then
            lngMvHwnd = GetForegroundWindow()
        End If
    End If
    GoTo OverError
    
ErrorHandler:
    Select Case Err.Number
    Case 438
        Resume Next
    Case Else
    End Select
    GoTo OverError

OverError:
End Sub
Private Sub SendKey(bKey As Byte)
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    
        KInput.wVk = bKey
        If KInput.wVk = 46 Then
            KInput.wVk = 110
        End If
        KInput.dwFlags = 0
        GInput(0).dwType = INPUT_KEYBOARD
        CopyMemory GInput(0).xi(0), KInput, Len(KInput)
        
        KInput.wVk = bKey
        KInput.dwFlags = KEYEVENTF_KEYUP
        GInput(1).dwType = INPUT_KEYBOARD
        CopyMemory GInput(1).xi(0), KInput, Len(KInput)
        Call SendInput(2, GInput(0), Len(GInput(0)))
End Sub
