VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   255
   ClientLeft      =   6000
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Interval        =   3000
      Left            =   660
      Top             =   -60
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1710
      Top             =   0
   End
   Begin VB.Label lblTexto 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000009&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' api para poner la ventana siempre visible
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    
'CargarUsuario
Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type
Private Const ERROR_SUCCESS            As Long = 0
Private Const MIB_TCP_STATE_CLOSED     As Long = 1
Private Const MIB_TCP_STATE_LISTEN     As Long = 2
Private Const MIB_TCP_STATE_SYN_SENT   As Long = 3
Private Const MIB_TCP_STATE_SYN_RCVD   As Long = 4
Private Const MIB_TCP_STATE_ESTAB      As Long = 5
Private Const MIB_TCP_STATE_FIN_WAIT1  As Long = 6
Private Const MIB_TCP_STATE_FIN_WAIT2  As Long = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
Private Const MIB_TCP_STATE_CLOSING    As Long = 9
Private Const MIB_TCP_STATE_LAST_ACK   As Long = 10
Private Const MIB_TCP_STATE_TIME_WAIT  As Long = 11
Private Const MIB_TCP_STATE_DELETE_TCB As Long = 12
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal addr As Long) As Long
'-----------------
'Recuperar_Nombre_Host
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    imaxsockets As Integer
    imaxudp As Integer
    lpszvenderinfo As Long
End Type
  
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Private strUsuarios As String
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Private Sub Form_Load()
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub Timer_Timer()
   Call CargarUsuario
   lblTexto.Caption = strUsuarios
End Sub

Private Sub Timer1_Timer()
   
   If lblTexto.Left <= -3500 Then
      lblTexto.Left = 3500
   End If
   lblTexto.Left = lblTexto.Left - 70
End Sub
Private Sub CargarUsuario()

Dim TcpRow As MIB_TCPROW
Dim buff() As Byte
Dim lngRequired As Long
Dim lngStrucSize As Long
Dim lngRows As Long
Dim lngCnt As Long
Dim strTmp As String


strUsuarios = "VNC: "

Call GetTcpTable(ByVal 0&, lngRequired, 1)

If lngRequired > 0 Then
    ReDim buff(0 To lngRequired - 1) As Byte
    If GetTcpTable(buff(0), lngRequired, 1) = ERROR_SUCCESS Then
        lngStrucSize = LenB(TcpRow)
        CopyMemory lngRows, buff(0), 4

        For lngCnt = 1 To lngRows
            CopyMemory TcpRow, buff(4 + (lngCnt - 1) * lngStrucSize), lngStrucSize
            If ntohs(TcpRow.dwLocalPort) = 5900 And TcpRow.dwState = MIB_TCP_STATE_ESTAB Then 'puerto del vnc
               strUsuarios = strUsuarios & " " & "(" & Replace(Right(GetInetAddrStr(TcpRow.dwRemoteAddr), 3), ".", "") & ") " & UCase(Recuperar_Nombre_Host(GetInetAddrStr(TcpRow.dwRemoteAddr)))
            End If
        Next
    End If
End If

End Sub
Private Function GetString(ByVal lpszA As Long) As String
    GetString = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetString, ByVal lpszA)
End Function
Private Function GetInetAddrStr(Address As Long) As String
    GetInetAddrStr = GetString(inet_ntoa(Address))
End Function

Private Function Recuperar_Nombre_Host(ByVal direccion_IP As String) As String
   Dim PH As Long, hDir As Long, nb As Long
   Dim W As WSADATA
   
   If WSAStartup(&H101, W) = 0 Then
      hDir = inet_addr(direccion_IP)
        
      If hDir <> -1 Then
        
         PH = gethostbyaddr(hDir, 4, 2)
         If PH <> 0 Then
            CopyMemory PH, ByVal PH, 4
            nb = lstrlen(ByVal PH)
            If nb > 0 Then
               direccion_IP = Space$(nb)
               CopyMemory ByVal direccion_IP, ByVal PH, nb
               Recuperar_Nombre_Host = Replace(direccion_IP, ".algoritmo.local", "")
            End If
         Else
             Recuperar_Nombre_Host = direccion_IP
         End If
         If WSACleanup() <> 0 Then
            Recuperar_Nombre_Host = direccion_IP
         End If
      Else
         Recuperar_Nombre_Host = direccion_IP
      End If
   Else
      Recuperar_Nombre_Host = direccion_IP
   End If
   
End Function


