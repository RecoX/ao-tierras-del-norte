VERSION 5.00
Begin VB.Form FrmPanelUsuario 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LinkTopic       =   "Form4"
   Picture         =   "FrmPanelUsuario.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2760
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgRecuPj 
      Height          =   495
      Left            =   480
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   480
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "FrmPanelUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

EstadoLogin = E_MODO.BorrarPJ
 
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub imgRecuPj_Click()
Call Audio.PlayWave(SND_CLICK)
EstadoLogin = E_MODO.RecuperarPJ
 
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
 
End Sub

