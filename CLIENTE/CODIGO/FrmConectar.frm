VERSION 5.00
Begin VB.Form FrmConectar 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   LinkTopic       =   "Form4"
   Picture         =   "FrmConectar.frx":0000
   ScaleHeight     =   3225
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkrecu 
      BackColor       =   &H8000000B&
      Caption         =   "Recordar Password"
      Height          =   195
      Left            =   1620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2510
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   370
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1770
      Width           =   4445
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   370
      Left            =   480
      TabIndex        =   0
      Top             =   745
      Width           =   4445
   End
   Begin VB.Label Volver 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image imgConectar 
      Height          =   375
      Left            =   4200
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "FrmConectar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cerrar_Click()

End Sub

Private Sub imgConectar_Click()
  '  Call CheckServers
    If checkrecu.value = 1 Then
    If Not StringIsRecup(UCase$(txtNombre.Text)) Then
        Call SaveRecu(txtNombre.Text, txtPasswd.Text)
    End If
End If
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String
    aux = txtPasswd.Text
    
#If SeguridadAlkon Then
    UserPassword = md5.GetMD5String(aux)
    Call md5.MD5Reset
#Else
    UserPassword = aux
#End If
    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    End If
    
End Sub

Private Sub txtNombre_Change()
 
Dim F As Long
For F = 1 To MaxRecu
    If UCase$(txtNombre.Text) = UCase$(Recu(F).Nick) Then
        txtPasswd.Text = Recu(F).Password
    End If
Next F
 
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectar_Click
End Sub

Private Sub Volver_Click()
Unload Me
End Sub
