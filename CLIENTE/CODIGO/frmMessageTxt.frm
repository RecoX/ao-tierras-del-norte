VERSION 5.00
Begin VB.Form frmMessageTxt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mensajes Predefinidos"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmMessageTxt.frx":0000
   ScaleHeight     =   4710
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   3840
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      Top             =   3400
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   6
      Top             =   2630
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   265
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   2220
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Text            =   " "
      Top             =   1400
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   290
      Index           =   2
      Left            =   1210
      TabIndex        =   2
      Top             =   960
      Width           =   3285
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   265
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   580
      Width           =   3315
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1210
      TabIndex        =   0
      Top             =   170
      Width           =   3285
   End
   Begin VB.Image ImgCancelar 
      Height          =   375
      Left            =   2880
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image ImgGuardar 
      Height          =   375
      Left            =   1080
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "frmMessageTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonGuardar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    Dim i As Long
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    For i = 0 To 9
        messageTxt(i) = CustomMessages.Message(i)
    Next i

  '  Me.Picture = LoadPicture(App.path & "\graficos\VentanaMensajesPersonalizados.jpg")
    
    LoadButtons
    
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonGuardar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

   ' Call cBotonGuardar.Initialize(imgGuardar, GrhPath & "BotonGuardarCustomMsg.jpg", GrhPath & "BotonGuardarRolloverCustomMsg.jpg", _
                                    GrhPath & "BotonGuardarClickCustomMsg.jpg", Me)
   ' Call cBotonCancelar.Initialize(ImgCancelar, GrhPath & "BotonCancelarCustomMsg.jpg", GrhPath & "BotonCancelarRolloverCustomMsg.jpg", _
                                    GrhPath & "BotonCancelarClickCustomMsg.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgGuardar_Click()
On Error GoTo ErrHandler
    Dim i As Long
    
    For i = 0 To 9
        CustomMessages.Message(i) = messageTxt(i)
    Next i
    
    Unload Me
Exit Sub

ErrHandler:
    'Did detected an invalid message??
    If Err.number = CustomMessages.InvalidMessageErrCode Then
        Call MsgBox("El Mensaje " & CStr(i + 1) & " es inválido. Modifiquelo por favor.")
    End If

End Sub

Private Sub messageTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub
