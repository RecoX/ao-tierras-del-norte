VERSION 5.00
Begin VB.Form FrmRECBORR 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   5235
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   3495
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrnRECBORR.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   325
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox DATO2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   325
      Left            =   360
      TabIndex        =   1
      Top             =   2630
      Width           =   2655
   End
   Begin VB.TextBox DATO1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   325
      Left            =   360
      TabIndex        =   0
      Top             =   1660
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3060
      Top             =   100
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   840
      Top             =   4440
      Width           =   1695
   End
End
Attribute VB_Name = "FrmRECBORR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modo As Byte

Private Sub Image1_Click()

If Modo = 1 Then
Me.Picture = LoadPicture(App.path & "\Recursos\VentanaBorrar.jpg")
        UserName = DATO1
        UserPassword = DATO2
        UserPin = txtPIN
ElseIf Modo = 2 Then
Call FrmRecuPJ.Show(vbModeless, FrmPanelUsuario)
        UserName = DATO1
        UserEmail = DATO2
        UserPin = txtPIN
Else


'Si por X razon llegamos aca y no se asigno el modo
MsgBox "Ocurrio un error En el proceso. Reintentelo."
        Unload Me
        Exit Sub
        End If
Call Login

Unload Me
End Sub

Private Sub Form_Load()
If EstadoLogin = E_MODO.BorrarPJ Then
Me.Picture = LoadPicture(App.path & "\Recursos\VentanaBorrar.jpg")
Modo = 1
ElseIf EstadoLogin = E_MODO.RecuperarPJ Then
Call FrmRecuPJ.Show(vbModeless, FrmPanelUsuario)
Modo = 2
End If
End Sub


Private Sub Image2_Click()
Unload Me
End Sub
