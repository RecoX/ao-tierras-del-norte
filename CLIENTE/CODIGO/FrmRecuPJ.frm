VERSION 5.00
Begin VB.Form FrmRecuPJ 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   LinkTopic       =   "Form4"
   Picture         =   "FrmRecuPJ.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   400
      Left            =   480
      TabIndex        =   2
      Top             =   3040
      Width           =   2175
   End
   Begin VB.TextBox DATO2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   400
      Left            =   480
      TabIndex        =   1
      Top             =   2210
      Width           =   2175
   End
   Begin VB.TextBox DATO1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   400
      Left            =   480
      TabIndex        =   0
      Top             =   1350
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2640
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRecuPJ"
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
