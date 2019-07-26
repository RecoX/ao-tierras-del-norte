VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creaci�n de un Clan"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildFoundation.frx":0000
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Legi�n Oscura"
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   2760
      TabIndex        =   4
      Top             =   3240
      Width           =   990
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Neutral"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "ArmadaReal"
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   990
   End
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtWeb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   495
      Left            =   2160
      Tag             =   "1"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image imgCancelar 
      Height          =   495
      Left            =   240
      Tag             =   "1"
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful, Porn Hub
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/ SABEEEEEEEEEEEEEEEEEEEEEEE
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit
Private Enum eAlineacion
    ieREAL = 0
    ieCAOS = 1
    ieNeutral = 2
    ieLegal = 4
    ieCriminal = 5
End Enum
Private clsFormulario As clsFormMovementManager

Private cBotonSiguiente As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

  '  Me.Picture = LoadPicture(App.path & "\graficos\VentanaNombreClan.jpg")
        
    If Len(txtClanName.Text) <= 30 Then
        If Not AsciiValidos(txtClanName) Then
            MsgBox "Nombre invalido."
            Exit Sub
        End If
    Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
    End If
    

End Sub



Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgSiguiente_Click()
        If Len(txtClanName.Text) <= 0 And Len(txtWeb.Text) <= 0 Then
    MsgBox "Debe rellenar todos los datos"
    Exit Sub
    End If
    If Option1.value = 0 And Option2.value = 0 And Option3.value = 0 Then
    MsgBox "Debes elegir alguna alineaci�n"
    Exit Sub
    End If
    ClanName = txtClanName.Text
    Site = txtWeb.Text
    Unload Me
    frmGuildDetails.Show , frmMain
End Sub

Private Sub Option1_Click()
If Option3.value = 1 Then Option3.value = 0
If Option2.value = 1 Then Option2.value = 0
WriteGuildFundation eAlineacion.ieREAL
End Sub

Private Sub Option2_Click()
'PAJ ALSJFLASJKFASFASFASF
If Option3.value = 1 Then Option3.value = 0
If Option1.value = 1 Then Option1.value = 0
Call WriteGuildFundation(eAlineacion.ieNeutral)
End Sub

Private Sub Option3_Click()
If Option1.value = 1 Then Option1.value = 0
If Option2.value = 1 Then Option2.value = 0
WriteGuildFundation eAlineacion.ieCAOS
End Sub

