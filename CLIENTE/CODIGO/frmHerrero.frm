VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHerrero.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmaduras 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1740
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3840
   End
   Begin VB.ListBox lstArmas 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   360
      TabIndex        =   0
      Top             =   1420
      Width           =   3855
   End
   Begin VB.Image Command3 
      Height          =   375
      Left            =   285
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image Command4 
      Height          =   375
      Left            =   2805
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image Command2 
      Height          =   495
      Left            =   2640
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image Command1 
      Height          =   495
      Left            =   360
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmHerrero"
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
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
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

Private Sub Command1_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next

    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1))
        
        If frmMain.macrotrabajo.Enabled Then _
            MacroBltIndex = ArmasHerrero(lstArmas.ListIndex + 1)
    Else
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1))
        
        If frmMain.macrotrabajo.Enabled Then _
            MacroBltIndex = ArmadurasHerrero(lstArmaduras.ListIndex + 1)
    End If

    Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

