VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   4650
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmCarp.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H0080C0FF&
      Height          =   1950
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3840
   End
   Begin VB.Image Command4 
      Height          =   375
      Left            =   240
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image Command3 
      Height          =   495
      Left            =   2880
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "frmCarp"
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

Private Sub Command3_Click()
    On Error Resume Next

    Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1))
    If frmMain.macrotrabajo.Enabled Then _
        MacroBltIndex = ObjCarpintero(lstArmas.ListIndex + 1)
    
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub
