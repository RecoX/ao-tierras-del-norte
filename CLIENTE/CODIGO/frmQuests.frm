VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Misiones"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
      Height          =   3735
      Left            =   2340
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   60
      Width           =   3015
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3420
      Width           =   2235
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Abandonar misi�n"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   2235
   End
   Begin VB.ListBox lstQuests 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2235
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Private Sub cmdOptions_Click(Index As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click de los CommandButtons cmdOptions.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Select Case Index
        Case 0 'Bot�n ABANDONAR MISI�N
            'Chequeamos si hay items.
            If lstQuests.ListCount = 0 Then
                MsgBox "�No tienes ninguna misi�n!", vbOKOnly + vbExclamation
                Exit Sub
            End If
            
            'Chequeamos si tiene algun item seleccionado.
            If lstQuests.ListIndex < 0 Then
                MsgBox "�Primero debes seleccionar una misi�n!", vbOKOnly + vbExclamation
                Exit Sub
            End If
            
            Select Case MsgBox("�Est�s seguro que deseas abandonar la misi�n?", vbYesNo + vbExclamation)
                Case vbYes  'Bot�n S�.
                    'Enviamos el paquete para abandonar la quest
                    Call WriteQuestAbandon(lstQuests.ListIndex + 1)
                    
                Case vbNo   'Bot�n NO.
                    'Como seleccion� que no, no hace nada.
                    Exit Sub
            End Select
            
        Case 1 'Bot�n VOLVER
            Unload Me
    End Select
End Sub

Private Sub lstQuests_Click()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click del ListBox lstQuests.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If lstQuests.ListIndex < 0 Then Exit Sub
    
    Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)
End Sub
