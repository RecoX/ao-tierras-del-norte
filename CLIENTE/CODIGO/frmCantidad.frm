VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   1560
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCantidad.frx":0000
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   221
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   405
      TabIndex        =   0
      Top             =   600
      Width           =   2550
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   360
      Tag             =   "1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   2040
      Tag             =   "1"
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X As Single
Private Y As Single
Private tX As Byte
Private tY As Byte
Private MouseX As Long
Private MouseY As Long

Private Sub Command1_Click()
 If LenB(frmCantidad.Text1) > 0 Then
        If Not IsNumeric(frmCantidad.Text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then
        If Text1 > Inventario.amount(Inventario.SelectedItem) Then
        ShowConsoleMsg "No tienes esa cantidad!"
        Unload Me
        Exit Sub
        End If
        
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, Text1
        frmCantidad.Text1.Text = ""
        
        Else
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.Text1.Text)
        frmCantidad.Text1.Text = ""
    End If
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
If Inventario.SelectedItem = 0 Then Exit Sub
    ' If LenB(frmcantidad.Text1) > 0 Then
       ' If Not IsNumeric(frmcantidad.Text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then 'drag and drop
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem)
        frmCantidad.Text1.Text = ""
Unload Me
    
Else 'tirar al piso :D
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem))
        Unload Me
    Else
        If UserGLD > 100000 Then
            Call WriteDrop(Inventario.SelectedItem, 100000)
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            Unload Me
        End If
    End If
    End If
    frmCantidad.Text1.Text = ""
    
    Unload Me
End Sub

Private Sub Text1_Change()
On Error GoTo ErrHandler
    If Val(frmCantidad.Text1) < 0 Then
        frmCantidad.Text1 = "1"
    End If
    
    If Val(frmCantidad.Text1) > 100000 Then
        frmCantidad.Text1 = "100000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    frmCantidad.Text1 = "1"
End Sub


