VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4560
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   1425
      Left            =   7800
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   3240
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   2514
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCargando.frx":48302
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3120
      Left            =   1560
      ScaleHeight     =   208
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   536
      TabIndex        =   0
      Top             =   10800
      Width           =   8040
   End
   Begin VB.Image barra 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   900
      Picture         =   "frmCargando.frx":48381
      Top             =   2040
      Width           =   4110
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim F As Integer


Private Sub Form_Load()

barra.Width = 1

'Analizar

End Sub

Function Analizar()
     On Error Resume Next
    
     Dim iX As Integer
     Dim tX As Integer
     Dim DifX As Integer
    
    iX = Inet1.OpenURL("http://tdn-ao-ofi.ucoz.es/VEREXE") 'Host
    tX = LeerInt(App.path & "\INIT\Update.ini")
    DifX = iX - tX
 
    If Not (DifX = 0) Then
        'If MsgBox("Hay una actualizaci�n disponible. �Desea ejecutar el AutoUpdate para descargarla?", vbYesNo, "Tierras del Norte AO") = vbYes Then
            Call ShellExecute(0, "open", App.path & "/Autoupdate.exe", "", "", 1)
            End
        'End If
    End If

End Function
Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function
 
Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, data
    Close #F
End Sub

