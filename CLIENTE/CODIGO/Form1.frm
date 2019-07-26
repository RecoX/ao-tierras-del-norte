VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   -360
      TabIndex        =   3
      Top             =   600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCommand3 
      Caption         =   "Command2"
      Height          =   360
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   990
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UltimaCadena As Integer

Private Function Eliminar_Item_ListView(ListView As ListView) As Long

' -- Variables
    Dim i As Long
    Dim J As Long
    Dim ret As Long         ' -- total de items que se eliminan

    With ListView
        ' -- Recorrer todos los items
        For i = 1 To .ListItems.Count
            ' -- Comparar uno a uno con todos los dem�s
            For J = i + 1 To .ListItems.Count
                If .ListItems.Item(i) = .ListItems.Item(J) Then
                    ' -- Si es igual eliminar
                    .ListItems.Remove .ListItems.Item(J).Index
                    J = J - 1
                    ret = ret + 1
                End If
                If J = .ListItems.Count Then
                    Exit For
                End If
            Next

            If i = .ListItems.Count Then
                ' -- Retorna el valor de la funci�n con _
                  la cantidad de elementos eliminados
                CANTt = .ListItems.Count
                Eliminar_Item_ListView = ret
                Exit Function
            End If
        Next
    End With

End Function

Public Sub AnalizarThreads()

'MEMORY_BASIC_INFORMATION mbi;
    Dim mbi As MEMORY_BASIC_INFORMATION
    'MODULE_INFORMATION mi;
    Dim mi As MODULE_INFORMATION
    'BYTE szBuffer[MAX_PATH * 2 + 4] = { 0 };
    Dim szBuffer(523) As Byte
    Dim i As Long

    'PUNICODE_STRING usSectionName;
    Dim usSectionName As UNICODE_STRING
    Dim hProcess As Long
    hProcess = GetCurrentProcess()
    Dim Addr As Long
    Dim READABLE As Long
    READABLE = (PAGE_EXECUTE_READ + PAGE_EXECUTE_READWRITE + PAGE_EXECUTE_WRITECOPY + PAGE_READONLY + PAGE_READWRITE + PAGE_WRITECOPY)
    Form1.ListView1.ListItems.Clear
    Addr = 0

    Dim hRet As Long
    Dim zBytes() As Byte
    ReDim zBytes(0) As Byte
    While VirtualQuery(Addr, mbi, 28)
        DoEvents
        '        Form1.lblAddress.Caption = "0x" & Hex(Addr)

        If (mbi.State And MEM_COMMIT) Then

            If (mbi.AllocationProtect And READABLE) Then

                hRet = ZwQueryVirtualMemory(hProcess, Addr, MemoryBasicInformation, VarPtr(mbi), &H1C, 0&)

                For i = LBound(szBuffer) To UBound(szBuffer)
                    szBuffer(i) = 0
                Next i

                For i = LBound(zBytes) To UBound(zBytes)
                    zBytes(i) = 0
                Next i

                If (hRet >= 0) Then
                    If (mbi.Type <> MEM_FREE) Then
                        hRet = ZwQueryVirtualMemory(hProcess, Addr, MemorySectionName, VarPtr(szBuffer(0)), &H20C, 0&)

                        If (hRet >= 0) Then
                            Call ZeroMemory(mi, &H234)
                            Call RtlMoveMemory(mi, mbi, &H1C)
                            Call ReadProcessMemory(hProcess, VarPtr(szBuffer(0)), usSectionName.length, &H2, 0&)
                            Call ReadProcessMemory(hProcess, VarPtr(szBuffer(2)), usSectionName.MaximumLength, &H2, 0&)
                            ReDim zBytes(usSectionName.length * 2)

                            'How do I know is offset 8? It's simple.... "Aliens"
                            Call ReadProcessMemory(hProcess, VarPtr(szBuffer(8)), zBytes(0), usSectionName.length * 2, 0&)

                            Dim TempString As String

                            TempString = Trim$(ByteArrayToString(zBytes))

                            If EsNormal(TempString) Then        'Evito ciertas DLL al pedo

                                CANTt = CANTt + 1

                                ListView1.ListItems.Add , , TempString        ' Asi lo tenia y funcionaba jaja.

                            End If

                        End If
                    End If
                End If
            End If
        End If

        If Addr >= &H7FFF0000 Then
            GoTo Salir
        End If
        Addr = (mbi.BaseAddress) + mbi.RegionSize
    Wend
Salir:
    Call Eliminar_Item_ListView(ListView1)    ' Esto uso para evitar duplicados
End Sub

Public Function EsNormal(ByVal Cadena As String) As Boolean

    Cadena = UCase$(Cadena)

    If Cadena = vbNullString Then Cadena = "Vacio": Exit Function

    ' ++ Optimizacion comprobamos que no se repita la cadena anterior y usamos InStrRev ya que despues del .dll no hay otra ruta xdxd logica mistica de la nasa ahre
    If Len(Cadena) <> UltimaCadena Then

        If InStrRev(Cadena, ".DLL") Then

            EsNormal = True

            UltimaCadena = Len(Cadena)    ' ++ Mejor comparar bytes que strings

            Exit Function

        End If

    End If

    EsNormal = False

End Function

