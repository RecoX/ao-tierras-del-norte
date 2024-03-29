VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsParty.cls
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer
'el creador

Private p_CantMiembros As Integer
'cantidad de miembros

Private p_SumaNivelesElevados As Single
'suma de todos los niveles elevados a la ExponenteNivelParty > Esta variable se usa para calcular la experiencia repartida en la Party.

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 p_SumaNiveles added (Tavo)
' - 07/04/08 p_SumaNiveles changed to p_SumaNivelesElevados (MarKoxX)
'***************************************************
    p_expTotal = 0
    p_CantMiembros = 0
    p_SumaNivelesElevados = 0
    
    ' // maTih.- //
    Call initPorcentajes
    
End Sub

'Destructor de clase
Public Sub Class_Terminate()
End Sub

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param lvl Specifies reference to user level
' @remarks When a user level up and he is in a party, we update p_sumaNivelesElavados so the formula still works.
Public Sub UpdateSumaNivelesElevados(ByVal Lvl As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 11/24/09
'11/24/09: Pato - Change the exponent to a variable with the exponent
'*************************************************
    p_SumaNivelesElevados = p_SumaNivelesElevados - ((Lvl - 1) ^ ExponenteNivelParty) + Lvl ^ ExponenteNivelParty
End Sub

Public Function MiExperiencia(ByVal Userindex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: 11/27/09
'Last Modification By: Budi
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
' - 11/27/09 Arregl� el Out of Range.
'***************************************************
'Me dice cuanta experiencia tengo colectada ya en la party
Dim i As Integer
i = 1

While i <= PARTY_MAXMEMBERS And p_members(i).Userindex <> Userindex
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    MiExperiencia = Fix(p_members(i).Experiencia)
Else        'esto no deberia pasar :p
    MiExperiencia = -1
End If

End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal mapa As Integer, X As Integer, Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 New formula for calculating the experience point of each user
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
' - 04/04/08 Ahora antes de calcular la experiencia a X usuario se fija si ese usuario existe (MarKoxX)
' - 07/04/08 New formula to calculate Experience for each user. (MarKoxX)
'***************************************************
'Se produjo un evento que da experiencia en la wp referenciada
Dim i As Integer
Dim UI As Integer
Dim expThisUser As Double

p_expTotal = p_expTotal + ExpGanada

For i = 1 To PARTY_MAXMEMBERS
    UI = p_members(i).Userindex
    If UI > 0 Then
        ' Formula: Exp* (Nivel ^ ExponenteNivelParty) / sumadeNivelesElevados
        ' expThisUser = CDbl(ExpGanada * (UserList(p_members(i).UserIndex).Stats.ELV ^ ExponenteNivelParty) / p_SumaNivelesElevados)
        
        expThisUser = Porcentaje(ExpGanada, p_members(i).bPorcentaje)
        
        If mapa = UserList(UI).Pos.map And UserList(UI).flags.Muerto = 0 Then
            If Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                p_members(i).Experiencia = p_members(i).Experiencia + expThisUser
                If p_members(i).Experiencia < 0 Then
                    p_members(i).Experiencia = 0
                End If
                If PARTY_EXPERIENCIAPORGOLPE Then
                    UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + Fix(expThisUser)
                    If UserList(UI).Stats.Exp > MAXEXP Then _
                        UserList(UI).Stats.Exp = MAXEXP
                    Call CheckUserLevel(UI)
                    Call WriteUpdateUserStats(UI)
                End If
            End If
        End If
    End If
Next i

End Sub

Public Sub initPorcentajes()

    '
    ' @ maTih.-
    
    Dim pNumUsers   As Integer
    
    Dim j As Long
    
    For j = 1 To PARTY_MAXMEMBERS
        If p_members(j).Userindex > 0 Then pNumUsers = (pNumUsers + 1)
    Next j
    
    If (pNumUsers > 0) Then
        For j = 1 To PARTY_MAXMEMBERS
            If p_members(j).Userindex > 0 Then
               p_members(j).bPorcentaje = getPorcentaje(pNumUsers)
               
               Call Protocol.WriteConsoleMsg(p_members(j).Userindex, "Party> Porcentajes actualizados." & vbNewLine & "Porcentaje de experiencia por usuario:" & CStr(p_members(j).bPorcentaje) & ".", FontTypeNames.FONTTYPE_PARTY)
            End If
        Next j
    End If
    
End Sub

Public Function preparePorcentajeString() As String

    '
    ' @ maTih.-
    
    Dim Str As String
    Dim j   As Long
    Dim tmp As Integer
    
    For j = 1 To PARTY_MAXMEMBERS
        tmp = p_members(j).Userindex
        
        If (tmp > 0) Then
            If (Str = vbNullString) Then
                Str = UserList(tmp).name & "*" & CStr(p_members(j).Experiencia) & "*" & CStr(p_members(j).bPorcentaje)
            Else
                Str = Str & "," & UserList(tmp).name & "*" & CStr(p_members(j).Experiencia) & "*" & (p_members(j).bPorcentaje)
            End If
        End If
    Next j
    
    preparePorcentajeString = Str

End Function

Public Function stringToArray(ByRef strCompare As String) As Byte()

    '
    ' @ maTih.- Chupa la pija, aca manda Fran loco
    
    Dim tempArray() As String
    Dim endArray()  As Byte
    Dim tmpPosition As Integer
    Dim nowMember   As String
    Dim MemberIndex As Integer
    Dim TempInt     As Integer
    
    ' initialize ;
    ReDim endArray(1 To PARTY_MAXMEMBERS) As Byte
    
    ' set array ;
    tempArray = Split(strCompare, ",")
                  Dim pt As String
                  Dim pt1 As String
                  Dim pt2 As String
                  Dim pt3 As String
    Dim j As Long
    
    For j = 0 To UBound(tempArray())
        If Not tempArray(j) = vbNullString Then
           tmpPosition = InStr(1, tempArray(j), "*")
           
           
           If (tmpPosition > 0) Then
               ' set member name
               nowMember = Left$(tempArray(j), (tmpPosition - 1))
               
               ' find memberIndex ;
               If Not (nowMember = "Nadie") Then
                  MemberIndex = findMemberIndex(UCase$(nowMember))
               Else
                  MemberIndex = 0
               End If

                pt = mid$(tempArray(j), (tmpPosition + 1), 2)
          
               
               If (MemberIndex > 0) Then
                   TempInt = val(mid$(tempArray(j), tmpPosition + 1, 2))
                '  TempInt = val(Left$(JIji, pt)) matIh.- pT, El fran vieja^^ mira arriba papa
                
                   endArray(MemberIndex) = CByte(TempInt)
               End If

               
           End If
        End If
    Next j
    
    stringToArray = endArray()

End Function

Private Function findMemberIndex(ByRef mName As String) As Integer

    '
    ' @ maTih.-
    
    ' default ;
    findMemberIndex = 0
    
    Dim j As Long
    Dim N As Integer
    
    N = NameIndex(mName)
    
    If (N = 0) Then Exit Function
    
    For j = 1 To PARTY_MAXMEMBERS
        If p_members(j).Userindex = N Then
            findMemberIndex = CInt(j)
            Exit Function
        End If
    Next j

End Function

Public Sub setPorcentajes(ByRef bPorcentajeArray() As Byte)

    '
    ' @ maTih.-
    
    Dim j As Long
    Dim N As Integer
    Dim f As String
    
    For j = 1 To PARTY_MAXMEMBERS
        p_members(j).bPorcentaje = bPorcentajeArray(j)
        
        N = p_members(j).Userindex
        
        If (N > 0) Then
            If (f = vbNullString) Then
                f = UserList(N).name & " [" & p_members(j).bPorcentaje & "%] "
            Else
                f = f & UserList(N).name & " [" & p_members(j).bPorcentaje & "%] "
            End If
        End If
    Next j
    
    Call MandarMensajeAConsola("Porcentajes de la party actualizados." & vbNewLine & f & ".", "PARTY")
    
End Sub

Private Function getPorcentaje(ByVal bNumUsers As Byte) As Byte

    '
    ' @ maTih.-
    
    Dim bValue As Byte
    
    Select Case bNumUsers
           Case 1
                bValue = 100
                
           Case 2
                bValue = 50
            
           Case 3
                bValue = 33
            
           Case 4
                bValue = 25
                
           Case 5
                bValue = 20
    End Select
    
    getPorcentaje = bValue

End Function


Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
Dim i As Integer

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).Userindex > 0 Then
        Call WriteConsoleMsg(p_members(i).Userindex, "[" & Sender & "] " & texto, FontTypeNames.FONTTYPE_PARTY)
    End If
Next i

End Sub

Public Function EsPartyLeader(ByVal Userindex As Integer) As Boolean
    EsPartyLeader = (Userindex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 There is no level prohibition (Tavo)
' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
'***************************************************

Dim i As Integer
i = 1
While i <= PARTY_MAXMEMBERS And p_members(i).Userindex > 0
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    p_members(i).Experiencia = 0
    p_members(i).Userindex = Userindex
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
    p_SumaNivelesElevados = p_SumaNivelesElevados + (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
    
    Call initPorcentajes
Else
    NuevoMiembro = False
End If

End Function

Public Function SaleMiembro(ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number (Tavo)
' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
'***************************************************
'el valor de retorno representa si se disuelve la party
Dim i As Integer
Dim j As Integer
i = 1
SaleMiembro = False
While i <= PARTY_MAXMEMBERS And p_members(i).Userindex <> Userindex
    i = i + 1
Wend

If i = 1 Then
    'sale el founder, la party se disuelve
    SaleMiembro = True
    Call MandarMensajeAConsola("El lider disuelve la party.", "Servidor")
    For j = PARTY_MAXMEMBERS To 1 Step -1
        If p_members(j).Userindex > 0 Then
            
            Call WriteConsoleMsg(p_members(j).Userindex, " Abandonas la party liderada por " & UserList(p_members(1).Userindex).name, FontTypeNames.FONTTYPE_PARTY)
    
            Call WriteConsoleMsg(p_members(j).Userindex, " Durante la misma has conseguido " & CStr(Fix(p_members(j).Experiencia)) & " puntos de experiencia!", FontTypeNames.FONTTYPE_PARTY)
    
            If Not PARTY_EXPERIENCIAPORGOLPE Then
                UserList(p_members(j).Userindex).Stats.Exp = UserList(p_members(j).Userindex).Stats.Exp + Fix(p_members(j).Experiencia)
                If UserList(p_members(j).Userindex).Stats.Exp > MAXEXP Then _
                    UserList(p_members(j).Userindex).Stats.Exp = MAXEXP
                Call CheckUserLevel(p_members(j).Userindex)
                Call WriteUpdateUserStats(p_members(j).Userindex)
            End If
            Call MandarMensajeAConsola(UserList(p_members(j).Userindex).name & " abandona la party.", "Servidor")
            UserList(p_members(j).Userindex).PartyIndex = 0
            p_CantMiembros = p_CantMiembros - 1
            p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
            p_members(j).Userindex = 0
            p_members(j).Experiencia = 0
        End If
    Next j
Else
    If i <= PARTY_MAXMEMBERS Then
        If Not PARTY_EXPERIENCIAPORGOLPE Then
            UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
            If UserList(p_members(i).Userindex).Stats.Exp > MAXEXP Then _
                UserList(p_members(i).Userindex).Stats.Exp = MAXEXP
            Call CheckUserLevel(p_members(i).Userindex)
            Call WriteUpdateUserStats(p_members(i).Userindex)
        End If
        Call MandarMensajeAConsola(UserList(p_members(i).Userindex).name & " abandona la party.", "Servidor")
        'TODO: Revisar que esto este bien, y no este faltando/sobrando un mensaje, ahora solo los estoy corrigiendo
        Call WriteConsoleMsg(p_members(i).Userindex, " Durante la misma has conseguido " & CStr(Fix(p_members(i).Experiencia)) & " puntos de experiencia!", FontTypeNames.FONTTYPE_PARTY)
        p_CantMiembros = p_CantMiembros - 1
        p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
        p_members(i).Userindex = 0
        p_members(i).Experiencia = 0
        CompactMemberList
    End If
End If

End Function

Public Function HacerLeader(ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 There is no level prohibition
'***************************************************
Dim i As Integer
Dim OldLeader As Integer
Dim oldExp As Double
Dim UserIndexIndex As Integer

UserIndexIndex = 0
HacerLeader = True

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).Userindex > 0 Then
        If p_members(i).Userindex = Userindex Then
            UserIndexIndex = i
        End If
    End If
Next i

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(" Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!", FontTypeNames.FONTTYPE_GUILD))
    HacerLeader = False
    Exit Function
End If


'aca esta todo bien y doy vuelta las collections
OldLeader = p_members(1).Userindex
oldExp = p_members(1).Experiencia

p_members(1).Userindex = p_members(UserIndexIndex).Userindex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

p_members(UserIndexIndex).Userindex = OldLeader
p_members(UserIndexIndex).Experiencia = oldExp

p_Fundador = p_members(1).Userindex

'no need to compact
End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList() As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number (Tavo)
' - 08/18/08 Now TotalExperience is fixed (MarKoxX)
' - 11/27/09 Rehice la funci�n, ahora devuelve el array con los UI online (Budi)
'***************************************************

Dim i As Integer

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).Userindex > 0 Then
        MemberList(i) = p_members(i).Userindex
    End If
Next i

End Sub

Public Function ObtenerExperienciaTotal() As Long
'***************************************************
'Author: Budi
'Last Modification: 11/27/09
'Retrieves the total experience acumulated in the party
'***************************************************
ObtenerExperienciaTotal = p_expTotal
End Function

Public Function PuedeEntrar(ByVal Userindex As Integer, ByRef razon As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 There is no level prohibition
'***************************************************
'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
Dim esArmada As Boolean
Dim esCaos As Boolean
Dim MyLevel As Integer
Dim i As Integer
Dim rv As Boolean
Dim UI As Integer

rv = True
esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)
esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)
MyLevel = UserList(Userindex).Stats.ELV

rv = Distancia(UserList(p_members(1).Userindex).Pos, UserList(Userindex).Pos) <= MAXDISTANCIAINGRESOPARTY
If rv Then
    rv = (p_members(PARTY_MAXMEMBERS).Userindex = 0)
    If rv Then
        For i = 1 To PARTY_MAXMEMBERS
            UI = p_members(i).Userindex
            'pongo los casos que evitarian que pueda entrar
            'aspirante armada en party crimi
            If UI > 0 Then
                If esArmada And criminal(UI) Then
                    razon = "Los miembros de la armada no entran a una party con criminales."
                    rv = False
                End If
                'aspirante caos en party ciuda
                If esCaos And Not criminal(UI) Then
                    razon = "Los miembros de la legi�n oscura no entran a una party con ciudadanos."
                    rv = False
                End If
                'aspirante crimi en party armada
                If UserList(UI).Faccion.ArmadaReal = 1 And criminal(Userindex) Then
                    razon = "Los criminales no entran a partys con miembros de la armada real."
                    rv = False
                End If
                'aspirante ciuda en party caos
                If UserList(UI).Faccion.FuerzasCaos = 1 And Not criminal(Userindex) Then
                    razon = "Los miembros de la legi�n oscura no entran a una party con ciudadanos."
                    rv = False
                End If
            
                If Not rv Then Exit For 'violate una programacion estructurada
            End If
        Next i
    Else
        razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
    End If
Else
    razon = "Te encuentras muy lejos del fundador."
End If

PuedeEntrar = rv

End Function


Public Sub FlushExperiencia()
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
'***************************************************
'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
'acumulada a los usuarios.

Dim i As Integer
If Not PARTY_EXPERIENCIAPORGOLPE Then   'esto sirve SOLO cuando acumulamos la experiencia!
    For i = 1 To PARTY_MAXMEMBERS
        If p_members(i).Userindex > 0 Then
            If p_members(i).Experiencia > 0 Then
                UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
                If UserList(p_members(i).Userindex).Stats.Exp > MAXEXP Then _
                    UserList(p_members(i).Userindex).Stats.Exp = MAXEXP
                Call CheckUserLevel(p_members(i).Userindex)
            Else
                If Abs(UserList(p_members(i).Userindex).Stats.Exp) > Abs(Fix(p_members(i).Experiencia)) Then
                    UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
                Else
                    UserList(p_members(i).Userindex).Stats.Exp = 0
                End If
            End If
            p_members(i).Experiencia = 0
            Call WriteUpdateUserStats(p_members(i).Userindex)
        End If
    Next i
End If

End Sub

Private Sub CompactMemberList()
Dim i As Integer
Dim freeIndex As Integer
i = 1
While i <= PARTY_MAXMEMBERS
    If p_members(i).Userindex = 0 And freeIndex = 0 Then
        freeIndex = i
    ElseIf p_members(i).Userindex > 0 And freeIndex > 0 Then
        p_members(freeIndex).Experiencia = p_members(i).Experiencia
        p_members(freeIndex).Userindex = p_members(i).Userindex
        p_members(i).Userindex = 0
        p_members(i).Experiencia = 0
        'muevo el de la pos i a freeindex
        i = freeIndex
        freeIndex = 0
    End If
    i = i + 1
Wend

End Sub

Public Function CantMiembros() As Integer
CantMiembros = p_CantMiembros
End Function




