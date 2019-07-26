Attribute VB_Name = "NPCs"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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


'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
         UserList(UserIndex).MascotasIndex(i) = 0
         UserList(UserIndex).MascotasType(i) = 0
         
         UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
         Exit For
      End If
    Next i
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        '********************************************************
        'Author: Unknown
        'Llamado cuando la vida de un NPC llega a cero.
        'Last Modify Date: 24/01/2007
        '22/06/06: (Nacho) Chequeamos si es pretoriano
        '24/01/2007: Pablo (ToxicWaste): Agrego para actualizaci�n de tag si cambia de status.
        '********************************************************
     
        '<EhHeader>
        On Error GoTo MuereNpc_Err
 
        '</EhHeader>
        Dim MiNPC As npc
100     MiNPC = Npclist(NpcIndex)
        Dim EraCriminal  As Boolean
        Dim IsPretoriano As Boolean
102     If (esPretoriano(NpcIndex) = 4) Then
            'Solo nos importa si fue matado en el mapa pretoriano.
104         IsPretoriano = True
 
106         If Npclist(NpcIndex).Pos.map = MAPA_PRETORIANO Then
                'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
                Dim i    As Integer
                Dim j    As Integer
                Dim NPCI As Integer
     
108             For i = 8 To 90
110                 For j = 8 To 90
             
112                     NPCI = MapData(Npclist(NpcIndex).Pos.map, i, j).NpcIndex
 
114                     If NPCI > 0 Then
116                         If esPretoriano(NPCI) > 0 And NPCI <> NpcIndex Then
118                             If Npclist(NpcIndex).Pos.X > 50 Then
120                                 If Npclist(NPCI).Pos.X > 50 Then Npclist(NPCI).Invent.ArmourEqpSlot = 1
                                Else
 
122                                 If Npclist(NPCI).Pos.X <= 50 Then Npclist(NPCI).Invent.ArmourEqpSlot = 5
 
                                End If
 
                            End If
 
                        End If
 
124                 Next j
126             Next i
 
128             Call CrearClanPretoriano(Npclist(NpcIndex).Pos.X)
 
            End If
 
130     ElseIf esPretoriano(NpcIndex) > 0 Then
132         IsPretoriano = True
 
134         If Npclist(NpcIndex).Pos.map = MAPA_PRETORIANO Then
136             Npclist(NpcIndex).Invent.ArmourEqpSlot = 0
138             pretorianosVivos = pretorianosVivos - 1
 
            End If
 
        End If
        'Quitamos el npc
140     Call QuitarNPC(NpcIndex)
 
142     If UserIndex > 0 Then ' Lo mato un usuario?
 
144         With UserList(UserIndex)
     
146             If MiNPC.flags.Snd3 > 0 Then
148                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, _
                            MiNPC.Pos.X, MiNPC.Pos.Y))
 
                End If
 
150             .flags.TargetNPC = 0
152             .flags.TargetNpcTipo = eNPCType.Comun
         
                'El user que lo mato tiene mascotas?
154             If .NroMascotas > 0 Then
                    Dim t As Integer
 
156                 For t = 1 To MAXMASCOTAS
 
158                     If .MascotasIndex(t) > 0 Then
160                         If Npclist(.MascotasIndex(t)).TargetNPC = NpcIndex Then
162                             Call FollowAmo(.MascotasIndex(t))
 
                            End If
 
                        End If
 
164                 Next t
 
                End If
         
                '[KEVIN]
 If MiNPC.flags.ExpCount > 0 Then
                If .PartyIndex > 0 Then
                    Call mdParty.ObtenerExito(UserIndex, MiNPC.flags.ExpCount, MiNPC.Pos.map, MiNPC.Pos.X, MiNPC.Pos.Y)
                Else
                    .Stats.Exp = .Stats.Exp + MiNPC.flags.ExpCount
                    If .Stats.Exp > MAXEXP Then _
                        .Stats.Exp = MAXEXP
                    Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
                     If UserList(UserIndex).flags.Oro = 1 Then
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + (MiNPC.flags.ExpCount * 0.4)
                 WriteConsoleMsg UserIndex, "Aumento de exp 40%> Has ganado " & (MiNPC.flags.ExpCount * 0.4) & " puntos de experiencia.", FontTypeNames.fonttype_dios
        End If
                End If
                MiNPC.flags.ExpCount = 0
            End If
         
                '[/KEVIN]
180             Call WriteConsoleMsg(UserIndex, "�Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
 
182             If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1
         
184             EraCriminal = criminal(UserIndex)
         
186             If MiNPC.Stats.Alineacion = 0 Then
         
188                 If MiNPC.Numero = Guardias Then
190                     .Reputacion.NobleRep = 0
192                     .Reputacion.PlebeRep = 0
194                     .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + 500
 
196                     If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
 
                    End If
             
198                 If MiNPC.MaestroUser = 0 Then
200                     .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + vlASESINO
 
202                     If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
 
                    End If
 
204             ElseIf MiNPC.Stats.Alineacion = 1 Then
206                 .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR
 
208                 If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
                 
210             ElseIf MiNPC.Stats.Alineacion = 2 Then
212                 .Reputacion.NobleRep = .Reputacion.NobleRep + vlASESINO / 2
 
214                 If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
                 
216             ElseIf MiNPC.Stats.Alineacion = 4 Then
218                 .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR
 
220                 If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
                 
                End If
         
222             If criminal(UserIndex) And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
224             If Not criminal(UserIndex) And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
         
226             If EraCriminal And Not criminal(UserIndex) Then
228                 Call RefreshCharStatus(UserIndex)
230             ElseIf Not EraCriminal And criminal(UserIndex) Then
232                 Call RefreshCharStatus(UserIndex)
 
                End If
         
234             Call CheckUserLevel(UserIndex)
         
            End With
            
            For i = 1 To MAXUSERQUESTS
 
246         With UserList(UserIndex).QuestStats.Quests(i)
 
248             If .QuestIndex Then
250                 If QuestList(.QuestIndex).RequiredNPCs Then
 
252                     For j = 1 To QuestList(.QuestIndex).RequiredNPCs
 
254                         If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
256                             If QuestList(.QuestIndex).RequiredNPC(j).Amount > .NPCsKilled(j) Then
258                                 .NPCsKilled(j) = .NPCsKilled(j) + 1
 
                                End If
 
                            End If
 
260                     Next j
 
                    End If
 
                End If
 
            End With
 
262     Next i
 
        End If ' Userindex > 0
        
236     If MiNPC.MaestroUser = 0 Then
            'Tiramos el oro
238         Call NPCTirarOro(MiNPC)
            'Tiramos el inventario
240         Call NPC_TIRAR_ITEMS(MiNPC, IsPretoriano)
            'ReSpawn o no
242         Call ReSpawnNpc(MiNPC)
 
        End If
        
244

        '<EhFooter>
        Exit Sub
 
MuereNpc_Err:
        LogError Err.description & vbCrLf & "MuereNpc " & "at line " & Erl
 
        '</EhFooter>
End Sub
Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
    End With
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
    End With
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim j As Long
    
    With Npclist(NpcIndex)
        For j = 1 To .NroCriaturas
            .Criaturas(j).NpcIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0
    End With
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim j As Long
    
    With Npclist(NpcIndex)
        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0
    End With
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        .Attackable = 0
        .CanAttack = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        .QuestNumber = 0
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        If .Owner > 0 Then Call PerdioNpc(.Owner)
        
        .MaestroUser = 0
        .MaestroNpc = 0
        
        .Mascotas = 0
        .Movement = 0
        .name = vbNullString
        .NPCtype = 0
        .Numero = 0
        .Orig.map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .desc = vbNullString
        
        
        Dim j As Long
        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j
    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Now npcs lose their owner
'***************************************************
On Error GoTo errhandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False
        
        .Owner = 0 ' Murio, no necesita mas due�os :P.
        
        If InMapBounds(.Pos.map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)
        End If
    End With
        
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If
Exit Sub

errhandler:
    Call LogError("Error en QuitarNPC")
End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 18/11/2009
'Kills a pet
'***************************************************
On Error GoTo errhandler

    Dim i As Integer
    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        ' Busco el indice de la mascota
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) = NpcIndex Then PetIndex = i
        Next i
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        ' Limpio el slot de la mascota
        .NroMascotas = .NroMascotas - 1
        .MascotasIndex(PetIndex) = 0
        .MascotasType(PetIndex) = 0
        
        ' Elimino la mascota
        Call QuitarNPC(NpcIndex)
    End With
    
    Exit Sub

errhandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    
    If LegalPos(Pos.map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = _
        MapData(Pos.map, Pos.X, Pos.Y).trigger <> 3 And _
        MapData(Pos.map, Pos.X, Pos.Y).trigger <> 2 And _
        MapData(Pos.map, Pos.X, Pos.Y).trigger <> 1
    End If
    
End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Sub
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.map, OrigPos.X, OrigPos.Y) Then
        
        map = OrigPos.map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.map = mapa 'mapa
        altpos.map = mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 And newpos.Y <> 0 Then
                altpos.X = newpos.X
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si ten�a que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)
                If newpos.X <> 0 And newpos.Y <> 0 Then
                    altpos.X = newpos.X
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
                End If
            End If
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.map, newpos.X, newpos.Y, PuedeAgua) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.map = newpos.map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
                
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    map = altpos.map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).Pos.map = map
                    Npclist(nIndex).Pos.X = X
                    Npclist(nIndex).Pos.Y = Y
                    Call MakeNPCChar(True, map, nIndex, map, X, Y)
                    Exit Sub
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.map = newpos.map
                        Npclist(nIndex).Pos.X = newpos.X
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(True, newpos.map, nIndex, newpos.map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
            
        'asignamos las nuevas coordenas
        map = newpos.map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
    End If
            
    'Crea el NPC
    Call MakeNPCChar(True, map, nIndex, map, X, Y)
            
End Sub

Public Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
 
    Dim CharIndex As Integer
    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
  
    Dim GG              As String
    Dim showName    As Byte
    showName = val(GetVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "ShowName"))
 
    GG = IIf(showName > 0, Npclist(NpcIndex).name, vbNullString)
 
    MapData(map, X, Y).NpcIndex = NpcIndex
    With Npclist(NpcIndex)
    If Not toMap Then
        Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.Heading, Npclist(NpcIndex).Char.CharIndex, X, Y, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim, GG, 0, 0)
        Call FlushBuffer(sndIndex)
    Else
        Call AgregarNpc(NpcIndex)
    End If
    End With
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If NpcIndex > 0 Then
        With Npclist(NpcIndex).Char
            .body = body
            .Head = Head
            .Heading = Heading
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(body, Head, Heading, .CharIndex, 0, 0, 0, 0, 0))
        End With
    End If
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los clientes
Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex))

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Public Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/04/2009
'06/04/2009: ZaMa - Now npcs can force to change position with dead character
'01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
'***************************************************

On Error GoTo errh
    Dim nPos As WorldPos
    Dim UserIndex As Integer
    
    With Npclist(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.map, nPos.X, nPos.Y) Then Exit Sub
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.map, nPos.X, nPos.Y) Then Exit Sub
            
            UserIndex = MapData(.Pos.map, nPos.X, nPos.Y).UserIndex
            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.map, nPos.X, nPos.Y) And Not HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then Exit Sub
                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.map, nPos.X, nPos.Y) And HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then Exit Sub
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(NpcIndex).Pos.X
                    .Pos.Y = Npclist(NpcIndex).Pos.Y
                    MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))
                End With
            End If
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            MapData(.Pos.map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        ElseIf .MaestroUser = 0 Then
            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0
            End If
        End If
    End With
Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)
End Sub

Function NextOpenNPC() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errhandler
    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC
Exit Function

errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 10/07/2010
'10/07/2010: ZaMa - Now npcs can't poison dead users.
'***************************************************
 
    Dim N As Integer
   
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then Exit Sub
       
        N = RandomNumber(1, 100)
        If N < 30 Then
            .flags.Envenenado = 1
            Call WriteConsoleMsg(UserIndex, "��La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
   
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/15/2008
'23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
'06/15/2008 -> Optimiz� el codigo. (NicoNZ)
'***************************************************
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim map As Integer
Dim X As Integer
Dim Y As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

PuedeAgua = Npclist(nIndex).flags.AguaValida
PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
Call ClosestLegalPos(Pos, altpos, PuedeAgua)
'Si X e Y son iguales a 0 significa que no se encontro posicion valida

If newpos.X <> 0 And newpos.Y <> 0 Then
    'Asignamos las nuevas coordenas solo si son validas
    Npclist(nIndex).Pos.map = newpos.map
    Npclist(nIndex).Pos.X = newpos.X
    Npclist(nIndex).Pos.Y = newpos.Y
    PosicionValida = True
Else
    If altpos.X <> 0 And altpos.Y <> 0 Then
        Npclist(nIndex).Pos.map = altpos.map
        Npclist(nIndex).Pos.X = altpos.X
        Npclist(nIndex).Pos.Y = altpos.Y
        PosicionValida = True
    Else
        PosicionValida = False
    End If
End If

If Not PosicionValida Then
    Call QuitarNPC(nIndex)
    SpawnNpc = 0
    Exit Function
End If

'asignamos las nuevas coordenas
map = newpos.map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y

'Crea el NPC
Call MakeNPCChar(True, map, nIndex, map, X, Y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
End If

SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.map, MiNPC.Orig)

End Sub

Private Sub NPCTirarOro(ByRef MiNPC As npc)

    If MiNPC.GiveGLD > 0 Then
        Dim MiObj As Obj
        Dim MiAux As Long
        MiAux = MiNPC.GiveGLD * Oroc
        Do While MiAux > MAX_INVENTORY_OBJS
            MiObj.Amount = MAX_INVENTORY_OBJS
            MiObj.objindex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            MiAux = MiAux - MAX_INVENTORY_OBJS
        Loop
        If MiAux > 0 Then
            MiObj.Amount = MiAux
            MiObj.objindex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
        End If
    End If
End Sub
Public Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ���� NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendr� que ver
'conmigo. Para leer los NPCS se deber� usar la
'nueva clase clsinimanager.
'
'Alejo
'
'###################################################
    Dim NpcIndex As Integer
    Dim Leer As clsIniManager
    Dim LoopC As Long
    Dim ln As String
    Dim aux As String
    
    Set Leer = LeerNPCs
    
    'If requested index is invalid, abort
    If Not Leer.KeyExists("NPC" & NpcNumber) Then
        OpenNPC = MAXNPCS + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNPC
    
    If NpcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NpcIndex
        Exit Function
    End If
    
    With Npclist(NpcIndex)
        .Numero = NpcNumber
        .name = Leer.GetValue("NPC" & NpcNumber, "Name")
        .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        
        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement
        
        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
        .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "EscudoAnim"))
        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "ArmaAnim"))
        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
        .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
        
        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
        .flags.OldHostil = .Hostile
        
        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * Expc
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
        
        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
        .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
        
        .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
        
        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
        With .Stats
            .MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
            .MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
            .MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
            .Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
        End With
        
        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
        For LoopC = 1 To .Invent.NroItems
            ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
            .Invent.Object(LoopC).objindex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        Next LoopC
        
        For LoopC = 1 To MAX_NPC_DROPS
            ln = Leer.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
            .Drop(LoopC).objindex = val(ReadField(1, ln, 45))
            .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
        Next LoopC

        
        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC
        
        If .NPCtype = eNPCType.Entrenador Then
            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC
        End If
        
        With .flags
            .NPCActive = True
            
            If Respawn Then
                .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
                .Respawn = 1
            End If
            
            .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
            
            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))
        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
        .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
    End With
    
    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0
        End If
    End With
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0
    End With
End Sub

Public Sub ValidarPermanenciaNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'Chequea si el npc continua perteneciendo a alg�n usuario
'***************************************************

    With Npclist(NpcIndex)
        If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)
    End With
End Sub