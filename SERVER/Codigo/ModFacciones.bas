Attribute VB_Name = "ModFacciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public GLD As Long
Public ArmaduraImperial1 As Integer
Public ArmaduraImperial2 As Integer
Public ArmaduraImperial3 As Integer
Public TunicaMagoImperial As Integer
Public TunicaMagoImperialEnanos As Integer
Public ArmaduraCaos1 As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer



Public Const NUM_RANGOS_FACCION As Integer = 5
Private Const NUM_DEF_FACCION_ARMOURS As Byte = 3

Public Enum eTipoDefArmors
    ieBaja
    ieMedia
    ieAlta
End Enum

Public Type tFaccionArmaduras
    Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
    Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer
End Type

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION) As Long

Private Function GetArmourAmount(ByVal Rango As Integer, ByVal TipoDef As eTipoDefArmors) As Integer
'***************************************************
'Autor: ZaMa
'Last Modification: 15/04/2010
'Returns the amount of armours to give, depending on the specified rank
'***************************************************

    Select Case TipoDef
        
        Case eTipoDefArmors.ieBaja
            GetArmourAmount = 1
            
        Case eTipoDefArmors.ieMedia
            GetArmourAmount = 1
            
        Case eTipoDefArmors.ieAlta
            GetArmourAmount = 1
            
    End Select
    
End Function

Private Sub GiveFactionArmours(ByVal UserIndex As Integer, ByVal IsCaos As Boolean)
'***************************************************
'Autor: ZaMa
'Last Modification: 15/04/2010
'Gives faction armours to user
'***************************************************
    
    Dim ObjArmour As Obj
    Dim Rango As Integer
    
    With UserList(UserIndex)
    
        Rango = val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1
    
    
        ' Entrego armaduras de defensa baja
        'ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)
        'If IsCaos = True And ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta And eTipoDefArmors.ieMedia And eTipoDefArmors.ieBaja) >= 1 Then Exit Sub
        'If IsCaos = False And ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta And eTipoDefArmors.ieMedia And eTipoDefArmors.ieBaja) >= 1 Then Exit Sub
       ' If IsCaos Then
       '     ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieBaja)
      '  Else
       '     ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieBaja)
       ' End If
        
       ' If Not MeterItemEnInventario(userIndex, ObjArmour) Then
       '     Call TirarItemAlPiso(.Pos, ObjArmour)
       ' End If
        
        
        ' Entrego armaduras de defensa media
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)
        If IsCaos = True Then
      If .Faccion.RecibioArmaduraCaos = 0 Then
 Dim MiObj As Obj
    MiObj.Amount = 1
   
   'CAOS
    Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 734
                Case eClass.Cleric
                    MiObj.objindex = 736
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 738
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 740
                        Case eGenero.Hombre
                            MiObj.objindex = 741
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 735
                Case eClass.Cleric
                    MiObj.objindex = 737
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 739
                Case eClass.Mage
                    MiObj.objindex = 742
        End Select
    End Select
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
   End If
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date
ElseIf IsCaos = False Then
If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    MiObj.Amount = 1
        
    'ARMADA
    Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 779
                Case eClass.Cleric
                    MiObj.objindex = 781
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 783
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 785
                        Case eGenero.Hombre
                            MiObj.objindex = 786
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 780
                Case eClass.Cleric
                    MiObj.objindex = 782
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 784
                Case eClass.Mage
                    MiObj.objindex = 787
        End Select
    End Select
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date
    'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
    UserList(UserIndex).Faccion.MatadosIngreso = UserList(UserIndex).Faccion.CiudadanosMatados

End If

    End With

End Sub

Public Sub GiveExpReward(ByVal UserIndex As Integer, ByVal Rango As Long)
'***************************************************
'Autor: ZaMa
'Last Modification: 15/04/2010
'Gives reward exp to user
'***************************************************
    
    Dim GivenExp As Long
    
    With UserList(UserIndex)
        
        GivenExp = RecompensaFacciones(Rango)
        
        .Stats.Exp = .Stats.Exp + GivenExp
        
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        Call WriteConsoleMsg(UserIndex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

        Call CheckUserLevel(UserIndex)
        
    End With
    
End Sub

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/04/2010
'Handles the entrance of users to the "Armada Real"
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
'27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
'15/04/2010: ZaMa - Cambio en recompensas iniciales.
'***************************************************

With UserList(UserIndex)
    If .Faccion.ArmadaReal = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.FuerzasCaos = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! Vete de aquí seguidor de las sombras.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If criminal(UserIndex) Then
        Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejército real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.CriminalesMatados < 50 Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 50 criminales, sólo has matado " & .Faccion.CriminalesMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Stats.ELV < 25 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
     
    If .Faccion.CiudadanosMatados > 0 Then
        Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.Reenlistadas > 4 Then
        Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Reputacion.NobleRep < 0 Then
        Call WriteChatOverHead(UserIndex, "Necesitas ser aún más noble para integrar el ejército real, sólo tienes " & .Reputacion.NobleRep & "/20.000 puntos de nobleza", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .GuildIndex > 0 Then
        If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Perteneces a un clan neutro, sal de él si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
    End If
    
    .Faccion.ArmadaReal = 1
    .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al ejército real!!! Aquí tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargaré de recompensarte.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Rey de Banderbill> Ahora le ofreceré estas vestimentas a " & .name & " por haberse enlistado a la Armada Real. Espero grandes logros de este noble guerrero.", FontTypeNames.FONTTYPE_CONSEJOVesA))
    
    ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
    If .Faccion.RecibioArmaduraReal = 0 Then
        
       
    Dim LiObj As Obj
    LiObj.Amount = 1
    
'[Wizard 03/09/05] no se quien hizo lo que estaba aca, pero por dios mandenlo a un curso de redaccion
'Habia 3 cases diciendo lo mismo, 1 If clause que nunca se accedia por suerte porque si se accedia daba armadura del caos
'ademas usan los Ucase$ para esto, que son cosas que los escribe el codigo y no pueden cambiar, gastan memoria ram al pedo.
Select Case .Raza
    Case Drow, Elfo, Humano
        If .clase = Cleric Or .clase = Druid Or .clase = Bard Then
            LiObj.objindex = 372
        ElseIf .Genero = Hombre And .clase = Mage Then
            LiObj.objindex = 517
        ElseIf .Genero = Mujer And .clase = Mage Then
            LiObj.objindex = 516
        ElseIf (.Genero = Mujer) And (.clase = Paladin Or .clase = Warrior Or .clase = Assasin Or .clase = Hunter) Then
            LiObj.objindex = 520
        ElseIf (.Genero = Hombre) And (.clase = Paladin Or .clase = Warrior Or .clase = Assasin Or .clase = Hunter) Then
            LiObj.objindex = 521
        End If
    
    Case Gnomo, Enano
        If .clase = Warrior Or .clase = Paladin Or .clase = Hunter Or .clase = Assasin Then
            LiObj.objindex = 492
        ElseIf .clase = Mage Or .clase = Bard Or .clase = Druid Or .clase = Cleric Then
            LiObj.objindex = 549
        Else 'Trabajadoras
            LiObj.objindex = 678
        End If
End Select
        
        If Not MeterItemEnInventario(UserIndex, LiObj) Then
            Call TirarItemAlPiso(.Pos, LiObj)
        End If
 .Faccion.RecibioArmaduraReal = 1
 
        Call GiveExpReward(UserIndex, 0)
        
        .Faccion.RecibioArmaduraReal = 1
        .Faccion.NivelIngreso = .Stats.ELV
        .Faccion.FechaIngreso = Date
        'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
        .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados
        
        .Faccion.RecibioExpInicialReal = 1
        .Faccion.RecompensasReal = 0
        .Faccion.NextRecompensa = 100
        
    End If
    
    If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
    
    Call LogEjercitoReal(.name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
End With

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Armada Real"
'***************************************************
Dim Crimis As Long
Dim Lvl As Byte
Dim NextRecom As Long
Dim Nobleza As Long
Dim MiObj As Obj
MiObj.Amount = 1
Lvl = UserList(UserIndex).Stats.ELV
Crimis = UserList(UserIndex).Faccion.CriminalesMatados
NextRecom = UserList(UserIndex).Faccion.NextRecompensa
Nobleza = UserList(UserIndex).Reputacion.NobleRep

If Crimis < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Ya has recibido tu recompensa, mata " & NextRecom - Crimis & " criminaales mas para recibir la proxima!!!", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
        Case 100:
         If Lvl < 32 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 32 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            If (UserList(UserIndex).Stats.GLD >= 500000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 500000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasReal = 1
            UserList(UserIndex).Faccion.NextRecompensa = 175
            'Call PerderItemsFaccionarios(Userindex)
            '2doARMY
    Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 788
                Case eClass.Cleric
                    MiObj.objindex = 790
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 792
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 794
                        Case eGenero.Hombre
                            MiObj.objindex = 795
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 789
                Case eClass.Cleric
                    MiObj.objindex = 791
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 793
                Case eClass.Mage
                    MiObj.objindex = 796
        End Select
    End Select
        
        Case 175:
         If Lvl < 36 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            If (UserList(UserIndex).Stats.GLD >= 1500000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1500000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 1.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasReal = 2
            UserList(UserIndex).Faccion.NextRecompensa = 250
            'Call PerderItemsFaccionarios(Userindex)
            '3cerARMY
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 797
                Case eClass.Cleric
                    MiObj.objindex = 799
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 801
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 803
                        Case eGenero.Hombre
                            MiObj.objindex = 804
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 798
                Case eClass.Cleric
                    MiObj.objindex = 800
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 802
                Case eClass.Mage
                    MiObj.objindex = 805
        End Select
    End Select
        
        Case 250:
         If Lvl < 38 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 38 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            If (UserList(UserIndex).Stats.GLD >= 2500000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 2500000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 2.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasReal = 3
            UserList(UserIndex).Faccion.NextRecompensa = 325
            'Call PerderItemsFaccionarios(Userindex)
            
'4toARMY
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 806
                Case eClass.Cleric
                    MiObj.objindex = 808
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 810
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 812
                        Case eGenero.Hombre
                            MiObj.objindex = 813
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 807
                Case eClass.Cleric
                    MiObj.objindex = 809
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 811
                Case eClass.Mage
                    MiObj.objindex = 814
        End Select
    End Select
        
        Case 325:
         If Lvl < 40 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 40 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            UserList(UserIndex).Faccion.RecompensasReal = 4
            UserList(UserIndex).Faccion.NextRecompensa = 415
           ' Call PerderItemsFaccionarios(Userindex)
            If (UserList(UserIndex).Stats.GLD >= 4000000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 4000000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 4.000.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            '5toARMY
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 815
                Case eClass.Cleric
                    MiObj.objindex = 817
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 819
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 821
                        Case eGenero.Hombre
                            MiObj.objindex = 822
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 816
                Case eClass.Cleric
                    MiObj.objindex = 818
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 820
                Case eClass.Mage
                    MiObj.objindex = 823
        End Select
    End Select
            '.Faccion.NextRecompensa = 17001
      
        
        Case 415:
            Exit Sub
    End Select
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
   End If
 
Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + TituloReal(UserIndex) + "!!!", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
'If UserList(UserIndex).Stats.Exp > MAXEXP Then
'    UserList(UserIndex).Stats.Exp = MAXEXP
'End If
'Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpX100 & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

'Call CheckUserLevel(UserIndex)


End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

With UserList(UserIndex)
    .Faccion.ArmadaReal = 0
    Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado del ejército real!!!", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado del ejército real!!!", FontTypeNames.FONTTYPE_FIGHT)
        .Faccion.RecibioArmaduraReal = 0
    End If
    
    If .Invent.ArmourEqpObjIndex <> 0 Then
        'Desequipamos la armadura real si está equipada
        If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpObjIndex)
   'Call QuitarObjetos(.Invent.ArmourEqpObjIndex, 1, UserIndex)
    End If
    
    If .Invent.EscudoEqpObjIndex <> 0 Then
        'Desequipamos el escudo de caos si está equipado
        If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpObjIndex)
    'Call QuitarObjetos(ObjData(.Invent.EscudoEqpObjIndex).Real = 1, 1, UserIndex)
    End If
    
    If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End With

End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

With UserList(UserIndex)
    .Faccion.FuerzasCaos = 0
    Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If .Invent.ArmourEqpObjIndex <> 0 Then
        'Desequipamos la armadura de caos si está equipada
        If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
   ' Call QuitarObjetos(.Invent.ArmourEqpObjIndex, 1, UserIndex)
    End If
    
    If .Invent.EscudoEqpObjIndex <> 0 Then
        'Desequipamos el escudo de caos si está equipado
        If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpObjIndex)
    'Call QuitarObjetos(.Invent.EscudoEqpObjIndex, 1, UserIndex)
    End If
    
    
    
    
    If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End With

End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Armada Real"
'***************************************************

Select Case UserList(UserIndex).Faccion.RecompensasReal
   
    Case 0
        TituloReal = "Aprendiz"
    Case 1
        TituloReal = "Caballero"
    Case 2
        TituloReal = "Capitán"
    Case 3
        TituloReal = "Guardián"
    Case Else
        TituloReal = "Campeón de la Luz"
End Select


End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 27/11/2009
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
'27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
'Handles the entrance of users to the "Legión Oscura"
'***************************************************

With UserList(UserIndex)
    If Not criminal(UserIndex) Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Lárgate de aquí, bufón!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.FuerzasCaos = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Faccion.ArmadaReal = 1 Then
        Call WriteChatOverHead(UserIndex, "Las sombras reinarán en las tierras Desterianas. ¡¡¡Fuera de aquí insecto real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
    If .Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
        Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    '[/Barrin]
    
    If Not criminal(UserIndex) Then
        Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tú no eres bienvenido aquí asqueroso ciudadano.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
   If .Faccion.CiudadanosMatados < 60 Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 60 ciudadanos, sólo has matado " & .Faccion.CiudadanosMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .Stats.ELV < 25 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If .GuildIndex > 0 Then
        If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Perteneces a un clan neutro, sal de él si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
    End If
    
    
    If .Faccion.Reenlistadas > 4 Then
        If .Faccion.Reenlistadas = 200 Then
            Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        End If
        Exit Sub
    End If
    
    .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
    .Faccion.FuerzasCaos = 1
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aquí tienes tus armaduras. Derrama sangre ciudadana y real, y serás recompensado, lo prometo.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
     'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Señor del Miedo> Es un gusto anunciar que " & .name & " se ha enlistado en la Legión Oscura. Su alma me pertenece y su equipamiento es la recompensa para sembrar el miedo en estas tierras.", FontTypeNames.FONTTYPE_CONSEJOCAOSVesA))
    
    If .Faccion.RecibioArmaduraCaos = 0 Then
                
        Call GiveFactionArmours(UserIndex, True)
        Call GiveExpReward(UserIndex, 0)
        
        .Faccion.RecibioArmaduraCaos = 1
        .Faccion.NivelIngreso = .Stats.ELV
        .Faccion.FechaIngreso = Date
    
        .Faccion.RecibioExpInicialCaos = 1
        .Faccion.RecompensasCaos = 0
        .Faccion.NextRecompensa = 110
    End If
    
    If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

    Call LogEjercitoCaos(.name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
End With

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Legión Oscura"
'***************************************************
Dim Ciudas As Long
Dim Lvl As Byte
Dim NextRecom As Long
Dim MiObj As Obj
MiObj.Amount = 1
Lvl = UserList(UserIndex).Stats.ELV
Ciudas = UserList(UserIndex).Faccion.CiudadanosMatados
NextRecom = UserList(UserIndex).Faccion.NextRecompensa

If Ciudas < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Ya has recibido tu recompensa, mata " & NextRecom - Ciudas & "  ciudadanos mas para recibir la proxima!!!", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
    Case 110:
        If Lvl < 27 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        If (UserList(UserIndex).Stats.GLD >= 500000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 500000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasCaos = 1
            UserList(UserIndex).Faccion.NextRecompensa = 180
           ' Call PerderItemsFaccionarios(Userindex)
            '2doCAOS
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 743
                Case eClass.Cleric
                    MiObj.objindex = 745
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 747
                Case eClass.Mage
                    If UserList(UserIndex).Genero = 2 Then
                    MiObj.objindex = 749
                    ElseIf UserList(UserIndex).Genero = 1 Then
                    MiObj.objindex = 750
                    End If
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 744
                Case eClass.Cleric
                    MiObj.objindex = 746
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 748
                Case eClass.Mage
                    MiObj.objindex = 751
        End Select
    End Select
        
        Case 180:
        If Lvl < 30 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 30 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            If (UserList(UserIndex).Stats.GLD >= 1500000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1500000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 1.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasCaos = 2
            UserList(UserIndex).Faccion.NextRecompensa = 270
          '  Call PerderItemsFaccionarios(Userindex)
            '3cerCAOS
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 752
                Case eClass.Cleric
                    MiObj.objindex = 754
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 756
                Case eClass.Mage
                    If UserList(UserIndex).Genero = 1 Then
                    MiObj.objindex = 758
                    ElseIf UserList(UserIndex).Genero = 2 Then
                    MiObj.objindex = 759
                    End If
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 753
                Case eClass.Cleric
                    MiObj.objindex = 755
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 757
                Case eClass.Mage
                    MiObj.objindex = 760
        End Select
    End Select
        
        Case 270:
        If Lvl < 34 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 34 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            If (UserList(UserIndex).Stats.GLD >= 2500000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 2500000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 2.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasCaos = 3
            UserList(UserIndex).Faccion.NextRecompensa = 350
          ' Call PerderItemsFaccionarios(Userindex)
            '4toCAOS
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 761
                Case eClass.Cleric
                    MiObj.objindex = 763
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 765
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 767
                        Case eGenero.Hombre
                            MiObj.objindex = 768
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 762
                Case eClass.Cleric
                    MiObj.objindex = 764
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 766
                Case eClass.Mage
                    MiObj.objindex = 769
        End Select
    End Select
            
        
        Case 350:
         If Lvl < 37 Then
                Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
            If (UserList(UserIndex).Stats.GLD >= 4000000) Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 4000000
        Call WriteUpdateGold(UserIndex)
        Else
        Call WriteChatOverHead(UserIndex, "Necesitas 4.000.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
        End If
            UserList(UserIndex).Faccion.RecompensasCaos = 4
            UserList(UserIndex).Faccion.NextRecompensa = 425
            'Call PerderItemsFaccionarios(Userindex)
            '5toCAOS
Select Case UserList(UserIndex).Raza
        Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 770
                Case eClass.Cleric
                    MiObj.objindex = 772
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 774
                Case eClass.Mage
                    Select Case UserList(UserIndex).Genero
                        Case eGenero.Mujer
                            MiObj.objindex = 776
                        Case eGenero.Hombre
                            MiObj.objindex = 777
                    End Select
              End Select
        Case eRaza.Gnomo, eRaza.Enano
            Select Case UserList(UserIndex).clase
                Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
                    MiObj.objindex = 771
                Case eClass.Cleric
                    MiObj.objindex = 773
                Case eClass.Paladin, eClass.Warrior
                    MiObj.objindex = 775
                Case eClass.Mage
                    MiObj.objindex = 778
        End Select
    End Select
Case 425:
        WriteChatOverHead UserIndex, "¡¡¡Bien hecho, ya no tengo más recompensas para ti. Sigue así!!!", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite
            Exit Sub
    End Select
If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
   End If

Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + TituloCaos(UserIndex) + ", aquí tienes tu recompensa!!!", Str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
'If UserList(UserIndex).Stats.Exp > MAXEXP Then
'    UserList(UserIndex).Stats.Exp = MAXEXP
'End If
'Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpX100 & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
'Call CheckUserLevel(UserIndex)


End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Legión Oscura"
'***************************************************
'Rango 1: Esbirro (20)
'Rango 2: Sanguinario (30 + 100k)
'Rango 3: Condenado (40 + 200k)
'Rango 4: Caballero de la oscuridad (50 + 375k)
'Rango 5: Devorador de almas (100 + 500k)


Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Esbirro"
    Case 1
        TituloCaos = "Sanguinario"
    Case 2
        TituloCaos = "Condenado"
    Case 3
        TituloCaos = "Caballero de la Oscuridad"
    Case 4
        TituloCaos = "Devorador de Almas"
End Select

End Function


Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
Dim i As Byte
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
ItemIndex = UserList(UserIndex).Invent.Object(i).objindex
If ItemIndex > 0 Then
    If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then
    QuitarUserInvItem UserIndex, i, UserList(UserIndex).Invent.Object(i).Amount
    UpdateUserInv False, UserIndex, i
        If ObjData(ItemIndex).OBJType = otarmadura Or ObjData(ItemIndex).OBJType = otescudo Then
        If ObjData(ItemIndex).Real = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
        If ObjData(ItemIndex).Caos = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
  Else
        UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Or UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
  End If
    End If
    
    End If
    
            Next i
End Sub



