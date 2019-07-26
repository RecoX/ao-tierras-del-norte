Attribute VB_Name = "modHungerGames"
Option Explicit
Public Type SurvivalGames
     HungerIndex As Integer 'Esta en los juegos del hambre
     HungerDie As Byte ' Murio en los SG
     'HungerDiePos As WorldPos ' Posicion para llevar dsp
     HungerGold As Long 'Oro que gana
End Type

Public Type HungerSG
     Cuonter As Byte 'Contador inicio
     Drop As Boolean 'Items
     Oro As Long ' Oro para ingresar
     Created As Byte 'Iniciaron los sg?
     Cupos As Byte 'Cuantos cupos
     Joined As Integer 'Los que entraron
     InPie As Byte
End Type
Public Const MAX_COFRES As Byte = 22
Public Hambriento As Integer

Public Const HungerMap As Integer = 192 'Mapa donde se hace
Public SurvivalG As HungerSG

Public Sub SecondSg()
If SurvivalG.Created = 0 Then Exit Sub
Dim i As Long
With SurvivalG
If .Cuonter > 0 Then
SendData SendTarget.toMap, HungerMap, PrepareMessageConsoleMsg("Juegos del Hambre> " & .Cuonter, FontTypeNames.FONTTYPE_Conteos)
.Cuonter = .Cuonter - 1
If .Cuonter <= 0 Then
.Cuonter = 0
SendData SendTarget.toMap, HungerMap, PrepareMessageConsoleMsg("Juegos del Hambre> YA", FontTypeNames.FONTTYPE_Conteos)
For i = 1 To NumUsers
If UserList(i).flags.SG.HungerIndex <> 0 Then
WritePauseToggle i
End If
Next i
End If
End If
End With
End Sub

Public Sub CleanSg()
With SurvivalG
Dim i As Long
For i = 1 To LastUser
If UserList(i).flags.SG.HungerIndex <> 0 Then
UserList(i).flags.SG.HungerIndex = 0
UserList(i).flags.SG.HungerDie = 0
UserList(i).flags.SG.HungerGold = 0
End If
Next i
.Cuonter = 0
.Drop = False
.Oro = 0
.Created = 0
.Joined = 0
.InPie = 0
.Cupos = 0
End With
End Sub

Public Function HungerGamesCanCreate(ByVal Cupos As Byte, ByVal Gold As Long, ByVal Drop As Boolean, ByRef Err As String) As Boolean

With SurvivalG

If .Created <> 0 Then
Err = "Los juegos del hambre ya est�n en curso!"
HungerGamesCanCreate = False
Exit Function
End If

If Cupos <= 0 Then Err = "Los cupos no son v�lidos": HungerGamesCanCreate = False: Exit Function

If Gold <= 0 Then Err = "El oro ingresado no es v�lido": HungerGamesCanCreate = False: Exit Function

HungerGamesCanCreate = True
End With
End Function

Public Sub HungerGamesCreate(ByVal Cupos As Byte, ByVal Gold As Long, ByVal Drop As Boolean)

With SurvivalG
.Oro = Gold
.Drop = Drop
.Cupos = Cupos


SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Juegos del hambre> Han dado inicio los Juegos del Hambre! El m�ximo de cupos es [" & Cupos & "], para entrar solo debes pagar " & Gold & " monedas de oro." & vbNewLine & _
IIf(Drop, "El ganador se queda con los items", "") & "Para ingresar escribe /SURVIVAL", FontTypeNames.fonttype_conteo)

.Created = 1
End With
End Sub

Public Function HungerGamesCanJoin(ByVal UI As Integer, ByRef Err As String) As Boolean

With UserList(UI)

If .death <> 0 Then Err = "Est�s en otro evento!": HungerGamesCanJoin = False: Exit Function
If SurvivalG.Created <> 1 Then Err = "El evento no tiene las inscripciones abiertas o no ha sido iniciado!": HungerGamesCanJoin = False: Exit Function
If .flags.Muerto <> 0 Then Err = "Est�s muerto!": HungerGamesCanJoin = False: Exit Function
If .flags.SG.HungerIndex <> 0 Then Err = "Ya est�s en los juegos del hambre!": HungerGamesCanJoin = False: Exit Function
If .Stats.Gld < SurvivalG.Oro Then Err = "No tienes el oro necesario.": HungerGamesCanJoin = False: Exit Function
If SurvivalG.Joined >= SurvivalG.Cupos Then Err = "Cupos llenos.": HungerGamesCanJoin = False: Exit Function
If .flags.SG.HungerDie <> 0 And SurvivalG.Created = 1 Then Err = "Ya moriste en los juegos del hambre": HungerGamesCanJoin = False: Exit Function
If .Invent.NroItems <> 0 Then Err = "No debes tener ning�n item en tu inventario!": HungerGamesCanJoin = False: Exit Function
HungerGamesCanJoin = True

End With
'HungerGamesJoin UI, SurvivalG.Oro, SurvivalG.Cupos
End Function


Public Sub HungerGamesJoin(ByVal UI As Integer, ByVal Gld As Long, ByVal Cupos As Byte)
With UserList(UI)
SurvivalG.Joined = SurvivalG.Joined + 1
.flags.SG.HungerIndex = UI
.flags.BeforeMap = .Pos.map
.flags.BeforeX = .Pos.X
.flags.BeforeY = .Pos.Y
WarpUserChar UI, HungerMap, RandomNumber(50, 75), 30, True
.Stats.Gld = .Stats.Gld - SurvivalG.Oro
WriteUpdateGold UI
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Juegos del Hambre> Ingresa " & .name & " a los juegos del hambre!", FontTypeNames.fonttype_conteo)
WritePauseToggle UI
If SurvivalG.Joined = SurvivalG.Cupos Then
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Juegos del Hambre> Cupos alcanzados! " & vbNewLine & "Juegos del Hambre> Damos inicio al Evento.", FontTypeNames.fonttype_conteo)
Dim i As Long
For i = 1 To NumUsers
If UserList(i).flags.SG.HungerIndex <> 0 Then
WarpUserChar i, HungerMap, RandomNumber(71, 78), 30, False
End If
Next i
Dim Cof As Obj
Cof.ObjIndex = 11
Cof.Amount = 1
For i = 1 To MAX_COFRES
Dim Xx As Integer
Dim Yy As Integer
Xx = RandomNumber(8, 90)
Yy = RandomNumber(8, 90)
If MapData(HungerMap, Xx, Yy).ObjInfo.ObjIndex = 0 Then
            MakeObj Cof, HungerMap, Xx, Yy
            End If
            Next i

SurvivalG.Cuonter = 5
SurvivalG.InPie = SurvivalG.Joined

End If
'CleanSg
'.flags.SG.HungerIndex = 0
'WritePauseToggle UI
End With

End Sub

Public Sub HungerDesconect(ByVal UI As Integer)
With UserList(UI)
 
If .flags.SG.HungerIndex <> 0 Then
TirarTodosLosItems UI
WarpUserChar UI, .flags.BeforeMap, .flags.BeforeX, .flags.BeforeY, False
SurvivalG.Joined = SurvivalG.Joined - 1
SurvivalG.InPie = SurvivalG.InPie - 1
End If
 
If SurvivalG.InPie = 1 Then
 
Dim H As Long
 
For H = 1 To NumUsers
 
If UserList(H).flags.SG.HungerIndex <> 0 Then
HungerWin H
End If
 
Next H
 
End If
 
End With
End Sub
Public Sub HungerDie(ByVal UI As Integer)
With UserList(UI)
 
If .flags.SG.HungerIndex <> 0 Then
TirarTodosLosItems UI
WarpUserChar UI, .flags.BeforeMap, .flags.BeforeX, .flags.BeforeY, False

.flags.SG.HungerDie = 1
SurvivalG.Joined = SurvivalG.Joined - 1
SurvivalG.InPie = SurvivalG.InPie - 1
End If
End With
 
If SurvivalG.InPie = 1 Then
Dim P As Long
For P = 1 To NumUsers
If UserList(P).flags.SG.HungerIndex <> 0 And Not UserList(P).flags.SG.HungerDie = 1 Then
HungerWin P
End If
Next P
End If
End Sub
Public Sub HungerWin(ByVal Win As Integer)
With UserList(Win)

If .flags.SG.HungerIndex <> 0 Then
Dim Pozo As Long
Pozo = SurvivalG.Oro * SurvivalG.Cupos
.Stats.Gld = .Stats.Gld + Pozo
'If SurvivalG.Drop = True Then
'SurvivalG.counteraregresar = 120
'Else
WarpUserChar Win, UserList(Win).flags.BeforeMap, UserList(Win).flags.BeforeX, UserList(Win).flags.BeforeY, False
CleanSg
If SurvivalG.Drop = False Then
CleanHGMap
End If
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Juegos del Hambre> El ganador ha sido " & .name & "! se lleva el pozo acumulado de " & Pozo & ".", FontTypeNames.fonttype_conteo)

End If
End With
End Sub
Public Function CleanHGMap()
Dim X As Long
Dim Y As Long
For X = 1 To 100
For Y = 1 To 100
With MapData(HungerMap, X, Y).ObjInfo

EraseObj .Amount, HungerMap, X, Y
End With
Next Y
Next X

End Function
'PARA LOS COFRES*********************************************************************
'**********************************************************************************
'**********************************************************************************
Public Function ArmaRandom() As Integer
 
' @ Devuelve un arma random, se puede cambiar..
 
Dim YATengo     As Boolean
Dim YAObjIndex  As Integer
 
Do While Not YATengo
   
    YAObjIndex = RandomNumber(1, NumObjDatas)
   
    'Ya tengo?
    YATengo = (ObjData(YAObjIndex).OBJType = eOBJType.otWeapon)
    If (ObjData(YAObjIndex).MinHIT >= 100) Or (ObjData(YAObjIndex).Premium <> 0) Or (ObjData(YAObjIndex).Real <> 0) Or (ObjData(YAObjIndex).Caos <> 0) Or (ObjData(YAObjIndex).NoSeCae <> 0) Then
    YATengo = False
    End If
Loop
 
ArmaRandom = YAObjIndex
 
End Function

Public Function ArmaduraRandom() As Integer
 
' @ Devuelve una armadira random, se puede cambiar..
 
Dim YATengo     As Boolean
Dim YAObjIndex  As Integer
 
Do While Not YATengo
   
    YAObjIndex = RandomNumber(1, NumObjDatas)
   
    'Ya tengo?
    YATengo = (ObjData(YAObjIndex).OBJType = eOBJType.otarmadura)
    If (ObjData(YAObjIndex).Premium <> 0) Or (ObjData(YAObjIndex).Real <> 0) Or (ObjData(YAObjIndex).Caos <> 0) Or (ObjData(YAObjIndex).NoSeCae <> 0) Or (ObjData(YAObjIndex).MinDef >= 30) Then
    YATengo = False
    End If
Loop
 
ArmaduraRandom = YAObjIndex
 
End Function

Public Function AnilloRandom() As Integer
 
' @ Devuelve una AnilloRandom random, se puede cambiar..
 
Dim YATengo     As Boolean
Dim YAObjIndex  As Integer
 
Do While Not YATengo
   
    YAObjIndex = RandomNumber(1, NumObjDatas)
   
    'Ya tengo?
    YATengo = (ObjData(YAObjIndex).OBJType = eOBJType.otAnillo)
    If (ObjData(YAObjIndex).Premium <> 0) Or (ObjData(YAObjIndex).Real <> 0) Or (ObjData(YAObjIndex).Caos <> 0) Or (ObjData(YAObjIndex).NoSeCae <> 0) Then
    YATengo = False
    End If
Loop
 
AnilloRandom = YAObjIndex
 
End Function

Public Function CascoRandom() As Integer
 
' @ Devuelve una AnilloRandom random, se puede cambiar..
 
Dim YATengo     As Boolean
Dim YAObjIndex  As Integer
 
Do While Not YATengo
   
    YAObjIndex = RandomNumber(1, NumObjDatas)
   
    'Ya tengo?
    YATengo = (ObjData(YAObjIndex).OBJType = eOBJType.otcasco)
    If (ObjData(YAObjIndex).MinDef >= 20) Or (ObjData(YAObjIndex).Premium <> 0) Or (ObjData(YAObjIndex).Real <> 0) Or (ObjData(YAObjIndex).Caos <> 0) Or (ObjData(YAObjIndex).NoSeCae <> 0) Then
    YATengo = False
    End If
Loop
 
CascoRandom = YAObjIndex
 
End Function

Public Function EscudoRandom() As Integer
 
' @ Devuelve una AnilloRandom random, se puede cambiar..
 
Dim YATengo     As Boolean
Dim YAObjIndex  As Integer
 
Do While Not YATengo
   
    YAObjIndex = RandomNumber(1, NumObjDatas)
   
    'Ya tengo?
    YATengo = (ObjData(YAObjIndex).OBJType = eOBJType.otescudo)
    If (ObjData(YAObjIndex).Premium <> 0) Or (ObjData(YAObjIndex).Real <> 0) Or (ObjData(YAObjIndex).Caos <> 0) Or (ObjData(YAObjIndex).NoSeCae <> 0) Then
    YATengo = False
    End If
Loop
 
EscudoRandom = YAObjIndex
 
End Function




