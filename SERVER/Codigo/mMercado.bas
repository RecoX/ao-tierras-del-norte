Attribute VB_Name = "mMercado"
Option Explicit

' @ Utilizado para la venta de personajes
Private Type tMercado
    valor As Long
    Nombre As String
    
    ' @ Venta de Pjs
    Recibidor As String
End Type

Public Enum eMercado
    Venta = 1
    Cambio = 2
End Enum

Public Const MAX_PJS_MERCADO As Integer = 1000

Public Mercado(1 To MAX_PJS_MERCADO) As tMercado

Public Function FreeSlotMercado() As Integer
    Dim i As Long
    ' @ Buscamos un slot libre para nuestro personaje

    For i = 1 To MAX_PJS_MERCADO
        If Mercado(i).Nombre = vbNullString Then
            FreeSlotMercado = i
            Exit For
        End If
    Next i

End Function
Sub CrearMercadofile()
Dim intFile As Integer

    intFile = FreeFile
    Dim i As Integer
    
    Open DatPath & "MERCADO.DAT" For Output As #intFile
        Print #intFile, "[INIT]"
        
        For i = 1 To 1000
            Print #intFile, "PERSONAJE" & i & "=---"
        Next i
    Close #intFile
End Sub
Sub LoadMercadoArgentum()
    ' @ Cargamos mercado de ventas y mercado de cambio de personajes.
    
    Dim i As Integer
    Dim ln As String
    
    If Not FileExist(DatPath & "MERCADO.DAT") Then
        Call CrearMercadofile
    End If
    For i = 1 To MAX_PJS_MERCADO
        ln = GetVar(App.Path & "\DAT\" & "Mercado.dat", "INIT", "PERSONAJE" & i)
        Mercado(i).Nombre = ReadField(1, ln, 45)
        Mercado(i).Recibidor = ReadField(2, ln, 45)
        Mercado(i).valor = val(ReadField(3, ln, 45))
    Next i
End Sub

Private Function CheckDatosMercado(ByVal UserIndex As Integer, ByVal UserName As String, ByVal Email As String, ByVal ClavePin As String, _
    ByVal Pw As String, ByVal Depositor As String) As Boolean
    
    CheckDatosMercado = False
    
    '¿Existe el personaje a postear?
    If Not FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Call WriteConsoleMsg(UserIndex, "No puedes postear un personaje inexistente.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    If Depositor <> vbNullString Then
        If Not FileExist(CharPath & UCase$(Depositor) & ".chr", vbNormal) Then
            Call WriteConsoleMsg(UserIndex, "El personaje que recibirá el dinero no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    
    
    If val(GetVar(CharPath & UCase$(UserName) & ".chr", "STATS", "ELV")) < 35 Then
        Call WriteConsoleMsg(UserIndex, "El personaje debe ser mayor a nivel 35.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    If UCase$(Email) <> UCase$(GetVar(CharPath & UCase$(UserName) & ".chr", "CONTACTO", "eMail")) Then
        Call WriteConsoleMsg(UserIndex, "El email ingresado no corresponde al personaje.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    If UCase$(Pw) <> UCase$(GetVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Password")) Then
        Call WriteConsoleMsg(UserIndex, "El password ingresado no corresponde al personaje.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    If UCase$(ClavePin) <> UCase$(GetVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Pin")) Then
        Call WriteConsoleMsg(UserIndex, "La clave Pin ingresada no corresponde al personaje.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    

    
        If UserList(UserIndex).flags.indexmercado = 1 Then
            Call WriteConsoleMsg(UserIndex, "Tu personaje ya está publicado, no puedes publicarlo otra vez.", FontTypeNames.FONTTYPE_WARNING)
            Exit Function
        End If
        
        Dim CharFile$
    CharFile = CharPath & ".chr"
    Dim GG As String
    GG = val(GetVar(CharFile, "VENTA", "iVenta"))
    If UserList(UserIndex).flags.EstaEnMercado = True And GG = "1" Then
    Call WriteConsoleMsg(UserIndex, "Tu personaje ya está publicado, no puedes publicarlo otra vez.", FontTypeNames.FONTTYPE_WARNING)
            Exit Function
            End If
   
    
    CheckDatosMercado = True

End Function
Public Sub ActualizarMercado(ByVal index As Integer)
    ' @ Quitamos el personaje del mercado.dat
    With Mercado(index)
        
        Mercado(index).Nombre = vbNullString
        Mercado(index).Recibidor = vbNullString
        Mercado(index).valor = 0
        
        Call WriteVar(DatPath & "Mercado.Dat", "INIT", "PERSONAJE" & index, "---")
    End With
End Sub

Public Sub QuitarPersonaje(ByVal UserIndex As Integer)
    ' @ Quitamos nuestro personaje del mercado
    With UserList(UserIndex)
        If .flags.indexmercado = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu personaje no está en el mercado de Tierras del Norte AO.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        ' @ Quitamos el personaje
        Call ActualizarMercado(.flags.indexmercado)
        Call WriteConsoleMsg(UserIndex, "El personaje fue quitado correctamente.", FontTypeNames.FONTTYPE_WARNING)
        .flags.indexmercado = 0
    End With
End Sub

Public Sub EnviarOfertaCambio(ByVal UserIndex As Integer, ByVal indexmercado As Integer)
    ' @ Enviamos oferta al personaje seleccionado
    With UserList(UserIndex)
        '¿Tiene lugar para recibir más ofertas?
        
        Dim OFFERS(1 To 10) As String
        Dim OfertaEnviada As Boolean
        Dim PuedeEnviar As Byte
        Dim tIndex As Integer
        Dim Freeslot As Byte
        Dim i As Integer
        tIndex = NameIndex(Mercado(indexmercado).Nombre)
        
        ' @ Personaje logueado, le enviamos oferta y notificamos
        If tIndex <> 0 Then
            ' @ Buscamos slot libre en el personaje para que reciba nuestra solicitud
            For i = 1 To 10
                If UserList(tIndex).Ofertas(i).OfertasRecibidas = vbNullString Then
                    Freeslot = i
                    Exit For
                End If
            Next i
            
            ' @ Chequeamos que nosotros podemos guardar la solicitud y enviar una oferta en caso de no haberla enviado antes.
            For i = 1 To 10
                ' ¿Tenemos lugar para guardar nuestra oferta en MIS datos?
                If .Ofertas(i).OfertasHechas = vbNullString Then
                    PuedeEnviar = i
                    Exit For
                End If
            Next i
            
            ' @ ¿Ya le hemos enviado una oferta?
            For i = 1 To 10
                If UserList(tIndex).Ofertas(i).OfertasRecibidas = .name Then
                    OfertaEnviada = True
                    Exit For
                End If
            Next i
            
                             If Not .Pos.map = 1 Then
    'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
    WriteConsoleMsg UserIndex, "Para enviar una solicitud de cambio debes estar en Ullathorpe.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
     End If
     
            If Mercado(indexmercado).Nombre = vbNullString Then
        Call WriteConsoleMsg(UserIndex, "No has seleccionado ningún personaje publicado en el Mercado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                        
            If Mercado(indexmercado).valor < 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje está publicado para aceptar cambios. Por lo tanto no recibe dinero a cambio de su personaje.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
            
            If PuedeEnviar = 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes espacio para enviar otra oferta de cambio. Antes debes eliminar solicitudes enviadas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Freeslot = 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no tiene más lugar para recibir ofertas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If OfertaEnviada = True Then
                    Call WriteConsoleMsg(UserIndex, "Ya le has enviado una oferta al personaje. Espera respuesta de él o bien cancela la oferta y envíasela de vuelta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                UserList(tIndex).Ofertas(Freeslot).OfertasRecibidas = .name
                .Ofertas(PuedeEnviar).OfertasHechas = UserList(tIndex).name
                Call WriteConsoleMsg(tIndex, "El personaje " & .name & " te ha enviado una solicitud para cambiar por tu personaje. La misma podrás verla desde el comando /MERCADO", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Le has enviado oferta de cambio a " & Mercado(indexmercado).Nombre & ". Espera respuesta del personaje", FontTypeNames.FONTTYPE_INFO)
                
            End If
        
        Else
            For i = 1 To 10
                ' ¿Tenemos lugar para guardar nuestra oferta en MIS datos?
                If .Ofertas(i).OfertasHechas = vbNullString Then
                    PuedeEnviar = i
                    Exit For
                End If
            Next i
            
            ' @ Chequeamos que nuestra oferta no haya sido recibida
            For i = 1 To 10
                OFFERS(i) = GetVar(CharPath & Mercado(indexmercado).Nombre & ".chr", "MERCADO", "OFERTARECIBIDA" & i)
                
                If OFFERS(i) = .name Then
                    Call WriteConsoleMsg(UserIndex, "Ya le has enviado una oferta al personaje. Espera respuesta de él", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Next i
            
            ' @ Buscamos un slot libre para almacenar nuestra oferta
            For i = 1 To 10
                OFFERS(i) = GetVar(CharPath & Mercado(indexmercado).Nombre & ".chr", "MERCADO", "OFERTARECIBIDA" & i)
                
                If OFFERS(i) = vbNullString Then
                    Freeslot = i
                    Exit For
                End If
            Next i
            
            If Freeslot = 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no tiene más lugar para recibir ofertas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If PuedeEnviar = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No tienes más lugar para enviar ofertas de cambios. Elimina algunas e intenta más tarde.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' @ Le guardamos la oferta
                Call WriteVar(CharPath & Mercado(indexmercado).Nombre & ".chr", "MERCADO", "OFERTARECIBIDA" & Freeslot, .name)
                
                ' @ Alacenamos nuestra oferta
                .Ofertas(PuedeEnviar).OfertasHechas = Mercado(indexmercado).Nombre
                Call WriteConsoleMsg(UserIndex, "Le has enviado oferta de cambio a " & Mercado(indexmercado).Nombre & ". Espera respuesta del personaje", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    
    End With
End Sub
Public Sub RechazarOfertaCambio(ByVal UserIndex As Integer, ByVal index As Byte)
    ' @ Rechazamos la oferta que nos ofreció el INDEX
    With UserList(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta de " & .Ofertas(index).OfertasRecibidas & ".", FontTypeNames.FONTTYPE_GUILD)
        
        ' @ Borramos la oferta
        .Ofertas(index).OfertasRecibidas = vbNullString
    End With
End Sub

Public Sub CancelarOfertaHecha(ByVal UserIndex As Integer, ByVal index As Byte)
    ' @ Cancelamos la oferta que realizamos al personaje.
    With UserList(UserIndex)
         Call WriteConsoleMsg(UserIndex, "Has cancelado la oferta que le mandaste anteriormente a " & .Ofertas(index).OfertasHechas & ".", FontTypeNames.FONTTYPE_INFO)

        ' @ Borramos la oferta que hicimos
        .Ofertas(index).OfertasHechas = vbNullString

    End With
End Sub
Public Sub TransferirDatosPersonaje(ByVal Comprador As String, ByVal Vendedor As String)
    ' @ Cargamos los datos del comprador
    Dim Comprador_Pw As String
    Dim Comprador_Pin As String
    Dim Comprador_Email As String
    Comprador_Pw = GetVar(CharPath & Comprador & ".chr", "INIT", "PASSWORD")
    Comprador_Pin = GetVar(CharPath & Comprador & ".chr", "INIT", "PIN")
    Comprador_Email = GetVar(CharPath & Comprador & ".chr", "CONTACTO", "EMAIL")
    
    ' @ Cargamos los datos del vendedor
    Dim Vendedor_Pw As String
    Dim Vendedor_Pin As String
    Dim Vendedor_Email As String
    Vendedor_Pw = GetVar(CharPath & Vendedor & ".chr", "INIT", "PASSWORD")
    Vendedor_Pin = GetVar(CharPath & Vendedor & ".chr", "INIT", "PIN")
    Vendedor_Email = GetVar(CharPath & Vendedor & ".chr", "CONTACTO", "EMAIL")
    
    
    '___________________________________________________________Benyi
    
    

    ' @ Intercambiamos datos
    Call WriteVar(CharPath & Comprador & ".chr", "CONTACTO", "EMAIL", Vendedor_Email)
    Call WriteVar(CharPath & Comprador & ".chr", "INIT", "PASSWORD", Vendedor_Pw)
    Call WriteVar(CharPath & Comprador & ".chr", "INIT", "PIN", Vendedor_Pin)

    Call WriteVar(CharPath & Vendedor & ".chr", "CONTACTO", "EMAIL", Comprador_Email)
    Call WriteVar(CharPath & Vendedor & ".chr", "INIT", "PASSWORD", Comprador_Pw)
    Call WriteVar(CharPath & Vendedor & ".chr", "INIT", "PIN", Comprador_Pin)
End Sub
Public Sub AceptarOfertaMercado(ByVal UserIndex As Integer, ByVal index As Integer)
    ' @ Aceptamos la oferta seleccionada
    ' @ Se realiza el intercambio
    ' @ Eliminamos todas las solicitudes hechas/recibidas.
    
    Dim i As Integer
    Dim OFFERS(1 To 10) As String
    Dim tIndex As Integer
    Dim PuedeAceptar As Boolean
    With UserList(UserIndex)
                     If Not .Pos.map = 1 Then
    'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
    WriteConsoleMsg UserIndex, "Para enviar una solicitud de cambio debes estar en Ullathorpe.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
     End If
    
        tIndex = NameIndex(.Ofertas(index).OfertasRecibidas)
        Dim Nombre(2) As String
            
            Nombre(1) = .name
            Nombre(2) = .Ofertas(index).OfertasRecibidas
            
            ' @ Actualizamosy borramos datos necesarios
            Call ActualizarMercado(index)
            For i = 1 To 10
                .Ofertas(i).OfertasHechas = vbNullString
                .Ofertas(i).OfertasRecibidas = vbNullString
                
                Call WriteVar(CharPath & Nombre(2) & ".chr", "MERCADO", "OfertasHecha" & i, vbNullString)
                Call WriteVar(CharPath & Nombre(2) & ".chr", "MERCADO", "OfertasRecibida" & i, vbNullString)
            Next i
            
            
     
            ' @ Cerramos conexiones
            Call CloseSocket(tIndex): Call CloseSocket(UserIndex)
            
            ' @ Transferimos el personaje
            Call TransferirDatosPersonaje(Nombre(1), Nombre(2))
            
            
            Nombre(1) = .name
            Nombre(2) = .Ofertas(index).OfertasRecibidas
            
            ' @ Actualizamosy borramos datos necesarios
            Call ActualizarMercado(index)
            For i = 1 To 10
                .Ofertas(i).OfertasHechas = vbNullString
                .Ofertas(i).OfertasRecibidas = vbNullString
                
                UserList(tIndex).Ofertas(i).OfertasHechas = vbNullString
                UserList(tIndex).Ofertas(i).OfertasRecibidas = vbNullString
            Next i
            
            ' @ Cerramos conexiones
            Call CloseSocket(UserIndex): Call CloseSocket(tIndex)

    End With
End Sub
Public Sub ComprarPersonajeMercado(ByVal UserIndex As Integer, ByVal indexmercado As Integer)

    ' @ Compramos el personaje seleccionado
    Dim GLD As Long
    Dim Target As Integer
    Dim Pw As String
    Dim Pin As String
    Dim Email As String
    With UserList(UserIndex)
        If Mercado(indexmercado).valor > .Stats.GLD Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero para comprar el personaje. Recuerda tenerlo en tu billetera a la hora de realizar la compra.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
                             If Not .Pos.map = 1 Then
    'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
    WriteConsoleMsg UserIndex, "Sólo puedes comprar personajes si estás en Ullathorpe.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
     End If
     
             If Mercado(indexmercado).valor < 100000 Then
            Call WriteConsoleMsg(UserIndex, "El valor mínimo para comprar un personaje es de 100.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
                If Mercado(indexmercado).valor = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje está publicado para aceptar cambios. Por lo tanto no recibe oro a cambio de su personaje.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Mercado(indexmercado).Nombre = vbNullString Then
        Call WriteConsoleMsg(UserIndex, "No has seleccionado ningún personaje publicado en el Mercado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        
        ' @ Personaje que recibe el dinero
        GLD = val(GetVar(CharPath & UCase$(Mercado(indexmercado).Recibidor) & ".chr", "STATS", "BANCO"))
        GLD = GLD + Mercado(indexmercado).valor
        Call WriteVar(CharPath & UCase$(Mercado(indexmercado).Recibidor) & ".chr", "STATS", "BANCO", GLD)

        
        ' @ Le quitamos el dinero al comprador
        .Stats.GLD = .Stats.GLD - Mercado(indexmercado).valor
        Call WriteUpdateGold(UserIndex)

       ' @ Actualizamos los datos del personaje
        Pw = GetVar(CharPath & UCase$(.name) & ".chr", "INIT", "PASSWORD")
        Call WriteVar(CharPath & UCase$(Mercado(indexmercado).Nombre) & ".chr", "INIT", "PASSWORD", Pw)
        
        Call WriteConsoleMsg(UserIndex, "Has comprado el personaje " & Mercado(indexmercado).Nombre & ".", FontTypeNames.FONTTYPE_GUILD)
        
        ' @ Personaje comprado logueado?
        Target = NameIndex(Mercado(indexmercado).Nombre)
        If Target <> 0 Then
            Call CloseSocket(Target)
        End If
        
        Call WriteVar(CharPath & UCase$(Mercado(indexmercado).Nombre) & ".chr", "INIT", "PIN", .Pin)
        Call WriteVar(CharPath & UCase$(Mercado(indexmercado).Nombre) & ".chr", "CONTACTO", "EMAIL", .Email)
        
        Call ActualizarMercado(indexmercado)
    End With
End Sub
Public Sub PublicarPersonaje(ByVal UserIndex As Integer, _
                                ByVal UserName As String, _
                                ByVal Email As String, _
                                ByVal Pin As String, _
                                ByVal Pw As String, _
                                ByVal valor As Long, _
                                Optional ByVal Depositor As String = vbNullString)

    ' @ Publicamos el personaje elegido
    With UserList(UserIndex)
                     If Not .Pos.map = 1 Then
    'Call WriteConsoleMsg(UserIndex, "¡¡No puedes ingresar si no estas en Ullathorpe!!.", FontTypeNames.FONTTYPE_INFO)
    WriteConsoleMsg UserIndex, "Para publicar tu personaje debes estar en Ullathorpe.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
     End If

    
    
        Dim SlotLibre As Integer
            SlotLibre = FreeSlotMercado
        
        If SlotLibre = 0 Then
            Call WriteConsoleMsg(UserIndex, "No hay mas lugar en el Mercado para tu personaje.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿Los datos ingresados son los correctos?
        If CheckDatosMercado(UserIndex, UserName, Email, Pin, Pw, Depositor) Then
        
            ' @ Guardamos la pos para futuros chequeos
            .flags.indexmercado = SlotLibre
            
            If valor <> 0 And Depositor <> vbNullString Then
                Mercado(SlotLibre).Recibidor = Depositor
                Mercado(SlotLibre).valor = valor
            End If
                
            Mercado(SlotLibre).Nombre = UserName
            Call WriteVar(DatPath & "MERCADO.DAT", "INIT", "PERSONAJE" & SlotLibre, .name & "-" & Depositor & "-" & valor)
            Call WriteConsoleMsg(UserIndex, "Has publicado el personaje " & UserName & " en el mercado de Tierras del Norte AO. Recuerda estar Offline para recibir el oro si el personaje fue publicado por venta.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Recuerda que el personaje que recibe el oro debe estar OFFLINE, en el caso de ventas.", FontTypeNames.FONTTYPE_GUILD)
        End If
    
    End With
End Sub



