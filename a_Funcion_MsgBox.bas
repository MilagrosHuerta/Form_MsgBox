Attribute VB_Name = "a_Funcion_MsgBox"
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                     Form MsgBox                      --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autoría.                     --- '
' ---    Sirve para enviar mensajes con otro formato       --- '
' ---    y poder posicionarlo donde quieras                --- '
' ------------------------------------------------------------ '
Option Explicit
Public Titulo_Mensaje, Mensaje_Mostrar, Nombre_Icono As String
Public Boton_1, Boton_2, Boton_3 As String
Public continuar As Integer
Public numBotones As Byte
Public Posicion_Izda, Posicion_Top  As Integer
Function frmMsgBox(texoMensaje As Variant, btnTexto As String, _
                   Optional tituloForm As String, Optional iconoForm As String, _
                   Optional posLeftForm As Integer, Optional posTopForm As Integer, _
                   Optional textRigth As Boolean, Optional readRigth As Boolean)
' ------------------------------------------------------------------------- '
' --- Nombres Botones: bO, bOC, bARI, bYN, bYNC, bRC                    --- '
' --- Nombres Iconos: Informacion, Interrogacion, Exclamacion, Critico  --- '
' ------------------------------------------------------------------------- '
Dim longMensajeCorto As Integer
Dim longMensajeMedio As Integer
' Comprobamos que el Texto del botón sea correcto, si no lo es se sale
    If btnTexto <> "bO" And btnTexto <> "bOC" And btnTexto <> "bART" And btnTexto <> "bYN" _
                        And btnTexto <> "bYNC" And btnTexto <> "bRC" Then
        MsgBox "No has puesto un Tipo de BOTÓN correcto."
        End
' Comprobamos que el Texto del botón sea correcto, si no lo es se sale
    ElseIf iconoForm <> "Informacion" And iconoForm <> "Interrogacion" And iconoForm <> "Exclamacion" _
                                  And iconoForm <> "Critico" And iconoForm <> "" Then
        MsgBox "No has puesto un Tipo de ICONO correcto."
        End
    End If

    numBotones = Len(btnTexto) - 1
    longMensajeCorto = 160
    longMensajeMedio = 360
    Boton_1 = ""
    Boton_2 = ""
    Boton_3 = ""

    Select Case btnTexto
    Case "bO"
        Boton_1 = "Aceptar"
    Case "bOC"
        Boton_1 = "Aceptar"
        Boton_2 = "Cancelar"
    Case "bARI"
        Boton_1 = "Abortar"
        Boton_2 = "Reintentar"
        Boton_3 = "Ignorar"
    Case "bYN"
        Boton_1 = "SÍ"
        Boton_2 = "NO"
    Case "bYNC"
        Boton_1 = "SÍ"
        Boton_2 = "NO"
        Boton_3 = "Cancelar"
    Case "bRC"
        Boton_1 = "Reintentar"
        Boton_2 = "Cancelar"
    Case Else
        Boton_1 = "NO VALIDO"
    End Select
    
    Titulo_Mensaje = tituloForm
    Mensaje_Mostrar = texoMensaje
    Posicion_Izda = posLeftForm
    Posicion_Top = posTopForm
    Nombre_Icono = iconoForm
        
    Form_MsgBox.Show
End Function
