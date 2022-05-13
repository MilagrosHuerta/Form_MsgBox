VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_MsgBox 
   Caption         =   "Form_MsgBox"
   ClientHeight    =   6015
   ClientLeft      =   4110
   ClientTop       =   4470
   ClientWidth     =   10875
   OleObjectBlob   =   "Form_MsgBox.frx":0000
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------ '
' ---              Formulario creado por                   --- '
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
Private Sub UserForm_Initialize()
Dim a, b, i As Integer
' --------------------------------- '
' --- Pone un título al mensaje --- '
' --------------------------------- '
    If Titulo_Mensaje <> "" Then Form_MsgBox.Caption = Titulo_Mensaje
' --------------------------------------------------------- '
' --- Posiciona el mensaje en funcion de los parámetros --- '
' --------------------------------------------------------- '
    If Posicion_Izda = 0 And Posicion_Top = 0 Then
        Form_MsgBox.StartUpPosition = 2     ' Centrar en pantalla
    Else
        Form_MsgBox.Left = Posicion_Izda
        Form_MsgBox.Top = Posicion_Top
    End If
' ----------------------------------------------------------- '
' --- Tamaño mensaje en funcion de la longitud del texto ---  '
' ----------------------------------------------------------- '
    If Len(Mensaje_Mostrar) < 200 Then
        i = 1
    ElseIf Len(Mensaje_Mostrar) < 300 Then
        i = 2
    ElseIf Len(Mensaje_Mostrar) < 400 Then
        i = 3
    ElseIf Len(Mensaje_Mostrar) < 500 Then
        i = 4
    ElseIf Len(Mensaje_Mostrar) < 600 Then
        i = 5
    Else
        i = 6
    End If
    a = 30 * (2 + i)
    b = 208 + 2 * 30 * i
    
    Form_MsgBox.Height = b
    Icono_Interrogacion.Top = a
    Icono_Informacion.Top = a
    Icono_Exclamacion.Top = a
    Icono_Critico.Top = a
    frmMensaje_Mostrar.Height = a - 10
    frmBoton_1.Top = a + 3
    frmBoton_2.Top = a + 3
    frmBoton_3.Top = a + 3
    TextoBoton_1.Top = a + 4.5
    TextoBoton_2.Top = a + 4.5
    TextoBoton_3.Top = a + 4.5
    frmMensaje_Mostrar = Mensaje_Mostrar
' ----------------------------------------------------------------------- '
' --- Nombre de Botones y Visibles en funcion de la variable btnTexto --- '
' ----------------------------------------------------------------------- '
    frmBoton_1.Visible = True
    TextoBoton_1.Value = Boton_1
    If numBotones = 1 Then
        frmBoton_2.Visible = False
        frmBoton_3.Visible = False
        TextoBoton_2.Visible = False
        TextoBoton_3.Visible = False
    ElseIf numBotones = 2 Then
        frmBoton_2.Visible = True
        frmBoton_3.Visible = False
        TextoBoton_2.Visible = True
        TextoBoton_3.Visible = False
        TextoBoton_2.Value = Boton_2
    Else
        frmBoton_2.Visible = True
        frmBoton_3.Visible = True
        TextoBoton_2.Visible = True
        TextoBoton_3.Visible = True
        TextoBoton_2.Value = Boton_2
        TextoBoton_3.Value = Boton_3
    End If
    
' --------------------------------------------------------- '
' --- Mostrar la imagen en funcion de la variable ICONO --- '
' --------------------------------------------------------- '
    Select Case Nombre_Icono
    Case "Informacion"
        Icono_Informacion.Visible = True
        Icono_Interrogacion.Visible = False
        Icono_Exclamacion.Visible = False
        Icono_Critico.Visible = False
    Case "Interrogacion"
        Icono_Informacion.Visible = False
        Icono_Interrogacion.Visible = True
        Icono_Exclamacion.Visible = False
        Icono_Critico.Visible = False
    Case "Exclamacion"
        Icono_Informacion.Visible = False
        Icono_Interrogacion.Visible = False
        Icono_Exclamacion.Visible = True
        Icono_Critico.Visible = False
    Case "Critico"
        Icono_Informacion.Visible = False
        Icono_Interrogacion.Visible = False
        Icono_Exclamacion.Visible = False
        Icono_Critico.Visible = True
    Case Else                                   ' No muestra ninguna imagen
        Icono_Informacion.Visible = False
        Icono_Interrogacion.Visible = False
        Icono_Exclamacion.Visible = False
        Icono_Critico.Visible = False
    End Select

End Sub
Private Sub frmBoton_1_Click()
' ----------------------------------------------------------------------------------- '
' --- Se asigna el valor de VBA que se requiera a los botones, según el que pulse --- '
' ----------------------------------------------------------------------------------- '
    Select Case Boton_1
    Case "SÍ"
        continuar = vbYes
    Case "Aceptar"
        continuar = vbOK
    Case "Abortar"
        continuar = vbAbort
    Case "Reintentar"
        continuar = vbRetry
    End Select
            
    Unload Me
End Sub
Private Sub frmBoton_2_Click()
' ----------------------------------------------------------------------------------- '
' --- Se asigna el valor de VBA que se requiera a los botones, según el que pulse --- '
' ----------------------------------------------------------------------------------- '
    Select Case Boton_2
    Case "NO"
        continuar = vbNo
    Case "Cancelar"
        continuar = vbCancel
    Case "Reintentar"
        continuar = vbRetry
    End Select
    
    Unload Me
End Sub
Private Sub frmBoton_3_Click()
' ----------------------------------------------------------------------------------- '
' --- Se asigna el valor de VBA que se requiera a los botones, según el que pulse --- '
' ----------------------------------------------------------------------------------- '
    Select Case Boton_3
    Case "Cancelar"
        continuar = vbCancel
    Case "Ignorar"
        continuar = vbIgnore
    End Select
        
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' Macro para evitar que se pueda cerrar el formulario en la X de arriba a la derecha
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
