VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmAñadirRecordatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir recordatorio"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmAñadirRecordatorio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAñadirRecordatorio.frx":08CA
   ScaleHeight     =   5655
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MaxLength       =   64
      TextRTF         =   $"frmAñadirRecordatorio.frx":2922
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1085
      Caption         =   "Añadir recordatorio"
      EstiloDelBoton  =   1
      Picture         =   "frmAñadirRecordatorio.frx":29A5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   548
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3908
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2588
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   548
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3908
      TabIndex        =   16
      Top             =   1320
      Width           =   330
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Día:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2588
      TabIndex        =   15
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   548
      TabIndex        =   14
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minutos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1860
      TabIndex        =   13
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   548
      TabIndex        =   12
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1725
      TabIndex        =   11
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Añadir un recordatorio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   600
      TabIndex        =   10
      Top             =   480
      Width           =   3660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del recordatorio:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora del recordatorio:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   548
      TabIndex        =   8
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha del recordatorio:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   548
      TabIndex        =   7
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   4680
      Picture         =   "frmAñadirRecordatorio.frx":327F
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmAñadirRecordatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swEmpecé As Boolean
Public swEditar As Boolean
Public swMes As Byte
Public swDía As Byte
Public swAño As Integer
Public swTexto As String
Public swHora As Byte
Public swMinutos As Byte
Dim controlPresionado As Boolean
Dim swPulsóEnterParaAvanzar As Boolean

Private Sub ButtonTransparent1_Click()
    Dim fecha As String, rec As Recordatorio, hora As String, repetición As Byte, posición As Long, auxFecha As Date, resultComparación As Byte
    If Trim(RichTextBox1.Text) = "" Then
        Decir "no le has escrito un texto al recordatorio, por favor escribilo ahora"
        RichTextBox1.SetFocus
    Else
        If Len(Combo2.Text) = 1 Then
            fecha = "0" + Combo2.Text
        Else
            fecha = Combo2.Text
        End If
        fecha = fecha + "-"
        If Len(Trim(Str(Combo1.ListIndex + 1))) = 1 Then
            fecha = fecha + "0" + Trim(Str(Combo1.ListIndex + 1))
        Else
            fecha = fecha + Trim(Str(Combo1.ListIndex + 1))
        End If
        fecha = fecha + "-" + Combo3.Text
        hora = Combo4.Text + ":" + Combo5.Text
        
        auxFecha = Format(fecha, "dd/mm/yyyy")
        resultComparación = compararFechas(auxFecha, Format(Date, "dd/mm/yyyy"))
        
        If resultComparación = comparación.primeroMenor Then 'se controla la fecha
            frmMsgBox.swMostrarCancelar = False
            frmMsgBox.cadenaAMostrar = "Imposible añadir un recordatorio para una fecha pasada"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
            Combo1.SetFocus
            Exit Sub
        End If
        
        If resultComparación = comparación.iguales Then 'si es el día de hoy se ve q no sea una hora ya pasada
            resultComparación = compararHora(Format(hora, "HH:mm"), Format(Time, "HH:mm"))
            
            If resultComparación = comparación.primeroMenor Then 'se controla la fecha
                frmMsgBox.swMostrarCancelar = False
                frmMsgBox.cadenaAMostrar = "Se está queriendo añadir un recordatorio para una hora ya pasada del día de hoy. Imposible añadir."
                frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
                frmMsgBox.Show 1
                Combo4.SetFocus
                Exit Sub
            End If
        End If
        
        rec.fecha = Format(fecha, "dd/mm/yyyy")
        rec.hora = Format(hora, "HH:mm")
        rec.texto = RichTextBox1.Text
        rec.sonido = "predeterminado"
        rec.yaAnunciado = False
        
        Dim archivo As Byte, auxRecordatorio As Recordatorio
        If swEditar = True Then
            posición = 1
            archivo = FreeFile
            Open App.path + "\recordatorios\" + Trim(Str(Me.swAño)) + "\" + Trim(Str(Me.swMes)) + "\recordatorios.gui" For Random As #archivo Len = Len(auxRecordatorio)
            Do While Not EOF(archivo)   ' Repite hasta el final del archivo.
               Get #archivo, posición, auxRecordatorio   ' Lee el registro siguiente.
               If Format(auxRecordatorio.fecha, "dd/mm/yyyy") = Format(Trim(Str(Me.swDía)) + "/" + Trim(Str(Me.swMes)) + "/" + Trim(Str(Me.swAño))) Then
                    If Format(auxRecordatorio.hora, "HH:mm") = Format(Trim(Str(Me.swHora)) + ":" + Trim(Str(Me.swMinutos))) Then 'si es del día seleccionado
                        If auxRecordatorio.texto = Me.swTexto Then
                            Exit Do 'si es el mismo recordatorio, se sale del bucle pues se lo ha encontrado
                        End If
                    End If
                End If
                posición = posición + 1
            Loop
            Close #archivo
        End If
        
        If swEditar = False Then
            Call GuardarRecordatorio(rec)
        Else
            Call GuardarRecordatorioEnPosición(rec, posición)
        End If
        Call cargarRecordatorios
        frmMsgBox.swMostrarCancelar = False
        If miMateria <> "" Then
            If swEditar = False Then
                frmMsgBox.cadenaAMostrar = "El recordatorio se ha guardado exitosamente. Ahora vas a volver a tu carpeta de " + miMateria + "."
            Else
                frmMsgBox.cadenaAMostrar = "El recordatorio se ha modificado exitosamente. Ahora vas a volver a tu carpeta de " + miMateria + "."
            End If
        Else
            If swEditar = False Then
                frmMsgBox.cadenaAMostrar = "El recordatorio se ha guardado exitosamente. Ahora vas a volver a los accesorios."
            Else
                frmMsgBox.cadenaAMostrar = "El recordatorio se ha modificado exitosamente. Ahora vas a volver a los accesorios."
            End If
        End If
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Unload Me
    End If
End Sub

Private Sub ButtonTransparent1_GotFocus()
    Decir "botón " + ButtonTransparent1.Caption + ". aceptá con enter para añadir el recordatorio"
End Sub

Private Sub Combo1_Click()
    Call cargarDías
End Sub

Private Sub Combo1_GotFocus()
    If swEmpecé = True Then
        If swEditar = False Then
            Decir "Aquí podés añadir un recordatorio para que la mochila virtual te avise en el día y a la hora que vos programes, primero elegí con las flechas el mes en que querés que se active la alarma del recordatorio y aceptá con enter. Estás en: " + Combo1.Text
        Else
            Decir "Abriste el recordatorio para cambiarle alguno de sus datos. Modificá lo que quieras y guardalo para que suene nuevamente. Estás en elegir el mes en que querés que se active la alarma del recordatorio, usá las flechas y enter para cambiarlo. Estás en: " + Combo1.Text
        End If
        swEmpecé = False
    Else
        Decir "elegí con las flechas el mes en que querés que se active la alarma del recordatorio y aceptá con enter. Estás en: " + Combo1.Text
    End If
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo1.Text
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo2.Text
End Sub

Private Sub Combo2_GotFocus()
    Decir "ahora elegí con las flechas el día del recordatorio y aceptá con enter, si querés cambiar el mes que elegiste, apretá las teclas yift y tab. Estás en: " + Combo2.Text
End Sub

Private Sub Combo3_Click()
    Call cargarDías
End Sub

Private Sub Combo3_GotFocus()
    Decir "usá las flechas para elegir el año del recordatorio y aceptá con enter, si querés cambiar el día que elegiste, apretá las teclas yift y tab. Estás en: " + Combo3.Text
End Sub


Private Sub Combo3_Keyup(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo3.Text
End Sub

Private Sub Combo4_GotFocus()
    Decir "elegí con las flechas la hora del recordatorio y aceptá con enter, si querés volver a la fecha, apretá yift más tab. Estás en: " + Combo4.Text
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo4.Text
End Sub

Private Sub Combo5_GotFocus()
    Decir "usá las flechas para elegir los minutos y aceptá con enter, si querés volver a la hora, apretá yift más tab. Estás en: " + Combo5.Text
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Decir Combo5.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If shiftkey = vbAltMask And KeyCode = 18 Then 'se neutraliza el menú de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyReturn Then
        If TypeOf Me.ActiveControl Is ComboBox Or TypeOf Me.ActiveControl Is RichTextBox Then
            SendKeys ("{tab}")
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
'        If swCuadernoAbierto = False Then frmPrincipal.Show
        Unload Me
    End If
    
    If shiftkey = vbAltMask And KeyCode = vbKeyF4 Then swSalir = True 'si presiona alt + f4 se termina el programa
    If shiftkey = 0 And KeyCode = vbKeyF7 Then Decir "para abrir o ir al reproductor de música, tenés que estar en el menú principal o en una carpeta. ahora estás en añadir un recordatorio"
    
    If shiftkey = vbCtrlMask Then Decir "" 'para callar la voz
    
    'If shiftkey = vbAltMask Then Decir "al apretar la tecla alt sin ninguna otra tecla, vas al menú de la aplicación. Para leer los ítems de este menú necesitás jaws u otro lector de pantallas. Para volver a la mochila, apretá escape"
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyF1 Then 'leer la ayuda
         frmAyuda.formulario = formularios.añadirRecordatorios
         frmAyuda.Show
         Exit Sub
    End If

End Sub

Private Sub Form_Load()
    Dim mes As Byte
    Dim i As Long, añoActual As Long ', día As Byte
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    
    If swEditar = False Then
        ButtonTransparent1.Caption = "Añadir Recordatorio"
    Else
        ButtonTransparent1.Caption = "Modificar Recordatorio"
    End If
    
    '++++++++++++++++++++++++++++++++++++++
    'Cargar la fecha
    Combo1.AddItem "Enero"
    Combo1.AddItem "Febrero"
    Combo1.AddItem "Marzo"
    Combo1.AddItem "Abril"
    Combo1.AddItem "Mayo"
    Combo1.AddItem "Junio"
    Combo1.AddItem "Julio"
    Combo1.AddItem "Agosto"
    Combo1.AddItem "Setiembre"
    Combo1.AddItem "Octubre"
    Combo1.AddItem "Noviembre"
    Combo1.AddItem "Diciembre"
    
    mes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2) 'seleccionar el mes actual
    If swEditar = False Then 'si se edita, se marca el mes enviado
        Combo1.ListIndex = mes - 1
    Else
        Combo1.ListIndex = swMes - 1
    End If
    
    añoActual = Year(Date) 'cargar 20 años desde el actual y seleccionar el año actual
    If swEditar = False Then
        For i = añoActual To añoActual + 20
            Combo3.AddItem i
        Next
    Else
        If swAño < añoActual Then
            For i = swAño To añoActual + 20
                Combo3.AddItem i
            Next
        Else
            For i = añoActual To añoActual + 20
                Combo3.AddItem i
            Next
        End If
    End If
    
    If swEditar = False Then
        Combo3.ListIndex = 0
    Else
        For i = 0 To Combo3.ListCount - 1 'si se está editando, se selecciona el año que se envía
            If Combo3.List(i) = swAño Then
                Combo3.ListIndex = i
                Exit For
            End If
        Next
        
        RichTextBox1 = swTexto 'se copia el texto
    End If
    
    Call cargarDías
    
    
    '++++++++++++++++++++++++++++++++++++++
    'cargar la hora
    For i = 0 To 23 'cargar la hora
        Combo4.AddItem i
    Next
    
    If swEditar = False Then 'se marca la hora, según se edite o no
        Combo4.ListIndex = 0
    Else
        For i = 1 To Combo4.ListCount - 1 'si se está editando, se selecciona el año que se envía
            If Combo4.List(i) = swHora Then
                Combo4.ListIndex = i
                Exit For
            End If
        Next
    End If
    
    For i = 0 To 59 'se cargan los minutos
        Combo5.AddItem i
    Next
    
    If swEditar = False Then
        Combo5.ListIndex = 0
    Else
        For i = 1 To Combo5.ListCount - 1 'si se está editando, se selecciona el año que se envía
            If Combo5.List(i) = swMinutos Then
                Combo5.ListIndex = i
                Exit For
            End If
        Next
    End If
    
    swEmpecé = True
End Sub

Sub seleccionarDíaAlEditar()
    Dim i As Integer
    For i = 0 To Combo2.ListCount - 1 'si se está editando, se selecciona el día que se envía
        If Combo2.List(i) = swDía Then
            Combo2.ListIndex = i
            Exit For
        End If
    Next
End Sub

Sub cargarDías()
    Dim día As Byte, i As Byte, díaEnCombo As Byte
    If Combo3.Text <> "" Then
        
        día = Day(Date) 'cargar los días correspondientes al mes elegido
        If Combo2.ListIndex <> -1 Then díaEnCombo = Combo2.ListIndex + 1 'se ve cuál es el día del combo para seleccionar el mismo al cambiar el mes
        Combo2.Clear
        For i = 1 To cantDíasMes(Combo1.ListIndex + 1, Combo3.Text)
            Combo2.AddItem i
        Next
        
        If díaEnCombo = 0 Then 'si no había día seleccionado, se selecciona el actual
            For i = 0 To Combo2.ListCount - 1 'se selecciona el día actual
                If Combo2.List(i) = día Then Combo2.ListIndex = i
            Next
            If Combo2.ListIndex = -1 Then Combo2.ListIndex = 0 'si no está el día actual, se selecciona el primer día
        Else
            If díaEnCombo <= Combo2.ListCount Then
                Combo2.ListIndex = díaEnCombo - 1
            Else
                Combo2.ListIndex = 0
            End If
        End If
        
        If swEditar = True Then Call seleccionarDíaAlEditar
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If swSalir = True Then
        If SalirDelPrograma = True Then
            chauPrograma
        Else
            Cancel = 1
            swSalir = False
        End If
        Exit Sub
    End If
    
    swEditar = False
    
'    If swPulsóEnterParaAvanzar = False Then
        If swCuadernoAbierto = False Then
            frmAccesorios.Show
        Else
            Decir "volviendo a tu carpeta"
        End If
'    End If
    'Call contarFormularios(False)
End Sub


Private Sub richTextbox1_GotFocus()
    Dim cadena As String
    cadena = "Ahora escribí el texto que querés que te diga en el recordatorio. Cuando termines apretá enter"
    If RichTextBox1.Text <> "" Then cadena = cadena + ". Ya está escrito: " + RichTextBox1.Text
    'SendKeys ("^{end}") 'se pasa al final del cuadro
    Decir cadena
End Sub

Private Sub richTextbox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    If shiftkey = vbCtrlMask Then controlPresionado = True
    If KeyCode = vbKeyReturn Then
        SendKeys "{BACKSPACE}"
        SendKeys "{tab}"
        Exit Sub
    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        Decir "está escrito " + RichTextBox1.Text
    End If
        
    Dim letra As String, auxString As String
    If KeyCode = vbKeyDelete Then
        If RichTextBox1.Text <> "" Then 'si no está vacío
            If RichTextBox1.SelStart <> Len(RichTextBox1.Text) Then 'y no está al final de la hoja
                If RichTextBox1.SelText <> "" Then 'si hay algo seleccionado
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
                letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart + 1, 1)
                If letra = " " Then
                    Decir "borrando a la derecha el espacio", False
                ElseIf letra = Chr(9) Then
                    Decir "borrando a la derecha un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando a la derecha " + auxString
                End If
            Else
                Decir "imposible borrar, estás al final del texto"
            End If
        Else
            Decir "no se puede borrar a la derecha porque el texto está vacío"
        End If
    End If
    
    If KeyCode = vbKeyBack Then
        If RichTextBox1.Text = "" Then
            Decir "Ya está todo borrado"
        Else
            If RichTextBox1.SelStart = 0 Then
                Decir "imposible borrar porque estás al principio de la hoja"
            Else
                If RichTextBox1.SelText = "" Then 'si no hay nada seleccionado
                    letra = Mid(RichTextBox1.Text, RichTextBox1.SelStart, 1)
                Else
                    Decir "borrando el texto seleccionado"
                    Exit Sub
                End If
    
                If letra = " " Then
                    Decir "borrando el espacio"
                ElseIf letra = Chr(9) Then
                    Decir "borrando un salto"
                Else
                    auxString = traducirParaBorrar(letra)
                    Decir "borrando " + auxString
                End If
            End If
        End If
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyEnd Then Decir "final del texto"
    If shiftkey = 0 And KeyCode = vbKeyHome Then Decir "principio del texto"
    
End Sub

Private Sub richTextbox1_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    If Len(RichTextBox1.Text) < 64 Then
        If KeyAscii >= 32 And KeyAscii <= 255 And controlPresionado = False Then cadena = quéLetraSeApretó(KeyAscii)

        If KeyAscii = 9 Then cadena = "salto hacia adelante" 'tab
        If KeyAscii = 39 Then cadena = "apóstrofo"
        If KeyAscii = 123 Then cadena = "abre llave"
        If KeyAscii = 125 Then cadena = "cierra llave"
        If KeyAscii = 91 Then cadena = "abre corchete"
        If KeyAscii = 93 Then cadena = "cierra corchete"
        If KeyAscii = 64 Then cadena = "arroba"
        
        'leer la palabra al apretar espacio, punto, coma, etc.
        If KeyAscii = 32 Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") _
        Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Then cadena = cadena + decirPalabraAnterior(RichTextBox1)
        
        If cadena <> "" Then
            If RichTextBox1.SelBold = True Then cadena = cadena + " en negrita"
            If RichTextBox1.SelUnderline = True Then cadena = cadena + " subrayada"
            Decir cadena
        End If
    Else
        Decir "ya se escribieron las 64 letras que le podés poner al título de tu hoja"
    End If
    controlPresionado = False 'se resetea la variable

End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
            
    If shiftkey = vbCtrlMask And KeyCode = vbKeyLeft Then 'leer por palabras retrocediendo
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyRight Then 'leer por palabras avanzando
        Decir decirPalabraSiguiente(RichTextBox1) 'cadena
    End If
        
    If shiftkey = 0 And KeyCode = vbKeyRight Then 'avanzar de a caracter
        Decir decirLetraSiguiente(RichTextBox1)
    End If
    
    If shiftkey = 0 And KeyCode = vbKeyLeft Then 'retroceder de a caracter
        If RichTextBox1.SelStart = 0 And RichTextBox1.Text <> "" Then
            Decir "Estás en el principio del texto, delante de la letra " + decirLetraSiguiente(RichTextBox1)
        Else
            Decir decirLetraSiguiente(RichTextBox1)
        End If
    End If
    
    Dim TeclaShift, TeclaControl
    TeclaShift = (Shift And vbShiftMask) > 0
    TeclaControl = (Shift And vbCtrlMask) > 0

    Dim teclaApretada As Byte, control As Boolean, shift2 As Boolean
    Select Case KeyCode
        Case vbKeyA
            teclaApretada = tecla.a
        Case vbKeyUp
            teclaApretada = tecla.flechaArriba
        Case vbKeyDown
            teclaApretada = tecla.flechaAbajo
        Case vbKeyLeft
            teclaApretada = tecla.flechaIzquierda
        Case vbKeyRight
            teclaApretada = tecla.flechaDerecha
        Case vbKeyPageUp
            teclaApretada = tecla.avancePágina
        Case vbKeyPageDown
            teclaApretada = tecla.retrocesoPágina
        Case vbKeyHome
            teclaApretada = tecla.inicio
        Case vbKeyEnd
            teclaApretada = tecla.fin
        Case vbKeyBack
            teclaApretada = tecla.borrar
        Case vbKeyDelete
            teclaApretada = tecla.borrar
    End Select

    If TeclaControl Then
        control = True
    Else
        control = False
    End If

    If TeclaShift Then
        shift2 = True
    Else
        shift2 = False
    End If

    Call evaluarSelección(RichTextBox1, control, shift2, teclaApretada) 'se ve si hay selección
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyC Then 'copiar
        If RichTextBox1.SelText <> "" Then
            Decir "se copió el texto seleccionado"
        Else
            Decir "No se puede copiar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyX Then 'cortar
        If RichTextBox1.SelText <> "" Then
            Decir "se cortó el texto seleccionado"
        Else
            Decir "No se puede cortar porque no hay texto seleccionado. para seleccionar, usar shift más las flechas"
        End If
    End If
    
    If shiftkey = vbCtrlMask And KeyCode = vbKeyV Then 'pegar
        If Clipboard.GetText <> "" Then
            Decir "texto pegado: " + Clipboard.GetText
        Else
            Decir "No se puede pegar porque no hay texto copiado o cortado. para copiar, usar control más ce. para cortar, usar control más équis"
        End If
    End If
    controlPresionado = False 'se resetea la variable
End Sub
