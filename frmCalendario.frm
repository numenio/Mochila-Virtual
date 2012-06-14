VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actividades Guardadas"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   ForeColor       =   &H00000000&
   Icon            =   "frmCalendario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCalendario.frx":08CA
   ScaleHeight     =   6945
   ScaleWidth      =   9930
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4320
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin TransparentButton.ButtonTransparent btnD�a 
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "ButtonTransparent1"
      EstiloDelBoton  =   1
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
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
      ShowFocusRect   =   0   'False
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A�o:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6600
      TabIndex        =   4
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Materia:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1440
      X2              =   3720
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swEmpez� As Boolean
Public swEstoyAbierto As Boolean

Private Sub btnD�a_Click(Index As Integer)
    Dim mes As Date
    
    mes = Format(Str(Index + 1) + "/" + Trim(Str(decodificarMes(Combo3.Text))) + "/" + Combo2.Text)
    frmA�adirActividad.d�aParaCargarActividades = mes
    frmA�adirActividad.materia = Combo1.Text
    frmA�adirActividad.swCargarFecha = True
    If Not IsNumeric(btnD�a(Index).Caption) Then
        frmA�adirActividad.swEditarActividades = True
'        frmA�adirActividad.Show 1
    Else
        frmA�adirActividad.swEditarActividades = False
    End If
    frmA�adirActividad.Show 1 'a�adir una actividad
End Sub

Private Sub Combo1_Click()
    Dim mes As Byte, d�a As Byte, a�o As Integer
    
    If swEmpez� = True Then
        a�o = Combo2.Text
        mes = decodificarMes(Combo3.Text)
        d�a = nombreDeD�a(1, mes, a�o)
        Call chequearD�a1(d�a)
        Call cargarMesEnPantalla(mes, a�o, d�a)
        Call llenarActividades(Trim(Combo1.Text), mes, a�o)
        Call arreglarForm(mes, a�o)
    End If
End Sub

Private Sub Combo2_Click() 'si se cambian los a�os
    Dim mes As Byte, d�a As Byte, a�o As Integer
    
    If swEmpez� = True Then
        a�o = Combo2.Text
        mes = decodificarMes(Combo3.Text)
        d�a = nombreDeD�a(1, mes, a�o)
        Call chequearD�a1(d�a)
        Call cargarMesEnPantalla(mes, a�o, d�a)
        Call llenarActividades(Trim(Combo1.Text), mes, a�o)
        Call arreglarForm(mes, a�o)
    End If
End Sub

Private Sub Combo3_Click() 'si se cambian los meses
    Dim mes As Byte, d�a As Byte, a�o As Integer
    
    If swEmpez� = True Then
        If Combo2.Text <> "" Then
            a�o = Combo2.Text
            mes = decodificarMes(Combo3.Text)
            d�a = nombreDeD�a(1, mes, a�o)
            Call chequearD�a1(d�a)
            Call cargarMesEnPantalla(mes, a�o, d�a)
            Call llenarActividades(Trim(Combo1.Text), mes, a�o)
            Call arreglarForm(mes, a�o)
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim shiftkey As Integer
'    shiftkey = Shift And 7
    If KeyCode = vbKeyF1 Then ShellExecute 0, "open", "hh.exe", App.Path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/ver actividades.htm", "", 1
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim j As Integer
       
    swEmpez� = False
    swEstoyAbierto = True
    
    For j = 2008 To Year(Date) + 10 'se llenan los a�os, desde el actual m�s otros 10 a�os
        Combo2.AddItem j
    Next
    
    Call llenarComboMaterias(Combo1) 'se llenan las materias
    
    Dim mes As Byte
    Combo3.AddItem "Enero" 'se llenan los meses
    Combo3.AddItem "Febrero"
    Combo3.AddItem "Marzo"
    Combo3.AddItem "Abril"
    Combo3.AddItem "Mayo"
    Combo3.AddItem "Junio"
    Combo3.AddItem "Julio"
    Combo3.AddItem "Agosto"
    Combo3.AddItem "Setiembre"
    Combo3.AddItem "Octubre"
    Combo3.AddItem "Noviembre"
    Combo3.AddItem "Diciembre"
    
    mes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2) 'se activa el mes actual
    Combo3.ListIndex = mes - 1
    
    Dim a�oABuscar As Integer 'se activa el a�o actual
    a�oABuscar = Year(Date)
    For j = 0 To Combo2.ListCount - 1
        If Combo2.List(j) = a�oABuscar Then
            Combo2.ListIndex = j
            Exit For
        End If
    Next
    
    Combo1.ListIndex = 0 'se activa la primer materia
    
    btnD�a(0).Caption = "1"
    For j = 1 To 30
        Load btnD�a(j)
        btnD�a(j).Caption = Str(j + 1) 'se carga en el bot�n 0 el d�a 1
        btnD�a(j).Visible = True
    Next
    
    For j = 1 To 6
        Load Label2(j)
    Next
    
    Dim d�a As Byte, a�o As Integer
    
    a�o = Combo2.Text
    mes = decodificarMes(Combo3.Text)
    d�a = nombreDeD�a(1, mes, a�o)
    Call chequearD�a1(d�a)
    Call cargarMesEnPantalla(mes, a�o, d�a)
    Call llenarActividades(Trim(Combo1.Text), mes, a�o)
    Call arreglarForm(mes, a�o)
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swEmpez� = True
End Sub

Sub cargarMesEnPantalla(mes As Byte, a�o As Integer, d�ndeEmpezar As Byte)
    Dim j As Integer, espacioEntreBotones As Integer, espacioHorizontalbotones As Byte
    Dim col1, numColumna As Byte, numFila As Byte, fila1X As Integer, fila1Y As Integer
       
    espacioEntreBotones = 200
    espacioHorizontalbotones = 200
    
    btnD�a(0).Caption = 1
    
    For j = 1 To 30
        btnD�a(j).Visible = False
        btnD�a(j).ForeColor = &H8000000F 'blanco
        btnD�a(j).Caption = j + 1
    Next
        
    With Label2(0)
        .Width = btnD�a(0).Width
        .Height = btnD�a(0).Height
        .Top = 1480
        .Left = 480
        .Visible = True
        .Caption = "L"
    End With
    
    numFila = 1
    
    fila1X = 480
    fila1Y = 2200
    
    btnD�a(0).Top = fila1Y
    btnD�a(0).Left = CInt(fila1X) + btnD�a(0).Width * (CInt(d�ndeEmpezar) - 1) + CInt(espacioHorizontalbotones) * (CInt(d�ndeEmpezar) - 1) '+ espacioHorizontalbotones
    
    d�ndeEmpezar = d�ndeEmpezar + 1
    
    For j = 1 To cantD�asMes(mes, a�o) - 1
        If d�ndeEmpezar = 8 Then
            numFila = numFila + 1
            btnD�a(j).Left = fila1X
            d�ndeEmpezar = 1
        Else
            btnD�a(j).Left = btnD�a(j - 1).Left + espacioHorizontalbotones + btnD�a(j).Width
        End If
        
        If numFila = 1 Then
            btnD�a(j).Top = fila1Y
        Else
            btnD�a(j).Top = fila1Y + btnD�a(j).Width * (numFila - 1)
        End If
        
        If d�ndeEmpezar = 6 Or d�ndeEmpezar = 7 Then btnD�a(j).ForeColor = &H4080& 'marr�n
        
        
        d�ndeEmpezar = d�ndeEmpezar + 1
        btnD�a(j).Visible = True
        btnD�a(j).Refresh
        
        If j < 7 Then 'se cargan los d�as, desde el martes al domingo
            With Label2(j)
                .Width = btnD�a(0).Width
                .Height = btnD�a(0).Height
                .Top = 1480
                .Left = Label2(j - 1).Left + espacioHorizontalbotones + .Width
                .Visible = True
                
                Select Case j
                    Case 1
                        .Caption = "M"
                    Case 2
                        .Caption = "M"
                    Case 3
                        .Caption = "J"
                    Case 4
                        .Caption = "V"
                    Case 5
                        .Caption = "S"
                    Case 6
                        .Caption = "D"
                End Select
            End With
        End If
    Next
        
    btnD�a(0).Refresh
    
    With Line1
        .X1 = fila1X
        .X2 = Label2(6).Left + btnD�a(0).Width + espacioHorizontalbotones
        .Y1 = fila1Y - espacioHorizontalbotones
        .Y2 = fila1Y - espacioHorizontalbotones
    End With
End Sub

Sub arreglarForm(mes As Byte, a�o As Integer)
    Dim d�asDelMes As Byte
    d�asDelMes = cantD�asMes(mes, a�o) '-1
    Me.Height = btnD�a(d�asDelMes - 1).Top + btnD�a(d�asDelMes - 1).Height + 1000
End Sub

Sub chequearD�a1(d�a As Byte)
    If d�a = 6 Or d�a = 7 Then
        btnD�a(0).ForeColor = &H4080&
    Else
        btnD�a(0).ForeColor = &H8000000F
    End If
End Sub

Private Function decodificarMes(numMes As String) As Byte
    Dim mes As Byte
    Select Case LCase(numMes)
        Case "enero"
            mes = 1
        Case "febrero"
            mes = 2
        Case "marzo"
            mes = 3
        Case "abril"
            mes = 4
        Case "mayo"
            mes = 5
        Case "junio"
            mes = 6
        Case "julio"
            mes = 7
        Case "agosto"
            mes = 8
        Case "setiembre"
            mes = 9
        Case "octubre"
            mes = 10
        Case "noviembre"
            mes = 11
        Case "diciembre"
            mes = 12
    End Select
    decodificarMes = mes
End Function

Sub llenarActividades(materia As String, mes As Byte, a�o As Integer)
    Dim j As Integer, archivolibre As Byte, n�meroD�a As Byte
    Dim d�as(1 To 31) As Integer
    
    For j = 1 To 31
        d�as(j) = 0
    Next
    
    File1.Refresh
    File1.Path = App.Path + "\trabajos\" + materia + "\actividades\" + Trim(Str(mes)) + "\"
    If File1.ListCount <> 0 Then
        For j = 0 To File1.ListCount - 1
            If Mid(File1.List(j), 7, 4) = a�o Then 'si es el mismo a�o
                    n�meroD�a = CByte(Mid(File1.List(j), 4, 2))
                    d�as(n�meroD�a) = d�as(n�meroD�a) + 1 'se cuentan cu�ntas activiades hay por d�a
            End If
        Next
        For j = 1 To 31
            If d�as(j) <> 0 Then 'se llenan los botones de los d�as con las actividades
                If d�as(j) = 1 Then
                    btnD�a(j - 1).Caption = btnD�a(j - 1).Caption + Chr(13) + Trim(Str(d�as(j))) + " actividad"
                Else
                    btnD�a(j - 1).Caption = btnD�a(j - 1).Caption + Chr(13) + Trim(Str(d�as(j))) + " actividades"
                End If
            End If
        Next
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    swEstoyAbierto = False
    'Call contarFormularios(False)
End Sub

Public Sub actualizarCalendario()
    Dim mes As Byte, d�a As Byte, a�o As Integer
    
    a�o = Combo2.Text
    mes = decodificarMes(Combo3.Text)
    d�a = nombreDeD�a(1, mes, a�o)
    Call chequearD�a1(d�a)
    Call cargarMesEnPantalla(mes, a�o, d�a)
    Call llenarActividades(Trim(Combo1.Text), mes, a�o)
    Call arreglarForm(mes, a�o)
End Sub
