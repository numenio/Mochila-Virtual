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
   Begin TransparentButton.ButtonTransparent btnDía 
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
      Caption         =   "Año:"
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
Dim swEmpezó As Boolean
Public swEstoyAbierto As Boolean

Private Sub btnDía_Click(Index As Integer)
    Dim mes As Date
    
    mes = Format(Str(Index + 1) + "/" + Trim(Str(decodificarMes(Combo3.Text))) + "/" + Combo2.Text)
    frmAñadirActividad.díaParaCargarActividades = mes
    frmAñadirActividad.materia = Combo1.Text
    frmAñadirActividad.swCargarFecha = True
    If Not IsNumeric(btnDía(Index).Caption) Then
        frmAñadirActividad.swEditarActividades = True
'        frmAñadirActividad.Show 1
    Else
        frmAñadirActividad.swEditarActividades = False
    End If
    frmAñadirActividad.Show 1 'añadir una actividad
End Sub

Private Sub Combo1_Click()
    Dim mes As Byte, día As Byte, año As Integer
    
    If swEmpezó = True Then
        año = Combo2.Text
        mes = decodificarMes(Combo3.Text)
        día = nombreDeDía(1, mes, año)
        Call chequearDía1(día)
        Call cargarMesEnPantalla(mes, año, día)
        Call llenarActividades(Trim(Combo1.Text), mes, año)
        Call arreglarForm(mes, año)
    End If
End Sub

Private Sub Combo2_Click() 'si se cambian los años
    Dim mes As Byte, día As Byte, año As Integer
    
    If swEmpezó = True Then
        año = Combo2.Text
        mes = decodificarMes(Combo3.Text)
        día = nombreDeDía(1, mes, año)
        Call chequearDía1(día)
        Call cargarMesEnPantalla(mes, año, día)
        Call llenarActividades(Trim(Combo1.Text), mes, año)
        Call arreglarForm(mes, año)
    End If
End Sub

Private Sub Combo3_Click() 'si se cambian los meses
    Dim mes As Byte, día As Byte, año As Integer
    
    If swEmpezó = True Then
        If Combo2.Text <> "" Then
            año = Combo2.Text
            mes = decodificarMes(Combo3.Text)
            día = nombreDeDía(1, mes, año)
            Call chequearDía1(día)
            Call cargarMesEnPantalla(mes, año, día)
            Call llenarActividades(Trim(Combo1.Text), mes, año)
            Call arreglarForm(mes, año)
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
       
    swEmpezó = False
    swEstoyAbierto = True
    
    For j = 2008 To Year(Date) + 10 'se llenan los años, desde el actual más otros 10 años
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
    
    Dim añoABuscar As Integer 'se activa el año actual
    añoABuscar = Year(Date)
    For j = 0 To Combo2.ListCount - 1
        If Combo2.List(j) = añoABuscar Then
            Combo2.ListIndex = j
            Exit For
        End If
    Next
    
    Combo1.ListIndex = 0 'se activa la primer materia
    
    btnDía(0).Caption = "1"
    For j = 1 To 30
        Load btnDía(j)
        btnDía(j).Caption = Str(j + 1) 'se carga en el botón 0 el día 1
        btnDía(j).Visible = True
    Next
    
    For j = 1 To 6
        Load Label2(j)
    Next
    
    Dim día As Byte, año As Integer
    
    año = Combo2.Text
    mes = decodificarMes(Combo3.Text)
    día = nombreDeDía(1, mes, año)
    Call chequearDía1(día)
    Call cargarMesEnPantalla(mes, año, día)
    Call llenarActividades(Trim(Combo1.Text), mes, año)
    Call arreglarForm(mes, año)
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    swEmpezó = True
End Sub

Sub cargarMesEnPantalla(mes As Byte, año As Integer, dóndeEmpezar As Byte)
    Dim j As Integer, espacioEntreBotones As Integer, espacioHorizontalbotones As Byte
    Dim col1, numColumna As Byte, numFila As Byte, fila1X As Integer, fila1Y As Integer
       
    espacioEntreBotones = 200
    espacioHorizontalbotones = 200
    
    btnDía(0).Caption = 1
    
    For j = 1 To 30
        btnDía(j).Visible = False
        btnDía(j).ForeColor = &H8000000F 'blanco
        btnDía(j).Caption = j + 1
    Next
        
    With Label2(0)
        .Width = btnDía(0).Width
        .Height = btnDía(0).Height
        .Top = 1480
        .Left = 480
        .Visible = True
        .Caption = "L"
    End With
    
    numFila = 1
    
    fila1X = 480
    fila1Y = 2200
    
    btnDía(0).Top = fila1Y
    btnDía(0).Left = CInt(fila1X) + btnDía(0).Width * (CInt(dóndeEmpezar) - 1) + CInt(espacioHorizontalbotones) * (CInt(dóndeEmpezar) - 1) '+ espacioHorizontalbotones
    
    dóndeEmpezar = dóndeEmpezar + 1
    
    For j = 1 To cantDíasMes(mes, año) - 1
        If dóndeEmpezar = 8 Then
            numFila = numFila + 1
            btnDía(j).Left = fila1X
            dóndeEmpezar = 1
        Else
            btnDía(j).Left = btnDía(j - 1).Left + espacioHorizontalbotones + btnDía(j).Width
        End If
        
        If numFila = 1 Then
            btnDía(j).Top = fila1Y
        Else
            btnDía(j).Top = fila1Y + btnDía(j).Width * (numFila - 1)
        End If
        
        If dóndeEmpezar = 6 Or dóndeEmpezar = 7 Then btnDía(j).ForeColor = &H4080& 'marrón
        
        
        dóndeEmpezar = dóndeEmpezar + 1
        btnDía(j).Visible = True
        btnDía(j).Refresh
        
        If j < 7 Then 'se cargan los días, desde el martes al domingo
            With Label2(j)
                .Width = btnDía(0).Width
                .Height = btnDía(0).Height
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
        
    btnDía(0).Refresh
    
    With Line1
        .X1 = fila1X
        .X2 = Label2(6).Left + btnDía(0).Width + espacioHorizontalbotones
        .Y1 = fila1Y - espacioHorizontalbotones
        .Y2 = fila1Y - espacioHorizontalbotones
    End With
End Sub

Sub arreglarForm(mes As Byte, año As Integer)
    Dim díasDelMes As Byte
    díasDelMes = cantDíasMes(mes, año) '-1
    Me.Height = btnDía(díasDelMes - 1).Top + btnDía(díasDelMes - 1).Height + 1000
End Sub

Sub chequearDía1(día As Byte)
    If día = 6 Or día = 7 Then
        btnDía(0).ForeColor = &H4080&
    Else
        btnDía(0).ForeColor = &H8000000F
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

Sub llenarActividades(materia As String, mes As Byte, año As Integer)
    Dim j As Integer, archivolibre As Byte, númeroDía As Byte
    Dim días(1 To 31) As Integer
    
    For j = 1 To 31
        días(j) = 0
    Next
    
    File1.Refresh
    File1.Path = App.Path + "\trabajos\" + materia + "\actividades\" + Trim(Str(mes)) + "\"
    If File1.ListCount <> 0 Then
        For j = 0 To File1.ListCount - 1
            If Mid(File1.List(j), 7, 4) = año Then 'si es el mismo año
                    númeroDía = CByte(Mid(File1.List(j), 4, 2))
                    días(númeroDía) = días(númeroDía) + 1 'se cuentan cuántas activiades hay por día
            End If
        Next
        For j = 1 To 31
            If días(j) <> 0 Then 'se llenan los botones de los días con las actividades
                If días(j) = 1 Then
                    btnDía(j - 1).Caption = btnDía(j - 1).Caption + Chr(13) + Trim(Str(días(j))) + " actividad"
                Else
                    btnDía(j - 1).Caption = btnDía(j - 1).Caption + Chr(13) + Trim(Str(días(j))) + " actividades"
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
    Dim mes As Byte, día As Byte, año As Integer
    
    año = Combo2.Text
    mes = decodificarMes(Combo3.Text)
    día = nombreDeDía(1, mes, año)
    Call chequearDía1(día)
    Call cargarMesEnPantalla(mes, año, día)
    Call llenarActividades(Trim(Combo1.Text), mes, año)
    Call arreglarForm(mes, año)
End Sub
