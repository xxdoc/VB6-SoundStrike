VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound-Strike"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOFF 
      Caption         =   "OFF"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5910
      TabIndex        =   15
      Top             =   5580
      Width           =   645
   End
   Begin VB.CommandButton cmdON 
      Caption         =   "ON"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5280
      TabIndex        =   14
      Top             =   5580
      Width           =   645
   End
   Begin VB.Frame frmConfiguracion 
      Caption         =   "Configuraci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   7335
      Begin VB.TextBox txtNombreFichero 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "voice_input.wav"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtTecla 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "120"
         Text            =   "F9"
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtRuta 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Ruta Sonidos"
         Top             =   300
         Width           =   5055
      End
      Begin VB.Image imgVerde 
         Height          =   480
         Left            =   4560
         Picture         =   "frmMain.frx":1E72
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgRojo 
         Height          =   480
         Left            =   4080
         Picture         =   "frmMain.frx":22B4
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgSemaforo 
         Height          =   480
         Left            =   6600
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblNombreFichero 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Fichero:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblTecla 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tecla de acci�n:"
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblRuta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ruta general:"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame frmExplorador 
      Caption         =   "Explorador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   3615
      Begin VB.FileListBox filArchivos 
         Height          =   2430
         Left            =   120
         Pattern         =   "*.wav"
         TabIndex        =   2
         Top             =   2160
         Width           =   3375
      End
      Begin VB.DirListBox dirDirectorios 
         Height          =   1440
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   3375
      End
      Begin VB.DriveListBox drvUnidades 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblAuxiliar 
         AutoSize        =   -1  'True
         Caption         =   "Etiqueta Auxiliar"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame frmLista 
      Caption         =   "Lista de Sonidos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin VB.ListBox lstSonidos 
         Height          =   4350
         ItemData        =   "frmMain.frx":26F6
         Left            =   120
         List            =   "frmMain.frx":26F8
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Menu menArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu menCargar 
         Caption         =   "&Cargar Configuraci�n"
      End
      Begin VB.Menu menGuardar 
         Caption         =   "&Guardar Configuraci�n"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu menAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu menGuia 
         Caption         =   "&Gu�a del usuario"
      End
      Begin VB.Menu menAcerca 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rutaEjecucion As String          'Directorio de trabajo en el que se ejecuta la aplicaci�n.
Dim indiceSonido  As Integer         '�ndice del sonido a renombrar.
Dim nombreTemp    As String          'Nombre temporal del fichero a renombrar.
Const carpetaConfigs = "configs"     'Nombre de la carpeta donde se almacenan las configuraciones.
Const extensionFichero = ".cfg"      'Extensi�n de los ficheros de configuraci�n.

Private Sub cmdA�adir_Click()
    Dim i As Integer
    Dim existe As Boolean
    
    existe = False
    
    On Error Resume Next
    If filArchivos.ListIndex >= 0 Then     'Se a�aden los ficheros a la lista.
        For i = 0 To lstSonidos.ListCount - 1     'No se permiten sonidos repetidos.
            If lstSonidos.List(i) = filArchivos.List(filArchivos.ListIndex) Then
                existe = True
                Debug.Print "Sonido existente"
            End If
        Next i
        If existe = False Then
            lstSonidos.AddItem (filArchivos.List(filArchivos.ListIndex))
            cmdON.Enabled = True
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    On Error Resume Next
    lstSonidos.RemoveItem lstSonidos.ListIndex
    If lstSonidos.ListCount = 0 Then     'Si no quedan elementos en la lista, se inhabilita el bot�n eliminar y el de activado.
        cmdON.Enabled = False
    End If
    
    Exit Sub
End Sub

Private Sub cmdOFF_Click()
    Name txtRuta.Text + "\" + txtNombreFichero As txtRuta.Text + "\" + nombreTemp    'Se renombra el fichero input_voice.wav al �ltimo origninal.
    nombreTemp = ""
    lstSonidos.Enabled = True          'Se habilitan los controles.
    lstSonidos.ListIndex = -1
    drvUnidades.Enabled = True
    dirDirectorios.Enabled = True
    filArchivos.Enabled = True
    frmConfiguracion.Enabled = True
    imgSemaforo.Picture = imgRojo.Picture
    cmdON.Enabled = True
    cmdOFF.Enabled = False
    UnHookKeyB     'Se deshabilita el gancho del teclado.
    
    Exit Sub
End Sub

Private Sub cmdON_Click()
    lstSonidos.Enabled = False         'Se inhabilitan los controles.
    drvUnidades.Enabled = False
    dirDirectorios.Enabled = False
    filArchivos.Enabled = False
    frmConfiguracion.Enabled = False
    imgSemaforo.Picture = imgVerde.Picture
    cmdON.Enabled = False
    cmdOFF.Enabled = True
    lstSonidos.ListIndex = 0
    renombrarFichero                 'Se comienza renombrando un fichero.
    HookKeyB App.hInstance           'Se habilita el gancho del teclado.
    
    Exit Sub
End Sub

Public Sub renombrarFichero()
    On Error Resume Next
    If lstSonidos.ListCount > 0 Then     'En caso de no haber sonidos agregados a la lista, no se renombra nada.
        indiceSonido = lstSonidos.ListIndex
        
        If nombreTemp <> "" Then     'Si ya existe un nombre temporal, es porque se renombr� ya el fichero.
            Name txtRuta.Text + "\" + txtNombreFichero As txtRuta.Text + "\" + nombreTemp
        End If
        
        nombreTemp = lstSonidos.List(indiceSonido)
        
        Name txtRuta.Text + "\" + nombreTemp As txtRuta.Text + "\" + txtNombreFichero     'Se renombra el fichero.
        
        indiceSonido = indiceSonido + 1
        lstSonidos.ListIndex = lstSonidos.ListIndex + 1
        
        If indiceSonido = (lstSonidos.ListCount) Then
            indiceSonido = 0
            lstSonidos.ListIndex = 0
        End If
    End If
End Sub

Private Sub dirDirectorios_Change()
    On Error Resume Next
    filArchivos.Path = dirDirectorios.Path     'Se asigna la ruta de la lista de directorios a la lista de ficheros.
    ChDir filArchivos.Path                     'Se cambia el directorio de trabajo del sistema al seleccionado.
    txtRuta = CurDir(drvUnidades.List(drvUnidades.ListIndex))     'Se asigna el directorio de trabajo del sistema al campo de texto de la ruta.
End Sub

Private Sub drvUnidades_Change()
    On Error Resume Next
    dirDirectorios.Path = drvUnidades.List(drvUnidades.ListIndex)     'Se asigna la ruta de la lista de directorios a la unidad seleccionada.
End Sub

Private Sub filArchivos_DblClick()
    cmdA�adir_Click     'Al hacer doble clic, se simula el clic del bot�n a�adir.
End Sub

'Al hacer clic sobre la lista de ficheros, se simula su arrastre con una etiqueta oculta.
Private Sub filArchivos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If filArchivos.ListCount > 0 Then
        lblAuxiliar.Move filArchivos.Left, filArchivos.Top + Y - CLng(TextHeight("X")) / 2, filArchivos.Width, CLng(TextHeight("X"))
        lblAuxiliar.Drag
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = App.EXEName + "   v." + CStr(App.Major) + "." + CStr(App.Minor)
    rutaEjecucion = CurDir     'Al cargar el formulario se asigna la ruta de ejecuci�n al directorio actual de trabajo
    dirDirectorios_Change      'y se simula un cambio de la lista de directorios.
    imgSemaforo.Picture = imgRojo.Picture
    banderaTecla = False       'Se inicia la bandera a FALSE, ya que se pondr� a TRUE cuando se pulse una tecla.
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("�Deseas salir?", vbQuestion + vbYesNo, "Salir") = vbNo Then
        Cancel = True
    Else
        If cmdOFF.Enabled = True Then     'Por seguridad, si el bot�n OFF est� habilitado, se ejecuta su deshabilitaci�n.
            cmdOFF_Click
        End If
        UnHookKeyB     'Por seguridad, se quita el gancho al teclado.
        End
    End If
End Sub

Private Sub lstSonidos_DblClick()
    cmdEliminar_Click     'Al hacer doble clic, se simula el clic sobre el bot�n eliminar.
End Sub

Private Sub lstSonidos_DragDrop(Source As Control, X As Single, Y As Single)
    cmdA�adir_Click     'Al hacer un arrastre de un elemento de la lista de ficheros sobre la lista de sonidos, se simula el clic sobre el bot�n a�adir.
End Sub

Private Sub menAcerca_Click()
    frmAcerca.Show vbModal
End Sub

Private Sub menCargar_Click()
    On Error GoTo errorCarga

    Dim nombreFichero As String
    Dim rutaTotal     As String
    Dim i             As Integer
    Dim numeroSonidos As Integer

    nombreFichero = InputBox("Introduce un nombre para CARGAR la configuraci�n:", "Cargar Configuraci�n")     'Se obtiene el nombre del fichero a cargar de una ventana de inserci�n de datos.
    
    If nombreFichero <> "" Then
        rutaTotal = rutaEjecucion + "\" + carpetaConfigs + "\" + nombreFichero + extensionFichero
        If ExisteArchivo(rutaTotal) = True Then
            drvUnidades.Drive = LeeINI(rutaTotal, "EXPLORER", "Drive")          'Unidad de los sonidos.
            dirDirectorios.Path = LeeINI(rutaTotal, "EXPLORER", "DirPath")      'Directorio de los sonidos.
            filArchivos.Path = LeeINI(rutaTotal, "EXPLORER", "FilPath")         'Directorio de los sonidos.
            txtRuta.Text = LeeINI(rutaTotal, "CONFIG", "Path")                  'Ruta total de los sonidos.
            txtTecla.Text = LeeINI(rutaTotal, "CONFIG", "Hotkey")               'Tecla de acci�n.
            txtTecla.Tag = LeeINI(rutaTotal, "CONFIG", "Hotkeycode")            'C�digo de la tecla de acci�n.
            txtNombreFichero.Text = LeeINI(rutaTotal, "CONFIG", "Filename")     'Nombre del fichero a sustitu�r.
            numeroSonidos = CLng(LeeINI(rutaTotal, "LIST", "Count"))            'N�mero total de sonidos de la lista.
            lstSonidos.Clear
            For i = 0 To numeroSonidos - 1
                lstSonidos.AddItem LeeINI(rutaTotal, "LIST", "Item" + CStr(i))  'Cada elemento de la lista de sonidos.
            Next i
            'Se habilitan o deshabilitan los controles en funci�n de si la lista de sonidos contiene elementos.
            If numeroSonidos > 0 Then     'La lista contiene elementos.
                cmdON.Enabled = True
                cmdOFF.Enabled = False
            Else                                 'La lista no contiene elementos.
                cmdON.Enabled = False
                cmdOFF.Enabled = False
            End If
            MsgBox "Configuraci�n '" + nombreFichero + extensionFichero + "' cargada.", vbInformation, "Configuraci�n"
        Else
            MsgBox "Configuraci�n no encontrada en '.\" + carpetaConfigs + "\'", vbCritical, "Configuraci�n no encontrada"
        End If
    Else
        'MsgBox "Nombre de fichero no v�lido.", vbCritical, "Error"
    End If

    
    Exit Sub

errorCarga:
    MsgBox "Se ha producido un error al cargar la configuraci�n." + vbCrLf + _
           "Es posible que el fichero est� corrupto, o que no se encuentre la carpeta con los sonidos.", _
           vbCritical, "Configuraci�n"
    Exit Sub
End Sub

Private Sub menGuardar_Click()
    On Error GoTo errorGuardar

    Dim nombreFichero As String
    Dim rutaTotal     As String
    Dim i             As Integer
    
    nombreFichero = InputBox("Introduce un nombre para GUARDAR la configuraci�n:", "Guardar Configuraci�n")     'Se obtiene el nombre del fichero a guardar de una ventana de inserci�n de datos.
    
    If nombreFichero <> "" Then
        If ExisteArchivo(rutaEjecucion + "\" + carpetaConfigs + "\") = False Then
            MkDir rutaEjecucion + "\" + carpetaConfigs
        End If
        rutaTotal = rutaEjecucion + "\" + carpetaConfigs + "\" + nombreFichero + extensionFichero
        GrabaINI rutaTotal, "EXPLORER", "Drive", drvUnidades.Drive               'Unidad.
        GrabaINI rutaTotal, "EXPLORER", "DirPath", dirDirectorios.Path           'Directorio.
        GrabaINI rutaTotal, "EXPLORER", "FilPath", filArchivos.Path              'Directorio.
        GrabaINI rutaTotal, "CONFIG", "Path", txtRuta.Text                       'Ruta total.
        GrabaINI rutaTotal, "CONFIG", "Hotkey", txtTecla.Text                    'Tecla de acci�n.
        GrabaINI rutaTotal, "CONFIG", "Hotkeycode", txtTecla.Tag                 'C�digo de la tecla de acci�n.
        GrabaINI rutaTotal, "CONFIG", "Filename", txtNombreFichero.Text          'Nombre del fichero a sustitu�r.
        GrabaINI rutaTotal, "LIST", "Count", lstSonidos.ListCount                'N�mero total de sonidos.
        For i = 0 To lstSonidos.ListCount - 1
            GrabaINI rutaTotal, "LIST", "Item" + CStr(i), lstSonidos.List(i)     'Cada elemento de la lista.
        Next i
        MsgBox "Configuraci�n guardada en '" + nombreFichero + extensionFichero + "'.", vbInformation, "Configuraci�n"
    Else
        'MsgBox "Nombre de fichero no v�lido.", vbCritical, "Error"
    End If
    
    Exit Sub

errorGuardar:
    MsgBox "Se ha producido un error al guardar la configuraci�n." + vbCrLf + _
           "Por favor, cierra la aplicaci�n e int�ntalo de nuevo." + vbCrLf + _
           "Si el error persiste, ponte en contacto con el programador.", _
           vbCritical, "Configuraci�n"
    Exit Sub
End Sub

Private Sub menGuia_Click()
    frmGuia.Show vbModal
End Sub

Private Sub menSalir_Click()
    Unload Me
End Sub

Private Sub txtTecla_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode     'En funci�n de la tecla pulsada, se captura el c�digo de tecla y su valor textual.
        Case vbKeyF1
            txtTecla.Text = "F1"
            
        Case vbKeyF2
            txtTecla.Text = "F2"
            
        Case vbKeyF3
            txtTecla.Text = "F3"
            
        Case vbKeyF4
            txtTecla.Text = "F4"
            
        Case vbKeyF5
            txtTecla.Text = "F5"
            
        Case vbKeyF6
            txtTecla.Text = "F6"
            
        Case vbKeyF7
            txtTecla.Text = "F7"
            
        Case vbKeyF8
            txtTecla.Text = "F8"
            
        Case vbKeyF9
            txtTecla.Text = "F9"
            
        Case vbKeyF10
            txtTecla.Text = "F10"
            
        Case vbKeyF11
            txtTecla.Text = "F11"
            
        Case vbKeyF12
            txtTecla.Text = "F12"
            
        Case vbKeyNumpad0
            txtTecla.Text = "0"
            
        Case vbKeyNumpad1
            txtTecla.Text = "1"
            
        Case vbKeyNumpad2
            txtTecla.Text = "2"
            
        Case vbKeyNumpad3
            txtTecla.Text = "3"
            
        Case vbKeyNumpad4
            txtTecla.Text = "4"
            
        Case vbKeyNumpad5
            txtTecla.Text = "5"
            
        Case vbKeyNumpad6
            txtTecla.Text = "6"
            
        Case vbKeyNumpad7
            txtTecla.Text = "7"
            
        Case vbKeyNumpad8
            txtTecla.Text = "8"
            
        Case vbKeyNumpad9
            txtTecla.Text = "9"
            
        Case vbKeyAdd
            txtTecla.Text = "+"
            
        Case vbKeyBack
            txtTecla.Text = "RETROCESO"
            
        Case vbKeyCapital
            txtTecla.Text = "BLOQ MAY�S"
            
        Case vbKeyClear
            txtTecla.Text = "BORRAR"
            
        Case vbKeyControl
            txtTecla.Text = "CONTROL"
            
        Case vbKeyDecimal
            txtTecla.Text = "."
            
        Case vbKeyDelete
            txtTecla.Text = "SUPRIMIR"
            
        Case vbKeyDivide
            txtTecla.Text = "/"
            
        Case vbKeyUp
            txtTecla.Text = "ARRIBA"
            
        Case vbKeyDown
            txtTecla.Text = "ABAJO"
            
        Case vbKeyRight
            txtTecla.Text = "DERECHA"
            
        Case vbKeyLeft
            txtTecla.Text = "IZQUIERDA"
            
        Case vbKeyEnd
            txtTecla.Text = "FIN"
            
        Case vbKeyEscape
            txtTecla.Text = "ESCAPE"
            
        Case vbKeyHome
            txtTecla.Text = "INICIO"
            
        Case vbKeyInsert
            txtTecla.Text = "INSERTAR"
            
        Case vbKeyMenu
            txtTecla.Text = "ALT"
            
        Case vbKeyMultiply
            txtTecla.Text = "*"
            
        Case vbKeyNumlock
            txtTecla.Text = "BLOQ NUM"
            
        Case vbKeyPageUp
            txtTecla.Text = "RE P�G"
            
        Case vbKeyPageDown
            txtTecla.Text = "AV P�G"
            
        Case vbKeyPause
            txtTecla.Text = "PAUSA"
            
        Case vbKeyPrint
            txtTecla.Text = "IMPR PANT"
            
        Case vbKeyReturn
            txtTecla.Text = "ENTER"
            
        Case vbKeySelect
            txtTecla.Text = "SELECT"
            
        Case vbKeyShift
            txtTecla.Text = "SHIFT"
            
        Case vbKeySpace
            txtTecla.Text = "ESPACIO"
            
        Case vbKeySubtract
            txtTecla.Text = "-"
            
        Case vbKeyTab
            txtTecla.Text = "TABULADOR"
            
        Case Else
            txtTecla.Text = Chr(KeyCode)
    End Select
    txtTecla.Tag = CStr(KeyCode)
End Sub

Public Sub trazaError(descripcionError As String)
    Dim numeroFichero As Integer
    Const nombreFichero = "error.log"
    
    'Se asigna un numero para el fichero a abrir (sistema).
    numeroFichero = FreeFile
    Open nombreFichero For Append As #numeroFichero
        Print #numeroFichero, descripcionError
    Close #numeroFichero
    
    MsgBox "Se ha producido un error inesperado." + vbCrLf + _
           "Por favor, env�e el fichero 'error.log' al programador.", vbCritical, "Error inesperado"
End Sub
