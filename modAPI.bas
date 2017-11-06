Attribute VB_Name = "modAPI"
Option Explicit

'Se guarda el gancho creado con SetWindowsHookEx.
Private mHook As Long

'Variable que sirve de bandera para evitar que se detecten una pulsación como doble.
Public banderaTecla As Boolean

'Se indica a SetWindowsHookEx qué tipo de gancho queremos instalar (TECLADO).
Private Const WH_KEYBOARD_LL As Long = 13&
'Este es para el ratón.
'Private Const WH_MOUSE_LL As Long = 14&

Private Type tagKBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Const VK_TAB As Long = &H9          'TAB
Private Const VK_CONTROL As Long = &H11     'CONTROL
Private Const VK_MENU As Long = &H12        'ALT
Private Const VK_ESCAPE As Long = &H1B      'ESCAPE
Private Const VK_DELETE As Long = &H2E      'SUPR

Private Const LLKHF_ALTDOWN As Long = &H20&

'Códigos para los ganchos, es decir, la acción a tomar en el gancho del teclado.
Private Const HC_ACTION As Long = 0&

'Se asigna un gancho.
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long

'Se quita el gancho creado con SetWindowsHookEx.
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'Se llama al siguiente gancho.
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Se sabe si se ha pulsado una tecla.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Se copia la estructura en un long.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Funciones INI.
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Lee del fichero y sección que se le pasa como parámetros una clave.
Public Function LeeINI(archivo As String, Seccion As String, Clave As String) As String
    Dim iRetLen As Integer
    Dim sRet As String
    
    '255 espacios en blanco.
    sRet = Space(255)
    
    'Llama a una función de la API de acceso a archivos INI.
    iRetLen = GetPrivateProfileString(Seccion, Clave, "", sRet, Len(sRet), archivo)
    
    'Recorta los "iRetLen" primeros caracteres de "sRet".
    sRet = Left(sRet, iRetLen)
    
    'Devuelve el valor buscado.
    LeeINI = sRet
End Function

'Escribe una clave en función del fichero y la sección que se le pasa como parámetros.
Public Sub GrabaINI(archivo As String, Seccion As String, Clave As String, Text As String)
    'Llama a una función de la API de acceso a archivos INI.
    WritePrivateProfileString Seccion, Clave, Text, archivo
End Sub

'Comprueba que existe el archivo que se le pasa como parámetro.
Public Function ExisteArchivo(cArchivo As String) As Boolean
    'Devuelve si existe el fichero o no.
    ExisteArchivo = IIf(Dir$(cArchivo) = "", False, True)
End Function

'Función para usar el gancho del teclado.
Public Function LLKeyBoardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pkbhs As tagKBDLLHOOKSTRUCT
    Dim ret As Long
    
    ret = 0
    
    'Se copian los parámetros en la estructura.
    CopyMemory pkbhs, ByVal lParam, Len(pkbhs)
    
    If nCode = HC_ACTION Then
        If pkbhs.vkCode = frmMain.txtTecla.Tag Then
            If banderaTecla = False Then
                frmMain.renombrarFichero
                ret = 1
                banderaTecla = True
            Else
                banderaTecla = False
            End If
        End If
    End If
    
    If ret = 0 Then
        ret = CallNextHookEx(mHook, nCode, wParam, lParam)
    End If
    
    LLKeyBoardProc = ret
End Function

Public Sub HookKeyB(ByVal hMod As Long)
    'Se instala el gancho para el teclado.
    'hMod será el valor de App.hInstance de la aplicación.
    mHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LLKeyBoardProc, hMod, 0&)
End Sub

Public Sub UnHookKeyB()
    'Se desinstala el gancho para el teclado.
    If mHook <> 0 Then
        UnhookWindowsHookEx mHook
    End If
End Sub

