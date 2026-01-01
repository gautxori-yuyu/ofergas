VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' clsDBManager - Gestión de Base de Datos como Entidad
' ============================================================================
' Responsabilidades:
' - Conexión/desconexión (ADODB, DAO)
' - Transacciones (BeginTrans, Commit, Rollback vía DAO.Workspace)
' - Backup/Restore de archivos de base de datos
' - Gestión de tecnología de acceso a datos (DAO/ADO)
' ============================================================================
Option Explicit

Private m_conn As Object           ' ADODB.Connection
Private m_daoWs As Object          ' DAO.Workspace (para transacciones)
Private m_dbPath As String         ' Ruta completa a la BD

' ============================================================================
' PROPIEDADES
' ============================================================================

Public Property Get Connection() As Object
    Set Connection = m_conn
End Property

Public Property Get DAOWorkspace() As Object
    Set DAOWorkspace = m_daoWs
End Property

Public Property Get DatabasePath() As String
    DatabasePath = m_dbPath
End Property

' ============================================================================
' CONEXIÓN Y DESCONEXIÓN
' ============================================================================

Public Sub Connect(ByVal fullPathDB As String)
    On Error GoTo ErrHandler

    m_dbPath = fullPathDB

    ' Conexión ADODB (para consultas SQL)
    Set m_conn = CreateObject("ADODB.Connection")
    m_conn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & fullPathDB & ";"

    ' Workspace DAO (para transacciones)
    Set m_daoWs = CreateObject("DAO.DBEngine.120").Workspaces(0)

    Debug.Print "[OK] Conexión abierta: " & fullPathDB
    Exit Sub

ErrHandler:
    Debug.Print "[ERR] Connect: " & Err.Description
    Err.Raise Err.Number, "clsDBManager.Connect", Err.Description
End Sub

Public Sub Disconnect()
    On Error Resume Next

    If Not m_conn Is Nothing Then
        m_conn.Close
        Set m_conn = Nothing
    End If

    Set m_daoWs = Nothing
    m_dbPath = ""

    Debug.Print "[OK] Conexión cerrada"
End Sub

' ============================================================================
' BACKUP Y RESTORE
' ============================================================================

Public Function Backup(ByVal backupPath As String) As Boolean
    ' Crea backup de la BD actual
    On Error GoTo ErrHandler

    If m_dbPath = "" Then
        Debug.Print "[ERR] Backup: No hay base de datos conectada"
        Backup = False
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(m_dbPath) Then
        Debug.Print "[ERR] Backup: Base de datos origen no existe: " & m_dbPath
        Backup = False
        Exit Function
    End If

    ' Crear directorio de destino si no existe
    Dim backupDir As String
    backupDir = fso.GetParentFolderName(backupPath)
    If backupDir <> "" And Not fso.FolderExists(backupDir) Then
        fso.CreateFolder backupDir
    End If

    ' Copiar archivo
    fso.CopyFile m_dbPath, backupPath, True ' Overwrite = True

    Set fso = Nothing

    Debug.Print "[OK] Backup creado: " & backupPath
    Backup = True
    Exit Function

ErrHandler:
    Debug.Print "[ERR] Backup: " & Err.Description
    Backup = False
End Function

Public Function BackupTo(ByVal sourceDBPath As String, ByVal backupPath As String) As Boolean
    ' Crea backup de una BD específica (sin necesidad de estar conectado)
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(sourceDBPath) Then
        Debug.Print "[ERR] BackupTo: Base de datos origen no existe: " & sourceDBPath
        BackupTo = False
        Exit Function
    End If

    ' Crear directorio de destino si no existe
    Dim backupDir As String
    backupDir = fso.GetParentFolderName(backupPath)
    If backupDir <> "" And Not fso.FolderExists(backupDir) Then
        fso.CreateFolder backupDir
    End If

    ' Copiar archivo
    fso.CopyFile sourceDBPath, backupPath, True ' Overwrite = True

    Set fso = Nothing

    Debug.Print "[OK] Backup creado: " & backupPath
    BackupTo = True
    Exit Function

ErrHandler:
    Debug.Print "[ERR] BackupTo: " & Err.Description
    BackupTo = False
End Function

Public Function Restore(ByVal backupPath As String, ByVal destDBPath As String) As Boolean
    ' Restaura BD desde backup
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(backupPath) Then
        Debug.Print "[ERR] Restore: Backup no existe: " & backupPath
        Restore = False
        Exit Function
    End If

    ' Advertencia: esto sobrescribirá la base de datos de destino
    If fso.FileExists(destDBPath) Then
        Debug.Print "[WARN] Restore: Sobrescribiendo base de datos existente: " & destDBPath
    End If

    ' Crear directorio de destino si no existe
    Dim destDir As String
    destDir = fso.GetParentFolderName(destDBPath)
    If destDir <> "" And Not fso.FolderExists(destDir) Then
        fso.CreateFolder destDir
    End If

    ' Copiar archivo
    fso.CopyFile backupPath, destDBPath, True ' Overwrite = True

    Set fso = Nothing

    Debug.Print "[OK] Base de datos restaurada desde: " & backupPath
    Restore = True
    Exit Function

ErrHandler:
    Debug.Print "[ERR] Restore: " & Err.Description
    Restore = False
End Function

' ============================================================================
' GESTIÓN DE TRANSACCIONES (a través de DAO.Workspace)
' ============================================================================

Public Sub BeginTransaction()
    On Error GoTo ErrHandler
    If Not m_daoWs Is Nothing Then
        m_daoWs.BeginTrans
        Debug.Print "[OK] Transacción iniciada"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERR] BeginTransaction: " & Err.Description
End Sub

Public Sub CommitTransaction()
    On Error GoTo ErrHandler
    If Not m_daoWs Is Nothing Then
        m_daoWs.CommitTrans
        Debug.Print "[OK] Transacción confirmada"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERR] CommitTransaction: " & Err.Description
End Sub

Public Sub RollbackTransaction()
    On Error GoTo ErrHandler
    If Not m_daoWs Is Nothing Then
        m_daoWs.Rollback
        Debug.Print "[WARN] Transacción revertida"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERR] RollbackTransaction: " & Err.Description
End Sub

' ============================================================================
' INFORMACIÓN DE BASE DE DATOS
' ============================================================================

Public Function IsConnected() As Boolean
    IsConnected = Not (m_conn Is Nothing)
End Function

Public Function GetDatabaseInfo() As Object
    ' Retorna Dictionary con información de la BD
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")

    info.Add "Connected", IsConnected()
    info.Add "Path", m_dbPath
    info.Add "Provider", IIf(IsConnected(), m_conn.Provider, "")

    Set GetDatabaseInfo = info
End Function
