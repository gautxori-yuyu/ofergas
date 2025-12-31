VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_conn As Object           ' ADODB.Connection
Private m_daoWs As Object          ' DAO.Workspace (para transacciones)

Public Property Get Connection() As Object
    Set Connection = m_conn
End Property

Public Property Get DAOWorkspace() As Object
    Set DAOWorkspace = m_daoWs
End Property

Public Sub Connect(ByVal fullPathDB As String)
    Set m_conn = CreateObject("ADODB.Connection")
    m_conn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & fullPathDB & ";"
    
    ' Inicializar DAO.Workspace (late binding)
    Set m_daoWs = CreateObject("DAO.DBEngine.120").Workspaces(0)
    Debug.Print "[OK] Conexión abierta"
End Sub

Public Sub Disconnect()
    If Not m_conn Is Nothing Then m_conn.Close
    Set m_conn = Nothing
    Set m_daoWs = Nothing
    Debug.Print "[OK] Conexión cerrada"
End Sub
