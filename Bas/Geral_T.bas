Attribute VB_Name = "Geral"
Option Explicit

'Strings da conexão
Public cn As ADODB.Connection
Public strBD As String

'String da Visualização
Global ValorTipo As String
Global ValorElemento As String
Global ValorLocal As String
Global ValorAtendente As String
Global ValorDescricao As String
Global ValorReporte As String
Global ValorResultado As String
Global ValorFinalizado As String
Global ValorCancelada As String


Global PermissaoAD As Boolean


Global sFullName As String
Global strDescription As String
Global VCompletoDN As String
Global vUsuario As String



'Strings do usuário
Global gintUsuarioID As Integer
Global gstrNome As String
Global gstrDepto As String
Global gstrEMail As String
Global gstrSenha As String
Global gstrTela As String
Global gblnEncaminharOS As Boolean
Global gblnSupervisor As Boolean

'Strings de OS
Global gintOSID As Long
Global VElemento As String



'Variáveis para criação de usuário
Global gintNovoUsuarioID As Integer

'Variáveis do frmChamados
'Combo status
Global gstrcboStatus As String
Global gblnTelaChamados As Boolean
Global gintStatusFinalizada As Integer

Public Sub ConectarBD()
    'Conexão com a Elétricos
     strBD = "PROVIDER=SQLOLEDB;SERVER=192.168.0.249\SQLEXPRESS;DATABASE=SugestoesMelhorias;UID=cablena_user;PWD=C@bl3na;"
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open (strBD)
End Sub

Public Function CodDec(ByVal vText As String) As String
On Error Resume Next
Dim strAux As String
Dim intCount As Long
Dim strText As String
    
    strText = vText
    strAux = ""
    For intCount = 1 To Len(strText)
        strAux = Chr((Asc(Mid(strText, intCount, 1)) Xor 255)) & strAux
    Next
    CodDec = strAux
End Function


