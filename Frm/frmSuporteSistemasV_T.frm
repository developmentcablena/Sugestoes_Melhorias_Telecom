VERSION 5.00
Begin VB.Form frmSuporteSistemasV 
   Caption         =   "Suporte Simac - Ordem de Atendimento (Visualização)"
   ClientHeight    =   9735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445
   Icon            =   "frmSuporteSistemasV_T.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAtendimentoConcluido 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   7560
      Width           =   8175
   End
   Begin VB.TextBox txtPropostaMelhoria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3720
      Width           =   8175
   End
   Begin VB.ComboBox cboTipo 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboElemento 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   5175
   End
   Begin VB.TextBox txtObservacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1800
      Width           =   8175
   End
   Begin VB.CommandButton cmdCadastrar 
      Appearance      =   0  'Flat
      Caption         =   "&Cadastrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   9240
      Width           =   1335
   End
   Begin VB.TextBox txtReporteTecnico 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5640
      Width           =   8175
   End
   Begin VB.ComboBox cboAtendente 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSuporteSistemasV_T.frx":12CA
      Left            =   2760
      List            =   "frmSuporteSistemasV_T.frx":12CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1035
      Width           =   5175
   End
   Begin VB.ComboBox cboLocal 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Técnico - Atendimento Concluído:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Width           =   4095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   60
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Proposta de Melhoria e Resultado Esperado:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   4305
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Elemento 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Elemento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblDescricaoServico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição da Situação Atual:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2790
   End
   Begin VB.Label lblReporteTecnico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Tecnico:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Label lblAtendente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   795
      Width           =   1005
   End
   Begin VB.Label Local 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   510
   End
End
Attribute VB_Name = "frmSuporteSistemasV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private Sub cmdFechar_Click()
Call Unload(Me)

End Sub

Private Sub Form_Load()
Me.txtReporteTecnico.Enabled = False
Me.cboElemento.Enabled = False
Me.txtAtendimentoConcluido.Enabled = False
Me.txtPropostaMelhoria.Enabled = False

Call suListarTipo
Call suListarElemento
Call suListarLocal
Call suListarAtendimento

Me.cboTipo.Text = ValorTipo
Me.cboLocal.Text = ValorLocal

Me.cboAtendente.Text = ValorAtendente
Me.txtObservacao = ValorDescricao
Me.txtReporteTecnico = ValorReporte
Me.txtPropostaMelhoria = ValorResultado
Me.txtAtendimentoConcluido = ValorFinalizado

Me.cboTipo.Enabled = False
Me.txtObservacao.Enabled = False
Me.cmdCadastrar.Enabled = False
Me.cboLocal.Enabled = False
Me.cboAtendente.Enabled = False

If ValorElemento <> "" Then
    Me.cboElemento.Text = ValorElemento
Else
End If
End Sub



Private Sub suListarTipo()
Dim rs As New ADODB.Recordset
Call ConectarBD
    strSQL = "SELECT * FROM tb_Tipo"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboTipo.AddItem rs!Tipo
        rs.MoveNext
    Loop
    Set rs = Nothing
  
End Sub

Private Sub suListarElemento()
    strSQL = "SELECT * FROM tb_Elemento"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboElemento.AddItem rs!Elemento
        rs.MoveNext
    Loop
    Set rs = Nothing
    
End Sub

Private Sub suListarLocal()
    strSQL = "SELECT * FROM tb_Local"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboLocal.AddItem rs!Local
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub suListarAtendimento()
    strSQL = "SELECT * FROM tb_Atendente"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboAtendente.AddItem rs!Atendente
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

