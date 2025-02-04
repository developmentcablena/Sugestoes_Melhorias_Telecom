VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcompanharOS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acompanhar Sugestões"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcompanharOS_T.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15135
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwAcompanharOS 
      Height          =   8415
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14843
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   300
      Left            =   2985
      MaxLength       =   5
      TabIndex        =   4
      Top             =   420
      Width           =   1590
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      Caption         =   "&Pesquisar"
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   405
      Width           =   1095
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   2775
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   13560
      TabIndex        =   1
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Atendimento"
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
      Left            =   2985
      TabIndex        =   5
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      TabIndex        =   2
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "frmAcompanharOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As ADODB.Recordset

Private Sub cboStatus_Click()
Select Case Me.cboStatus
    Case Is = "Em Aberto"
        Me.lvwAcompanharOS.ForeColor = vbBlack
    Case Is = "Em Atendimento"
        Me.lvwAcompanharOS.ForeColor = vbBlack
    Case Is = "Finalizada"
        Me.lvwAcompanharOS.ForeColor = vbBlue
    Case Is = "Cancelada"
        Me.lvwAcompanharOS.ForeColor = vbRed
End Select

Call suListarChamados(Me.cboStatus.Text, gintUsuarioID, sFullName)

End Sub

Private Sub Form_Load()
With Me.lvwAcompanharOS
    .ColumnHeaders.Add , , "N°AT", 710
    .ColumnHeaders.Add , , "Nome", 1800, 2
    .ColumnHeaders.Add , , "Tipo", 1300, 2
    .ColumnHeaders.Add , , "Elemento", 3000, 2
    .ColumnHeaders.Add , , "Depto.", 1500, 2
    .ColumnHeaders.Add , , "Data Cadastro", 2500, 2
    .ColumnHeaders.Add , , "Local", 1800, 2
    .ColumnHeaders.Add , , "Descrição da Situação Atual", 3000, 2
    .ColumnHeaders.Add , , "Proposta de Melhoria e Resultado Esperado", 4000, 2
    .ColumnHeaders.Add , , "Atendente", 1500, 2
    .ColumnHeaders.Add , , "Reporte Técnico", 4000, 2
    .ColumnHeaders.Add , , "RT - Atendimento Concluído", 4000, 2
    .ColumnHeaders.Add , , "Data Finalizada", 2500, 2
    .ColumnHeaders.Add , , "", 0
End With

Call suListarStatus
Me.txtOSID.Visible = False
Me.cmdPesquisar.Visible = False
Me.lblOSID.Visible = False
End Sub

Private Sub suListarStatus()
Me.cboStatus.AddItem "Em Aberto"
Me.cboStatus.AddItem "Em Atendimento"
Me.cboStatus.AddItem "Finalizada"
Me.cboStatus.AddItem "Cancelada"
End Sub

Private Sub suListarChamados(ByVal vStatus As String, ByVal vUsuarioID As String, ByVal vUsuarioAD As String)
Call ConectarBD
Dim lvwItem As ListItem

If vUsuarioID = 0 Then
Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 AND Nome = '" & vUsuarioAD & "'"
        Case Is = "Em Atendimento"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 1 AND Nome = '" & vUsuarioAD & "'"
        Case Is = "Finalizada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 2 AND Nome = '" & vUsuarioAD & "'"
        Case Is = "Cancelada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 3 AND Nome = '" & vUsuarioAD & "'"
    End Select

Else
    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 AND UsuarioID = " & vUsuarioID & ""
        Case Is = "Em Atendimento"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 1 AND UsuarioID = " & vUsuarioID & ""
        Case Is = "Finalizada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 2 AND UsuarioID = " & vUsuarioID & ""
        Case Is = "Cancelada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 3 AND UsuarioID = " & vUsuarioID & ""
    End Select
    
End If

Set rs = New ADODB.Recordset
rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic

Me.lvwAcompanharOS.ListItems.Clear

If rs.EOF = False Then
    Do While Not rs.EOF
        Set lvwItem = Me.lvwAcompanharOS.ListItems.Add(, , Format(rs!ATID, "0000"))
            If Len(rs!Nome) > 0 Then
                lvwItem.SubItems(1) = rs!Nome
            Else
            '    lvwItem.SubItems(1) = sFullName
            End If

            lvwItem.SubItems(2) = rs!Tipo
            If IsNull(rs!Elemento) Then
                lvwItem.SubItems(3) = ""
            Else
                lvwItem.SubItems(3) = rs!Elemento
            End If
            If Len(rs!Departamento) > 0 Then
                lvwItem.SubItems(4) = rs!Departamento
            Else
                lvwItem.SubItems(4) = strDescription
            End If
            lvwItem.SubItems(5) = rs!DataCadastro
            lvwItem.SubItems(6) = rs!Local
            lvwItem.SubItems(7) = rs!DescricaoMelhorias
                
            If IsNull(rs!PropostaResultado) Then
                rs!PropostaResultado = ""
            Else
                lvwItem.SubItems(8) = rs!PropostaResultado
            End If
                
            lvwItem.SubItems(9) = rs!Atendente
                
            If IsNull(rs!ReporteTecnico) Then
                lvwItem.SubItems(10) = ""
            Else
                lvwItem.SubItems(10) = rs!ReporteTecnico
            End If
            If IsNull(rs!FinalizacaoOS) Then
                lvwItem.SubItems(11) = ""
            Else
                lvwItem.SubItems(11) = rs!FinalizacaoOS
            End If
                
            If IsNull(rs!DataBaixaAtual) Then
                lvwItem.SubItems(12) = ""
            Else
                lvwItem.SubItems(12) = rs!DataBaixaAtual
            End If
                
            If IsNull(rs!MotivoCancelamento) Then
                lvwItem.SubItems(13) = ""
            Else
                lvwItem.SubItems(13) = rs!MotivoCancelamento
            End If
        rs.MoveNext
    Loop
Else
    MsgBox "Nenhum registro foi encontrato!", vbOKOnly + vbInformation, "Consulta"
    Exit Sub
    
End If



rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
End Sub


Private Sub lvwAcompanharOS_DblClick()
Dim frm As New frmVisualizarAT

If Me.cboStatus = "Cancelada" Then
    If Not Me.lvwAcompanharOS.selectedItem Is Nothing Then
        Dim selectedItem As ListItem
        Set selectedItem = Me.lvwAcompanharOS.selectedItem
        
        ValorTipo = selectedItem.SubItems(2)
        ValorElemento = selectedItem.SubItems(3)
        ValorLocal = selectedItem.SubItems(6)
        ValorAtendente = selectedItem.SubItems(9)
        ValorResultado = selectedItem.SubItems(10)
        ValorFinalizado = selectedItem.SubItems(13)
        frm.Label1.Caption = "Reporte Técnico - Atendimento Cancelado:"
        frm.txt02.ForeColor = vbRed
        frm.cmdCadastrar.Enabled = False
        
        frm.Show (vbModal)
    End If
Else
    If Not Me.lvwAcompanharOS.selectedItem Is Nothing Then
        Dim selectedItem2 As ListItem
        Set selectedItem2 = Me.lvwAcompanharOS.selectedItem
        
        ValorTipo = selectedItem2.SubItems(2)
        ValorElemento = selectedItem2.SubItems(3)
        ValorLocal = selectedItem2.SubItems(6)
        ValorAtendente = selectedItem2.SubItems(9)
       
        ValorResultado = selectedItem2.SubItems(10)
        ValorFinalizado = selectedItem2.SubItems(11)
        frm.txt02.ForeColor = vbBlue
         frm.cmdCadastrar.Enabled = False
        
        frm.Show (vbModal)
    End If
End If
End Sub
