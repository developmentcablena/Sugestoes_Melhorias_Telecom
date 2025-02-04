VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRelatorio 
   Caption         =   "Relatório"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_T.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8445
   ScaleWidth      =   12870
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Arquivo 
      Left            =   8280
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboStatusR 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar pgbPorgressoR 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportarR 
      Caption         =   "Exportar XLS"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdGerarR 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   7920
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvwRelatorio 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private strSQL As String

Private Sub cmdExportarR_Click()
On Error GoTo Erro
    With Arquivo
        .CancelError = True
        .filter = "Pasta de Trabalho do Excel"           ' "Pasta de Trabalho da Microsoft Office Excel (*.xls)|*.xls|All Files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNOverwritePrompt
        On Error Resume Next
        .ShowSave
        If Err.Number <> 0 Then
            MsgBox "Cancelada", vbInformation, "Suporte Simac"
            Err.Clear
            Exit Sub
        End If
        
    End With
    
    Dim filepath As String
    filepath = Arquivo.filename
    
    If filepath <> "" Then
        Me.MousePointer = 11
        Call salvarExcel(filepath)
        Me.MousePointer = 0
    End If
    Exit Sub
Erro:
    MsgBox "Erro: " & Err.Description, vbCritical, "Suporte Simac"
    On Error Resume Next
End Sub

Private Sub salvarExcel(filepath As String)
Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim j As Integer
    Dim totalItems As Integer
    Dim currentItem As Integer
    ' Cria uma nova instância do Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    ' Copia os cabeçalhos das colunas do ListView para a planilha
    For i = 1 To Me.lvwRelatorio.ColumnHeaders.Count
        xlSheet.Cells(1, i).Value = Me.lvwRelatorio.ColumnHeaders(i).Text
    Next i
    ' Formata os cabeçalhos em negrito e centralizado
    With xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, Me.lvwRelatorio.ColumnHeaders.Count))
        .Font.Bold = True
        .HorizontalAlignment = -4108 ' xlCenter
        .Interior.Color = RGB(222, 212, 27)
        '.WrapText = True
        
    End With
    ' Obtém o total de itens para a barra de progresso
    totalItems = Me.lvwRelatorio.ListItems.Count

    ' Copia os itens do ListView para a planilha
    For i = 1 To totalItems
        xlSheet.Cells(i + 1, 1).Value = Me.lvwRelatorio.ListItems(i).Text ' Primeiro subitem
        For j = 1 To Me.lvwRelatorio.ListItems(i).ListSubItems.Count
            xlSheet.Cells(i + 1, j + 1).Value = Me.lvwRelatorio.ListItems(i).SubItems(j)
        Next j
    Next i
    
    ' Ajusta a largura das colunas automaticamente
    xlSheet.Columns.AutoFit
    xlSheet.Columns.WrapText = True
    
    
    
    ' Salva o arquivo Excel
    xlBook.SaveAs filepath

    
    xlBook.Close False
    xlApp.Quit
    
    ' Libera a memória
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    ' Exibe a mensagem de sucesso
    MsgBox "Os dados foram salvos com sucesso em " & filepath, vbInformation, "Suporte Simac"
End Sub







Private Sub cmdGerarR_Click()
If Len(Trim(Me.cboStatusR.Text)) = 0 Then
    MsgBox "Selecione um Status!", vbOKOnly + vbExclamation, "Suporte Simac"
    Me.cboStatusR.SetFocus
    Exit Sub
End If

Call suGerarRelatorio(Me.cboStatusR)
    
    
End Sub

Private Sub suGerarRelatorio(ByVal vStatus As String)
Call ConectarBD
Dim lvwItem As ListItem
Select Case vStatus
    Case Is = "Em Aberto"
        strSQL = "SELECT * FROM tb_Atendimento WHERE Status = 0"
    Case Is = "Em Atendimento"
        strSQL = "SELECT 8 FROM  tb_Atendimento WHERE Status = 1"
    Case Is = "Finalizada"
        strSQL = "SELECT * FROM tb_Atendimento WHERE Status = 2"
    Case Is = "Cancelada"
        strSQL = "SELECT * FROM tb_Atendimento WHERE Status = 3"
    Case Is = "Relatório Geral"
        strSQL = "SELECT * FROM tb_Atendimento"
End Select

Set rs = New ADODB.Recordset
rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic

Me.lvwRelatorio.ListItems.Clear

If rs.EOF = False Then
    Me.pgbPorgressoR.Min = 0
    Me.pgbPorgressoR.Value = 0
    Me.pgbPorgressoR.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
    Screen.MousePointer = 11
Else
    MsgBox "Nenhum registro foi encontrado!", vbOKOnly + vbInformation, "Suporte Simac"
    Exit Sub
End If

    Do While Not rs.EOF
        Set lvwItem = Me.lvwRelatorio.ListItems.Add(, , Format(rs!ATID, "0000"))
            If IsNull(rs!Nome) Then
                lvwItem.SubItems(1) = ""
            Else
                lvwItem.SubItems(1) = rs!Nome
            End If
            
            If IsNull(rs!Tipo) Then
                lvwItem.SubItems(2) = ""
            Else
                lvwItem.SubItems(2) = rs!Tipo
            End If
            
            If IsNull(rs!Elemento) Then
                lvwItem.SubItems(3) = ""
            Else
                lvwItem.SubItems(3) = rs!Elemento
            End If
            
            lvwItem.SubItems(4) = rs!Departamento
            lvwItem.SubItems(5) = rs!DataCadastro
            lvwItem.SubItems(6) = rs!Local
            lvwItem.SubItems(7) = rs!DescricaoMelhorias
            
            If IsNull(rs!PropostaResultado) Then
                lvwItem.SubItems(8) = ""
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
        Me.pgbPorgressoR.Value = Me.pgbPorgressoR.Value + 1
    Loop
rs.Close
Set rs = Nothing
Screen.MousePointer = 0
MsgBox "Relatório gerado com sucesso!", vbInformation, "Suporte Simac"
Exit Sub
    
End Sub

Private Sub Command3_Click()
Call Unload(Me)
End Sub

Private Sub Form_Load()

With Me.lvwRelatorio
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
    .LabelEdit = lvwManual
End With

Call suListarStatus
End Sub
Private Sub suListarStatus()
Me.cboStatusR.AddItem "Em Aberto"
Me.cboStatusR.AddItem "Em Atendimento"
Me.cboStatusR.AddItem "Finalizada"
Me.cboStatusR.AddItem "Cancelada"
Me.cboStatusR.AddItem "Relatório Geral"
End Sub
