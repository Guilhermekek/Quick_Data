# Melhorias preventivas — código VBA

> Baseado nos achados do arquivo `QD 3.23 - P2_vf.xlsx` (69.537 nomes definidos, 415.297 fórmulas vivas, `DP_Segmento` com 63.645 linhas, arquivo salvo sem VBA). Código pronto pra colar no Quick Data original. Testar em cópia antes de aplicar no arquivo de produção.

---

## 1. `ThisWorkbook.cls` — detectar e travar os dois problemas mais graves

Substitui o `Workbook_Open` vazio atual e adiciona `Workbook_BeforeSave`.

```vb
Private Const LIMITE_NOMES_DEFINIDOS As Long = 300   ' saudável hoje: 166. Alerta bem antes de virar um caso "P2".
Private Const CHAVE_PROP_ULTIMA_CONTAGEM As String = "QD_UltimaContagemNomes"

Private Sub Workbook_Open()
    Call Verificar_Saude_Nomes_Definidos(AoAbrir:=True)
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    ' --- Bloqueia salvar em formato sem macro (o que aconteceu no P2) ---
    If SaveAsUI Then
        Dim caminhoEscolhido As Variant
        caminhoEscolhido = Application.GetSaveAsFilename( _
            InitialFileName:=ThisWorkbook.Name, _
            FileFilter:="Excel Binário (*.xlsb), *.xlsb")

        If caminhoEscolhido = False Then
            Cancel = True
            Exit Sub
        End If

        If Right$(caminhoEscolhido, 5) <> ".xlsb" Then
            MsgBox "Esse formato não guarda macro nenhuma do Quick Data (extração, ajustes, tudo)." & vbCrLf & _
                   "Salvando assim, o arquivo vira uma foto morta, sem nenhum botão funcionando." & vbCrLf & vbCrLf & _
                   "Salve como .xlsb.", vbCritical, "Formato bloqueado"
            Cancel = True
            Exit Sub
        End If

        ThisWorkbook.SaveAs Filename:=caminhoEscolhido, FileFormat:=xlExcel12  ' .xlsb
        Cancel = True   ' já salvamos manualmente acima; cancela o save automático do Excel
        Exit Sub
    End If

    ' --- Alerta se o número de nomes definidos saltou desde o último save ---
    Call Verificar_Saude_Nomes_Definidos(AoAbrir:=False)

End Sub

Private Sub Verificar_Saude_Nomes_Definidos(AoAbrir As Boolean)
Dim qtdeAtual As Long, qtdeAnterior As Long

    qtdeAtual = ThisWorkbook.Names.Count

    On Error Resume Next
    qtdeAnterior = CLng(ThisWorkbook.CustomDocumentProperties(CHAVE_PROP_ULTIMA_CONTAGEM).Value)
    On Error GoTo 0

    If qtdeAtual > LIMITE_NOMES_DEFINIDOS And (qtdeAtual - qtdeAnterior) > 50 Then
        MsgBox "O arquivo tem " & qtdeAtual & " nomes definidos (tinha " & qtdeAnterior & " da última vez)." & vbCrLf & _
               "Esse salto grande normalmente vem de colar célula de outra planilha (ex.: arquivo com Capital IQ/Bloomberg)." & vbCrLf & vbCrLf & _
               "Rode 'Limpar Defined Names com erro' antes de continuar.", vbExclamation, "Nomes definidos anormais"
    End If

    ' grava a contagem atual pra comparar na próxima vez
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(CHAVE_PROP_ULTIMA_CONTAGEM).Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=CHAVE_PROP_ULTIMA_CONTAGEM, LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=qtdeAtual

End Sub
```

---

## 2. Novo módulo `Manutencao_Preventiva.bas`

### 2a. Congelar fórmula em valor — versão genérica reutilizável

Hoje cada módulo de extração repete o padrão "copiar → colar valor" na mão. Isso centraliza, pra nenhuma coluna nova (tipo "Adicoes"/"Inorgânico" no P2) escapar do padrão.

```vb
Public Sub Congelar_Formulas(ByVal sh As Worksheet, ByVal rng As Range)
    ' Converte fórmula em valor num range específico — chamar sempre que uma coluna
    ' nova for adicionada fora do fluxo padrão dos módulos Extracao_*.
    With sh
        rng.Copy
        rng.PasteSpecial Paste:=xlPasteValues
    End With
    Application.CutCopyMode = False
End Sub
```

### 2b. Checagem de crescimento anormal — usar dentro de cada `Refresh_*`

```vb
Public Function Crescimento_E_Normal(ByVal NomeTabela As String, _
                                      ByVal Ult_Lin_Nova As Long, _
                                      ByVal Ult_Lin_Anterior As Long) As Boolean
Const FATOR_MAXIMO As Double = 3   ' acima de 3x o tamanho anterior, para e pergunta

    If Ult_Lin_Anterior <= 1 Then
        Crescimento_E_Normal = True   ' primeira carga, nada pra comparar
        Exit Function
    End If

    If Ult_Lin_Nova > Ult_Lin_Anterior * FATOR_MAXIMO Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox(NomeTabela & " tinha " & Ult_Lin_Anterior & " linhas e a atualização gerou " & _
                           Ult_Lin_Nova & "." & vbCrLf & _
                           "Isso pode ser um refresh que está somando em vez de substituir." & vbCrLf & vbCrLf & _
                           "Continuar mesmo assim?", vbExclamation + vbYesNo, "Crescimento fora do esperado")
        Crescimento_E_Normal = (resposta = vbYes)
    Else
        Crescimento_E_Normal = True
    End If

End Function
```

**Como plugar no `Refresh_Base_Segmento` existente** (`Refresh_DP_Segmento.bas`) — antes de gravar o resultado novo, guardar quantas linhas a tabela tinha e comparar:

```vb
Sub Refresh_Base_Segmento()
Dim Ult_Lin_Anterior As Long, Ult_Lin_Nova As Long
Dim fn As WorksheetFunction: Set fn = Application.WorksheetFunction

    Ult_Lin_Anterior = fn.CountA(Sheet21.Columns("A:A"))   ' conta ANTES de limpar/reescrever

    ' ... código existente que consulta o SQL e monta o resultado em memória/recordset ...

    Ult_Lin_Nova = rs.RecordCount   ' ou o total de linhas que o novo resultado vai ter

    If Not Crescimento_E_Normal("DP_Segmento", Ult_Lin_Nova, Ult_Lin_Anterior) Then
        Exit Sub   ' usuário cancelou — não grava
    End If

    ' ... resto do código existente que limpa Sheet21 e escreve o resultado novo ...

End Sub
```

### 2c. Botão de diagnóstico — saúde do arquivo, não só erro de fórmula

Grava na mesma aba `tk_Lista_de_erros`, reaproveitando a estrutura que já existe (`fn_ListAllErrors`).

```vb
Public Sub Diagnostico_Arquivo()
Dim sh As Worksheet: Set sh = Sheets("tk_Lista_de_erros")
Dim prox As Long: prox = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1
Dim fn As WorksheetFunction: Set fn = Application.WorksheetFunction

Dim qtdeNomes As Long, qtdeBase As Long, qtdeSupLinhas As Long, qtdeSegmento As Long
Dim tamanhoMB As Double

    qtdeNomes = ThisWorkbook.Names.Count
    qtdeBase = fn.CountA(Sheet3.Columns("A:A"))
    qtdeSupLinhas = fn.CountA(Sheet15.Columns("A:A"))
    qtdeSegmento = fn.CountA(Sheet21.Columns("A:A"))
    tamanhoMB = Round(FileLen(ThisWorkbook.FullName) / 1024 / 1024, 1)

    With sh
        .Cells(prox, 1) = "Diagnostico_Arquivo"
        .Cells(prox, 2) = "N/A"
        .Cells(prox, 3) = "N/A"
        .Cells(prox, 4) = "N/A"
        .Cells(prox, 5) = "N/A"
        .Cells(prox, 6) = "Nomes=" & qtdeNomes & " | Base=" & qtdeBase & " linhas | " & _
                          "Sup_Linhas=" & qtdeSupLinhas & " | DP_Segmento=" & qtdeSegmento & _
                          " | Arquivo=" & tamanhoMB & "MB"
        .Cells(prox, 7) = Environ("Username")
        .Cells(prox, 8) = Now
    End With

    MsgBox "Diagnóstico registrado:" & vbCrLf & _
           "Nomes definidos: " & qtdeNomes & vbCrLf & _
           "Linhas na Base: " & qtdeBase & vbCrLf & _
           "Linhas em Sup_Linhas: " & qtdeSupLinhas & vbCrLf & _
           "Linhas em DP_Segmento: " & qtdeSegmento & vbCrLf & _
           "Tamanho do arquivo: " & tamanhoMB & " MB", vbInformation, "Diagnóstico do Quick Data"

End Sub
```

Ligar num botão na aba Extração igual aos outros (`Shape → Atribuir Macro → Diagnostico_Arquivo`), pra rodar antes de cada fechamento e ter histórico de tendência — não só descobrir o problema quando o arquivo já está de 37MB.

---

## 3. Guarda simples contra colagem externa (opcional, mais invasivo)

Coloca no code-behind de `Sheet3` (Base) e `Sheet13` (Ajustes). Como os módulos de extração já rodam com `EnableEvents = False` (via `Desligar_Tudo`), isso só dispara em ação manual do usuário — não interfere na extração normal.

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
Const LIMITE_CELULAS_COLAGEM As Long = 20
Const LIMITE_NOMES_APOS_COLAR As Long = 300

    If Target.Cells.Count < LIMITE_CELULAS_COLAGEM Then Exit Sub   ' edição normal, não é colagem em massa

    If ThisWorkbook.Names.Count > LIMITE_NOMES_APOS_COLAR Then
        Application.EnableEvents = False
        If MsgBox("Essa colagem pode ter trazido nomes definidos de outra planilha (" & _
                  ThisWorkbook.Names.Count & " nomes no arquivo agora)." & vbCrLf & _
                  "Desfazer a colagem?", vbExclamation + vbYesNo, "Colagem suspeita") = vbYes Then
            Application.Undo
        End If
        Application.EnableEvents = True
    End If

End Sub
```

---

**Ordem sugerida pra aplicar:** 1 e 2c primeiro (baratos, só alertam, não travam nada) → testar uma semana → depois 2b nos `Refresh_*` → o item 3 (guarda de colagem) só se os alertas passivos não forem suficientes, porque é o mais invasivo dos cinco.
