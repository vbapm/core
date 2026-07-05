Attribute VB_Name = "PortugueseExample"
' Módulo de exemplo com caracteres acentuados em português
'
' Este módulo contém texto em português para testar a detecção
' de codificação CP1252. As funções abaixo são puramente ilustrativas.
' ------------------------------------------------------------------
' Função para processar descrições de produtos
' ------------------------------------------------------------------
Function ObterDescricao(codigo As String) As String
    ' Esta função retorna a descrição completa do produto
    ' incluindo informações sobre fabricação e composição
    Dim descricao As String
    Dim categoria As String
    Dim preco As Double
    Dim disponivel As Boolean
    ' Verificar se o código é válido
    If Len(codigo) < 3 Then
        ObterDescricao = "Código inválido"
        Exit Function
    End If
    ' Consultar a tabela de produtos
    descricao = "Produto não encontrado"
    ' Categorias disponíveis:
    ' - Eletrônicos e acessórios
    ' - Alimentação e bebidas
    ' - Vestuário e calçados
    ' - Material de escritório
    ' - Produtos de limpeza
    Select Case Left(codigo, 2)
        Case "EL"
            categoria = "Eletrônicos"
            descricao = "Equipamento eletrônico - garantia de 12 meses"
        Case "AL"
            categoria = "Alimentação"
            descricao = "Produto alimentício - verificar data de validade"
        Case "VS"
            categoria = "Vestuário"
            descricao = "Peça de vestuário - consultar tabela de tamanhos"
        Case "ES"
            categoria = "Escritório"
            descricao = "Material de escritório - verificar stock disponível"
        Case "LP"
            categoria = "Limpeza"
            descricao = "Produto de limpeza - consultar ficha técnica"
    End Select
    ' Adicionar informações complementares
    descricao = descricao & " | Categoria: " & categoria
    ObterDescricao = descricao
End Function
' ------------------------------------------------------------------
' Função para formatar valores monetários
' ------------------------------------------------------------------
Function FormatarMoeda(valor As Double, Optional simbolo As Boolean = True) As String
    ' Formata um valor numérico como moeda no formato brasileiro
    '
    ' Parâmetros:
    '   valor   - valor a ser formatado (ex: 1234.56)
    '   simbolo - incluir símbolo R$ (padrão: True)
    '
    ' Exemplos:
    '   1234.56 -> "R$ 1.234,56"
    '   99.9    -> "R$ 99,90"
    Dim resultado As String
    Dim inteiro As Long
    Dim centavos As Integer
    Dim strInteiro As String
    Dim strCentavos As String
    ' Separar parte inteira e centavos
    inteiro = Int(valor)
    centavos = Round((valor - inteiro) * 100, 0)
    ' Formatar centavos com dois dígitos
    If centavos < 10 Then
        strCentavos = "0" & centavos
    Else
        strCentavos = CStr(centavos)
    End If
    ' Formatar parte inteira com separador de milhar
    strInteiro = Format(inteiro, "#,##0")
    ' Construir resultado
    If simbolo Then
        resultado = "R$ " & strInteiro & "," & strCentavos
    Else
        resultado = strInteiro & "," & strCentavos
    End If
    FormatarMoeda = resultado
End Function
' ------------------------------------------------------------------
' Função para validar dados de entrada
' ------------------------------------------------------------------
Function ValidarEntrada(dados As Variant) As Boolean
    ' Valida dados de entrada verificando:
    ' - Não está vazio
    ' - Não contém caracteres proibidos
    ' - Tamanho mínimo e máximo
    Dim texto As String
    Dim i As Integer
    Dim tamanho As Integer
    ' Converter para texto
    texto = CStr(dados)
    tamanho = Len(texto)
    ' Verificar se não está vazio
    If tamanho = 0 Then
        ValidarEntrada = False
        Exit Function
    End If
    ' Verificar tamanho mínimo e máximo
    If tamanho < 3 Or tamanho > 100 Then
        ValidarEntrada = False
        Exit Function
    End If
    ' Verificar caracteres não permitidos
    ' (caracteres de controle e especiais)
    For i = 1 To tamanho
        Select Case Asc(Mid(texto, i, 1))
            Case 0 To 31  ' Caracteres de controle
                ValidarEntrada = False
                Exit Function
            Case 127      ' Delete
                ValidarEntrada = False
                Exit Function
        End Select
    Next i
    ValidarEntrada = True
End Function
' ------------------------------------------------------------------
' Registro de log para auditoria
' ------------------------------------------------------------------
Sub RegistrarAuditoria(operacao As String, Optional detalhes As String = "")
    ' Esta sub-rotina registra operações no log de auditoria
    ' para fins de rastreabilidade e conformidade regulatória.
    Dim wsLog As Worksheet
    Dim ultimaLinha As Long
    Dim dataHora As String
    ' Verificar se a planilha de log existe
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    On Error GoTo 0
    If wsLog Is Nothing Then
        ' Criar planilha de log
        Set wsLog = ThisWorkbook.Worksheets.Add
        wsLog.Name = "LOG"
        wsLog.Cells(1, 1).Value = "Data/Hora"
        wsLog.Cells(1, 2).Value = "Operação"
        wsLog.Cells(1, 3).Value = "Detalhes"
        wsLog.Cells(1, 4).Value = "Usuário"
    End If
    ' Registrar operação
    dataHora = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ultimaLinha = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(ultimaLinha, 1).Value = dataHora
    wsLog.Cells(ultimaLinha, 2).Value = operacao
    wsLog.Cells(ultimaLinha, 3).Value = detalhes
    wsLog.Cells(ultimaLinha, 4).Value = Environ("USERNAME")
End Sub