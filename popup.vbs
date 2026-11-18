Sub xc_PowerDoc_OnF_B_ConfigChanged()

' POWERDOC | MANUTENCAO DAS INSTRUCOES RECONHECIDAS
'
' Este script interpreta as frases recebidas do PowerDoc e cria os checks.
' O texto das instrucoes e mantido no PowerDoc; o menu e os IDs dos documentos
' ficam em button.vbs. Aqui sao mantidos o reconhecimento, os valores e as tags.
'
' ESQUELETO DO CODIGO (NA ORDEM DO ARQUIVO)
' 01 | xc_PowerDoc_OnF_B_ConfigChanged
'      Valida o disparo e DocID, consulta o PowerDoc, interpreta cada linha
'      e monta textos, checks, mensagem de autonomia e dimensoes da tela.
'      MontarClausulaAutonomia: comparacao e qualidade para o resumo final.
' 02 | ExtrairInstrucao
'      Aplica as regras por prioridade e retorna os dados da instrucao.
' 03 | Extrair* / Normalizar* / CodigoOperadorPowerDoc / NovoRegExp
'      Extrai referencias e locais, normaliza nomes e prepara operadores
'      e expressoes regulares usados pelas regras.
' 04 | ObterTag*: resolve as tags SCADA conforme equipamento e subestacao.
'      04.01 disjuntores; 04.02 potencia/religamento IVP;
'      04.03 PSD/diferenca angular PGR; 04.04 potencia IRN;
'      04.05 tensao de barras; 04.06 estado N-1.
' 05 | aspJSON
'      Interpreta as respostas JSON para acesso aos dados do documento.
' 06 | Vazia
'      Rotina sem implementacao.
'
' EDITAR UMA INSTRUCAO JA PREVISTA
' 1. Localize a regra no bloco 02, pela ordem abaixo:
'    02.01 abrir/fechar disjuntor; 02.02 sistema completo/N-1;
'    02.03 potencia IRN/PGR C1; 02.04 potencia LT PGR C1/C2;
'    02.05 diferenca angular; 02.06 religamento tripolar;
'    02.07 PSD dos transformadores; 02.08 comparacao numerica generica.
' 2. Para aceitar outra redacao da mesma instrucao, ajuste o padrao de
'    NovoRegExp da regra. Preserve a ordem dos grupos capturados: cada
'    SubMatches(n) e usado para identificar equipamento, SE ou referencia.
'    Na regra generica, as expressoes de comparacao ficam em ExtrairOperador
'    e ExtrairReferenciaNumerica (bloco 03); mantenha as duas coerentes.
' 3. Para mudar limite ou estado esperado, confira referencia_valor,
'    referencia_unidade, referencia_estado e operador_check na regra.
'    Operadores do check: 1 >, 2 >=, 3 =, 4 <=, 5 <.
'    Algumas referencias sao fixas na regra; outras vem da frase capturada.
' 4. Para trocar somente a leitura SCADA, edite a funcao do bloco 04:
'    04.01 ObterTagGPEDisjuntor (IVP/PGR/PGS/SMS e numeros compartilhados);
'    04.02 ObterTagGPEPotenciaAtiva / ObterTagGPEReligamentoTripolar (IVP);
'    04.03 ObterTagGPEPSD / ObterTagGPEDiferencaAngular (PGR);
'    04.04 ObterTagGPEPotenciaAtivaIRN (IRN);
'    04.05 ObterTagGPE (tensao de barras IVP/IRN/PGR);
'    04.06 ObterTagCLCN1 (estado N-1 IVP/IRN/PGR).
'
' ADICIONAR EQUIPAMENTO A UMA INSTRUCAO PREVISTA
' Reutilize a regra existente e amplie seus equipamentos/circuitos aceitos,
' quando necessario; depois inclua o Case na funcao de tags correspondente.
' Exemplo: um novo disjuntor exige o numero em ObterTagGPEDisjuntor e sua tag
' de posicao. Se o numero existir em mais de uma SE, explicite a selecao por
' subestacao e preserve os retornos Else existentes para os demais casos.
' Para novos nomes de SE, confira tambem NormalizarSiglaSubestacao;
' para barras, confira NormalizarEquipamento e os testes locais do bloco 04.
' As leituras sao agrupadas pela SE da tag, que pode diferir do nome da LT.
'
' Se a instrucao exigir uma regra adicional, use uma regra semelhante do
' bloco 02 como modelo: padrao de texto, campos do dicionario, tag e retorno
' (Set ExtrairInstrucao = info / Exit Function). Posicione-a antes da regra
' generica, respeitando as prioridades existentes. Inclua seu mapeamento no
' bloco 04 e confira referencia_valor, tag_gpe e operador_check do novo caso.
'
' REMOVER OU DEIXAR DE VALIDAR UMA INSTRUCAO
' Para retirar a linha do documento, remova-a no PowerDoc.
' Para manter o texto visivel, mas sem check de um equipamento, remova seu
' mapeamento de tag e garanta que esse caso retorne "", sem cair em um Else
' que devolva outra tag. No N-1, tag vazia permite continuar nas outras regras.
' Para desativar uma regra inteira, revise tambem as regras seguintes:
' retirar seu reconhecimento pode fazer a frase cair na comparacao generica.
'
' CUIDADOS AO MANTER AS REGRAS
' A ordem do bloco 02 define prioridade; preserve os Exit Function e deixe
' a comparacao generica por ultimo. O dicionario de retorno e o mesmo para
' todas as regras (02.00). O bloco 01.06.b cria o check apenas se houver
' referencia_valor, tag_gpe e operador_check > 0; o texto e sempre exibido.
' Confira frases atendidas e nao atendidas, abrir/fechar, circuitos C1/C2,
' numeros compartilhados entre SEs, operador, referencia e caminho da tag.
' Valide no Elipse o texto e o check resultantes apos cada alteracao.
'

' ============================================================================
' 01 | EVENTO PRINCIPAL
' Etapas comentadas dentro do evento original.
' ============================================================================


    ' ------------------------------------------------------------------------
    ' 01.01 | Disparo e validacoes iniciais
    ' ------------------------------------------------------------------------
    If Not F_B_Config Then Exit Sub
    F_B_Config = False

    If DocID = "" Then Exit Sub

    ' ------------------------------------------------------------------------
    ' 01.02 | Referencias de tela e inicializacao
    ' ------------------------------------------------------------------------
    Set scr = Parent
    Set margem = scr.Item("dwr_9")
    Set cabecalho = Item("dwr_8")

    Dim oJsonParser, instrucao, resultadoExtracao, contextoOperacional
    Dim textoAutonomia, temCheck, valorReferencia
    Dim comparacoesAutonomia, qualidadesAutonomia, autonomiaIncompleta
    Dim comparacao, qualidade
    Const mensagemRecompor = "Recompor a FT com Autonomia"
    Const mensagemDisponibilizar = "Disponibilizar a FT com Autonomia"

    Set oJsonParser = New aspJSON
    comparacoesAutonomia = ""
    qualidadesAutonomia = ""
    autonomiaIncompleta = False

    ' ------------------------------------------------------------------------
    ' 01.03 | HTTP: consultar o ID do documento
    ' ------------------------------------------------------------------------
    powerDocServerAddressToGetID = _
        "http://middleware-powerdocs.ds55.local/integrations/v1/sharing?tag=" & DocID

    Set aj_XmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    aj_XmlHttp.Open "GET", powerDocServerAddressToGetID, False
    aj_XmlHttp.setTimeouts 5000, 5000, 5000, 5000
    aj_XmlHttp.setRequestHeader "Content-Type", "application/json"
    aj_XmlHttp.setRequestHeader "Authorization", _
        "Basic ZWxpcHNlOkJ2Y0VtcnJ4UVphWG9iaVg3ZEI0alhnS0VEdXFuOHREYXlBNlpxZElpL3ZOajdZdQ=="
    aj_XmlHttp.Send

    If aj_XmlHttp.Status <> 200 Then
        MsgBox "Erro " & aj_XmlHttp.Status & " na requisicao do ID."
        Exit Sub
    End If

    oJsonParser.LoadJSONString aj_XmlHttp.responseText
    idUpdated = oJsonParser.data("content").Item(0).Item("id")

    ' ------------------------------------------------------------------------
    ' 01.04 | HTTP: consultar o conteudo do documento
    ' ------------------------------------------------------------------------
    powerDocServerAddressToGetInstruction = _
        "http://middleware-powerdocs.ds55.local/integrations/v1/sharing/" & _
        idUpdated & "/content"

    Set aj_XmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    aj_XmlHttp.Open "GET", powerDocServerAddressToGetInstruction, False
    aj_XmlHttp.setTimeouts 5000, 5000, 5000, 5000
    aj_XmlHttp.setRequestHeader "Content-Type", "application/json"
    aj_XmlHttp.setRequestHeader "Authorization", _
        "Basic ZWxpcHNlOkJ2Y0VtcnJ4UVphWG9iaVg3ZEI0alhnS0VEdXFuOHREYXlBNlpxZElpL3ZOajdZdQ=="
    aj_XmlHttp.Send

    If aj_XmlHttp.Status <> 200 Then
        MsgBox "Erro " & aj_XmlHttp.Status & " na requisicao das instrucoes."
        Exit Sub
    End If

    oJsonParser.LoadJSONString aj_XmlHttp.responseText

    ' ------------------------------------------------------------------------
    ' 01.05 | Preparar linhas e contexto operacional
    ' ------------------------------------------------------------------------
    rows = oJsonParser.data("rows").Count
    contextoOperacional = CStr(Item("titulo").P_S_TXT)

    ' ------------------------------------------------------------------------
    ' 01.06 | Percorrer instrucoes: interpretar, exibir e vincular check
    ' ------------------------------------------------------------------------
    For i = 0 To rows - 1
        instrucao = CStr(oJsonParser.data("rows").Item(i).Item(1))
        Set resultadoExtracao = ExtrairInstrucao(instrucao, contextoOperacional)

        ' 01.06.a | Texto da instrucao
        Set texto = scr.AddObject( _
            "Panel.DrawString", _
            True, _
            "texto00" _
        )

        texto.Value = instrucao
        texto.X = "900"
        texto.Y = (i + 1) * 1200
        ' Reservar 900 a esquerda para o check e 300 a direita.
        texto.Width = scr.Width - 1200
        texto.Height = "1200"
        texto.Multiline = True
        texto.WordWrap = True
        texto.StretchText = False
        texto.TextAlignment = 0
        texto.VertTextAlignment = 1
        texto.TextFont.Name = "Arial"
        texto.TextFont.Size = 10
        texto.TextFont.Bold = False
        texto.TextColor = RGB(90, 90, 90)
        texto.BringToFront()

        ' 01.06.b | Check somente quando referencia, tag e operador existem
        temCheck = Trim(CStr(resultadoExtracao("referencia_valor"))) <> "" And _
                   Trim(CStr(resultadoExtracao("tag_gpe"))) <> "" And _
                   resultadoExtracao("operador_check") > 0

        If temCheck Then
            ' PowerDoc aceita ponto ou virgula; CDbl usa o locale do Viewer.
            valorReferencia = CDbl(Replace(Replace( _
                resultadoExtracao("referencia_valor"), ",", "."), _
                ".", Mid(CStr(1.5), 2, 1)))

            Set check = scr.AddObject( _
                "cot-xclib-powerdoc.xc_PowerDoc_Check", _
                True, _
                "check00" _
            )

            check.param_LogicalOperator = _
                CInt(resultadoExtracao("operador_check"))

            check.X = "300"
            ' Centralizar o icone pela altura real na mesma faixa do texto.
            check.Y = texto.Y + (texto.Height - check.Height) / 2
            check.param_PowerDoc = valorReferencia

            Set Bind = check.Links.CreateLink( _
                "scada_Measurement", _
                CStr(resultadoExtracao("tag_gpe")), _
                0 _
            )

            check.BringToFront()
        End If

        ' 01.06.c | Ignorar manobras de disjuntor, mesmo sem tag mapeada.
        ' Uma linha nao vazia sem check impede confirmar o conjunto.
        If resultadoExtracao("tipo") <> "ACAO" And Trim(instrucao) <> "" Then
            If temCheck Then
                comparacao = MontarClausulaAutonomia( _
                    resultadoExtracao, valorReferencia, qualidade)
                If comparacao = "" Then
                    autonomiaIncompleta = True
                Else
                    If comparacoesAutonomia <> "" Then
                        comparacoesAutonomia = comparacoesAutonomia & " And "
                        qualidadesAutonomia = qualidadesAutonomia & " And "
                    End If
                    comparacoesAutonomia = comparacoesAutonomia & comparacao
                    qualidadesAutonomia = qualidadesAutonomia & qualidade
                End If
            Else
                autonomiaIncompleta = True
            End If
        End If
    Next

    ' ------------------------------------------------------------------------
    ' 01.07 | Resumo informativo das condicoes, sem comando de manobra
    ' ------------------------------------------------------------------------
    Set textoAutonomia = scr.AddObject( _
        "Panel.DrawString", _
        True, _
        "autonomia00" _
    )

    textoAutonomia.Value = mensagemDisponibilizar
    textoAutonomia.X = "900"
    textoAutonomia.Y = (rows + 1) * 1200
    ' Margens iguais para centralizar o texto no popup.
    textoAutonomia.Width = scr.Width - 1800
    textoAutonomia.Height = "1200"
    textoAutonomia.Multiline = True
    textoAutonomia.WordWrap = True
    textoAutonomia.StretchText = False
    textoAutonomia.TextAlignment = 1
    textoAutonomia.VertTextAlignment = 1
    textoAutonomia.TextFont.Name = "Arial"
    textoAutonomia.TextFont.Size = 10
    textoAutonomia.TextFont.Bold = True
    textoAutonomia.TextFont.Underline = True
    textoAutonomia.TextColor = RGB(90, 90, 90)
    textoAutonomia.BringToFront()

    ' Recompor somente quando todas as condicoes puderem ser confirmadas.
    If Not autonomiaIncompleta And comparacoesAutonomia <> "" Then
        textoAutonomia.Links.CreateLink "Value", _
            "IIf(" & qualidadesAutonomia & " And " & comparacoesAutonomia & _
            ", """ & mensagemRecompor & """, """ & mensagemDisponibilizar & """)", 0
    End If

    ' ------------------------------------------------------------------------
    ' 01.08 | Dimensoes da tela, margem e cabecalho
    ' ------------------------------------------------------------------------
    ' Mesma altura calculada em TamanhoTela nos botoes: autonomia e rodape.
    scr.Height = ((rows + 2) * 1200) + 800

    margem.Width = scr.Width
    margem.Height = scr.Height

    cabecalho.Width = scr.Width

    Item("titulo").X = 0
    Item("titulo").Width = scr.Width
    Item("titulo").P_L_Align = 1

    Item("data").X = 0
    Item("data").Width = scr.Width
    Item("data").P_L_Align = 2

    ' ------------------------------------------------------------------------
    ' 01.09 | Data de leitura e vinculo do titulo
    ' ------------------------------------------------------------------------
    auxTxt = "Leitura: " & _
        Right("0" & Day(Now), 2) & "/" & _
        Right("0" & Month(Now), 2) & "/" & _
        Year(Now) & " - " & _
        Right("0" & Hour(Now), 2) & "h" & _
        Right("0" & Minute(Now), 2) & "min" & _
        Right("0" & Second(Now), 2) & "s"

    Item("titulo").Links.RemoveLink "P_S_TXT"
    Item("titulo").Links.CreateLink "P_S_TXT", "Parent.P_S_Titulo"
    Item("data").P_S_TXT = auxTxt
End Sub


Function MontarClausulaAutonomia(ByVal info, ByVal valorReferencia, ByRef qualidade)
    Dim tagGPE, medida, operador, valor

    MontarClausulaAutonomia = ""
    qualidade = ""
    operador = info("operador")
    If CodigoOperadorPowerDoc(operador) <> info("operador_check") Then Exit Function
    If info("operador_check") = 0 Then Exit Function

    tagGPE = Trim(CStr(info("tag_gpe")))
    If tagGPE = "" Then Exit Function

    ' Potencia ativa usa Abs, mas a qualidade pertence a medida original.
    If Left(tagGPE, 4) = "Abs(" And Right(tagGPE, 1) = ")" Then
        medida = Mid(tagGPE, 5, Len(tagGPE) - 5)
        valor = "Abs(" & medida & ".Value)"
    Else
        medida = tagGPE
        valor = medida & ".Value"
    End If

    qualidade = "OPCIsGood(" & medida & ".Quality)"
    MontarClausulaAutonomia = "(" & valor & " " & operador & " " & _
        Replace(CStr(valorReferencia), ",", ".") & ")"
End Function


' ============================================================================
' 02 | REGRAS DE INTERPRETACAO
' Manter a prioridade e os pontos de retorno.
' ============================================================================

Function ExtrairInstrucao(ByVal frase, ByVal contextoOperacional)

    ' A prioridade das regras e parte do comportamento. Nao reagrupar por SE.
    ' A regra N-1 so retorna se encontrar tag; caso contrario, o fluxo continua.
    ' contextoOperacional e recebido, mas nao e utilizado nesta versao.

    Dim info, texto, re, correspondencias
    Dim operador, circuito, subestacao

    ' ------------------------------------------------------------------------
    ' 02.00 | Contrato de retorno: dicionario de dados da instrucao
    ' ------------------------------------------------------------------------
    Set info = CreateObject("Scripting.Dictionary")

    info.Add "tipo", "NAO_IDENTIFICADO"
    info.Add "medida", ""
    info.Add "equipamento", ""
    info.Add "local", ""
    info.Add "operador", ""
    info.Add "referencia_valor", ""
    info.Add "referencia_unidade", ""
    info.Add "referencia_estado", ""
    info.Add "acao", ""
    info.Add "valor_nominal", ""
    info.Add "tag_gpe", ""
    info.Add "operador_check", 0

    texto = Trim(frase)

    ' ------------------------------------------------------------------------
    ' 02.01 | REGRA 1: abrir/fechar disjuntor | todas as SEs mapeadas
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "^\s*(?:(?:solicitar\s+para\s+COPEL\s+GET)\s+)?" & _
        "(fechar|abrir)\s+(?:o\s+)?disjuntor\s+([A-Za-z0-9_-]+)" & _
        "(?:\s+(?:na|da|em)\s+(?:SE\s+)?" & _
        "(IVP|PGR|PGS|SMS|IRN|IVAIPOR[AÃ]|PONTA\s+GROSSA|IRATI\s+NORTE))?" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("tipo") = "ACAO"
        info("acao") = LCase(correspondencias(0).SubMatches(0))
        info("medida") = "Estado do disjuntor"
        info("equipamento") = _
            "Disjuntor " & correspondencias(0).SubMatches(1)
        info("operador") = "="
        info("operador_check") = 3

        subestacao = NormalizarSiglaSubestacao( _
            correspondencias(0).SubMatches(2) _
        )

        If info("acao") = "fechar" Then
            info("referencia_valor") = "1"
            info("referencia_estado") = "Fechado"
        Else
            info("referencia_valor") = "0"
            info("referencia_estado") = "Aberto"
        End If

        info("tag_gpe") = ObterTagGPEDisjuntor( _
            correspondencias(0).SubMatches(1), _
            subestacao _
        )

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' N-1: individual por subestacao e, em PGR, por grupo de barras.
    ' ------------------------------------------------------------------------
    ' 02.02 | REGRA 2: sistema completo / N-1 | IVP, IRN e PGR
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "\bsistema\s+completo\b.*?\bN-?1\b.*?" & _
        "\b(?:barra|barramento)\s+(BR(?:10|07)B[AB])\b.*?" & _
        "\bSE\s+(.+)$" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("tipo") = "VERIFICACAO_DE_ESTADO"
        info("medida") = "Sistema completo ou N-1"
        info("equipamento") = _
            "Barra " & UCase(correspondencias(0).SubMatches(0))
        info("local") = _
            "SE " & Trim(correspondencias(0).SubMatches(1))
        info("operador") = "="
        info("operador_check") = 3
        info("referencia_valor") = "1"
        info("referencia_estado") = "N-1 atendido"
        info("tag_gpe") = ObterTagCLCN1( _
            info("local"), _
            info("equipamento") _
        )

        If info("tag_gpe") <> "" Then
            Set ExtrairInstrucao = info
            Exit Function
        End If
    End If

    ' Potencia ativa da LT IRN/PGR C1.
    ' ------------------------------------------------------------------------
    ' 02.03 | REGRA 3: potencia ativa LT IRN/PGR C1 | IRN
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "\bfluxo\s+de\s+pot.ncia\s+ativa\b.*?" & _
        "\bLT\s+IRN/PGR\s+C1\b" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("tipo") = "VERIFICACAO_NUMERICA"
        info("medida") = "Potencia ativa"
        info("equipamento") = "LT IRN/PGR C1"
        info("local") = "SE Irati Norte"
        info("operador") = ">"
        info("operador_check") = 1
        info("referencia_valor") = "1"
        info("referencia_unidade") = "MW"
        info("tag_gpe") = ObterTagGPEPotenciaAtivaIRN("C1")

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' Potencia ativa das LTs IVP/PGR.
    ' ------------------------------------------------------------------------
    ' 02.04 | REGRA 4: potencia ativa LT PGR C1/C2 | IVP
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "\bfluxo\s+de\s+pot.ncia\s+ativa\b.*?" & _
        "\bLT\s+PGR\s+C([12])\b" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        circuito = "C" & correspondencias(0).SubMatches(0)

        info("tipo") = "VERIFICACAO_NUMERICA"
        info("medida") = "Potencia ativa"
        info("equipamento") = "LT PGR " & circuito
        info("operador") = ">"
        info("operador_check") = 1
        info("referencia_valor") = "1"
        info("referencia_unidade") = "MW"
        info("tag_gpe") = ObterTagGPEPotenciaAtiva(circuito)

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 02.05 | REGRA 5: diferenca angular LT / barra BR10BA | PGR
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "\bdiferen\S*a\s+angular\b.*?" & _
        "\bLT\s+PGR\s+C([12])\b.*?" & _
        "\bbarra\s+BR10BA\b.*?" & _
        "\bSE\s+Ponta\s+Grossa\b.*?" & _
        "(?:inferior\s+a|menor\s+que)\s*" & _
        "([0-9]+(?:[\.,][0-9]+)?)" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        circuito = "C" & correspondencias(0).SubMatches(0)

        info("tipo") = "VERIFICACAO_NUMERICA"
        info("medida") = "Diferenca angular"
        info("equipamento") = _
            "LT PGR " & circuito & " / Barra BR10BA"
        info("local") = "SE Ponta Grossa"
        info("operador") = "<"
        info("operador_check") = 5
        info("referencia_valor") = _
            correspondencias(0).SubMatches(1)
        info("referencia_unidade") = "graus"
        info("tag_gpe") = ObterTagGPEDiferencaAngular(circuito)

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 02.06 | REGRA 6: religamento tripolar LT PGR C1/C2 | IVP
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "\bhabilitar\s+o\s+religamento\s+tripolar\b.*?" & _
        "\bLT\s+PGR\s+C([12])\b" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        circuito = "C" & correspondencias(0).SubMatches(0)

        info("tipo") = "VERIFICACAO_DE_ESTADO"
        info("medida") = "Religamento tripolar"
        info("equipamento") = "LT PGR " & circuito
        info("operador") = "="
        info("operador_check") = 3
        info("referencia_valor") = "1"
        info("referencia_estado") = "Habilitado"
        info("tag_gpe") = ObterTagGPEReligamentoTripolar(circuito)

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 02.07 | REGRA 7: PSD dos transformadores 1 a 3 | PGR
    ' ------------------------------------------------------------------------
    Set re = NovoRegExp( _
        "\bPSD\s+Habilitado\b.*?" & _
        "\bTransformador\s+0?([123])\b" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("tipo") = "VERIFICACAO_DE_ESTADO"
        info("medida") = "PSD"
        info("equipamento") = _
            "Transformador " & correspondencias(0).SubMatches(0)
        info("local") = "SE Ponta Grossa"
        info("operador") = "="
        info("operador_check") = 3
        info("referencia_valor") = "1"
        info("referencia_estado") = "Habilitado"
        info("tag_gpe") = _
            ObterTagGPEPSD(correspondencias(0).SubMatches(0))

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 02.08 | REGRA 8: comparacao numerica generica (por ultimo)
    ' ------------------------------------------------------------------------
    operador = ExtrairOperador(texto)

    If operador <> "" Then
        info("tipo") = "VERIFICACAO_NUMERICA"
        info("operador") = operador
        info("operador_check") = CodigoOperadorPowerDoc(operador)

        If InStr(1, LCase(texto), "tens", vbTextCompare) > 0 Then
            info("medida") = "Tensao"
        End If

        ExtrairReferenciaNumerica texto, info
        ExtrairBarraELocal texto, info

        info("equipamento") = NormalizarEquipamento( _
            info("equipamento"), _
            info("local"), _
            info("valor_nominal") _
        )

        info("tag_gpe") = ObterTagGPE( _
            info("equipamento"), _
            info("local"), _
            info("valor_nominal") _
        )

        Set ExtrairInstrucao = info
        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 02.09 | Retorno final: preservar os dados acumulados pelas regras
    ' ------------------------------------------------------------------------
    Set ExtrairInstrucao = info
End Function


' ============================================================================
' 03 | EXTRACAO E UTILITARIOS COMPARTILHADOS
' Funcoes sem caminhos de tags.
' ============================================================================

Function ExtrairOperador(ByVal texto)

    texto = LCase(texto)

    If InStr(texto, "igual ou inferior a") > 0 Or _
       InStr(texto, "menor ou igual a") > 0 Then

        ExtrairOperador = "<="

    ElseIf InStr(texto, "igual ou superior a") > 0 Or _
           InStr(texto, "maior ou igual a") > 0 Then

        ExtrairOperador = ">="

    ElseIf InStr(texto, "superior a") > 0 Or _
           InStr(texto, "maior que") > 0 Then

        ExtrairOperador = ">"

    ElseIf InStr(texto, "inferior a") > 0 Or _
           InStr(texto, "menor que") > 0 Then

        ExtrairOperador = "<"

    ElseIf InStr(texto, "igual a") > 0 Then
        ExtrairOperador = "="

    Else
        ExtrairOperador = ""
    End If
End Function

Sub ExtrairReferenciaNumerica(ByVal texto, ByRef info)

    Dim re, correspondencias

    Set re = NovoRegExp( _
        "(?:igual\s+ou\s+inferior\s+a|menor\s+ou\s+igual\s+a|" & _
        "igual\s+ou\s+superior\s+a|maior\s+ou\s+igual\s+a|" & _
        "superior\s+a|maior\s+que|inferior\s+a|menor\s+que|igual\s+a)" & _
        "\s*([0-9]+(?:[\.,][0-9]+)?)\s*([A-Za-z]+)" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("referencia_valor") = correspondencias(0).SubMatches(0)
        info("referencia_unidade") = correspondencias(0).SubMatches(1)
    End If
End Sub

Sub ExtrairBarraELocal(ByVal texto, ByRef info)

    Dim re, correspondencias

    Set re = NovoRegExp( _
        "\bbarra\s+([A-Za-z0-9_-]+)" & _
        "(?:\s+de\s+([0-9]+(?:[\.,][0-9]+)?\s*[A-Za-z]+))?" _
    )

    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("equipamento") = "Barra " & correspondencias(0).SubMatches(0)
        info("valor_nominal") = correspondencias(0).SubMatches(1)
    End If

    Set re = NovoRegExp("\b(?:da|de)\s+SE\s+(.+)$")
    Set correspondencias = re.Execute(texto)

    If correspondencias.Count > 0 Then
        info("local") = "SE " & Trim(correspondencias(0).SubMatches(0))
    End If
End Sub

Function NormalizarSiglaSubestacao(ByVal subestacao)

    Select Case UCase(Trim(subestacao))
        Case "IVP", "IVAIPORA", "IVAIPORÃ"
            NormalizarSiglaSubestacao = "IVP"

        Case "PGR", "PONTA GROSSA"
            NormalizarSiglaSubestacao = "PGR"

        Case "PGS"
            NormalizarSiglaSubestacao = "PGS"

        Case "SMS"
            NormalizarSiglaSubestacao = "SMS"

        Case "IRN", "IRATI NORTE"
            NormalizarSiglaSubestacao = "IRN"

        Case Else
            NormalizarSiglaSubestacao = ""
    End Select
End Function

Function NormalizarEquipamento(ByVal equipamento, ByVal local, ByVal valorNominal)

    Dim chave

    chave = UCase( _
        Trim(local) & "|" & _
        Replace(Trim(valorNominal), " ", "") & "|" & _
        Trim(equipamento) _
    )

    Select Case chave
        Case "SE PONTA GROSSA|230KV|BARRA A"
            NormalizarEquipamento = "Barra BR07BA"

        Case "SE PONTA GROSSA|230KV|BARRA B"
            NormalizarEquipamento = "Barra BR07BB"

        Case Else
            NormalizarEquipamento = equipamento
    End Select
End Function

Function CodigoOperadorPowerDoc(ByVal operador)

    Select Case Trim(operador)
        Case ">"
            CodigoOperadorPowerDoc = 1

        Case ">="
            CodigoOperadorPowerDoc = 2

        Case "="
            CodigoOperadorPowerDoc = 3

        Case "<="
            CodigoOperadorPowerDoc = 4

        Case "<"
            CodigoOperadorPowerDoc = 5

        Case Else
            CodigoOperadorPowerDoc = 0
    End Select
End Function

Function NovoRegExp(ByVal padrao)

    Dim re

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = padrao
    re.Global = False
    re.IgnoreCase = True

    Set NovoRegExp = re
End Function


' ============================================================================
' 04.01 | MAPEAMENTOS: DISJUNTORES
' Agrupados por SE e por numeros compartilhados; Else originais preservados.
' ============================================================================

Function ObterTagGPEDisjuntor(ByVal numeroDisjuntor, ByVal subestacao)

    Select Case CStr(numeroDisjuntor)
        ' 04.01.a | IVP: retorno fixo no codigo atual

        Case "1070"
            ObterTagGPEDisjuntor = _
                "GPE.IVP.IVP.DJC_04.DJ1070.Measurements.PosicaoDisjuntor"

        ' 04.01.b | PGR: retorno fixo no codigo atual

        Case "1060"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.DJC_03.DJ1060.Measurements.PosicaoDisjuntor"

        Case "1072"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.LT_IVP_C2.DJ1072.Measurements.PosicaoDisjuntor"

        Case "732"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.TF_01.DJ732.Measurements.PosicaoDisjuntor"

        Case "1082"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.TF_02.DJ1082.Measurements.PosicaoDisjuntor"

        Case "1080"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.DJC_04.DJ1080.Measurements.PosicaoDisjuntor"

        Case "792"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.TF_03.DJ792.Measurements.PosicaoDisjuntor"

        Case "752"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.LT_PGS.DJ752.Measurements.PosicaoDisjuntor"

        Case "742"
            ObterTagGPEDisjuntor = _
                "GPE.PGR.PGR.LT_SMS.DJ742.Measurements.PosicaoDisjuntor"

        ' 04.01.c | PGR: retorno condicionado a subestacao

        Case "722"
            If UCase(subestacao) = "PGR" Then
                ObterTagGPEDisjuntor = _
                    "GPE.PGR.PGR.LT_IRN_C2.DJ722.Measurements.PosicaoDisjuntor"
            Else
                ObterTagGPEDisjuntor = ""
            End If

        ' 04.01.d | PGS: retorno fixo no codigo atual

        Case "52-158"
            ObterTagGPEDisjuntor = _
                "GPE.PGS.PGS.LT_PGR.DJ52_158.Measurements.PosicaoDisjuntor"

        ' 04.01.e | SMS: retorno fixo no codigo atual

        Case "52-126"
            ObterTagGPEDisjuntor = _
                "GPE.SMS.SMS.LT_PGR.DJ52_126.Measurements.PosicaoDisjuntor"

        ' 04.01.f | IVP/PGR: mesmo numero, selecao pela SE

        Case "1052"
            If UCase(subestacao) = "PGR" Then
                ObterTagGPEDisjuntor = _
                    "GPE.PGR.PGR.LT_IVP_C1.DJ1052.Measurements.PosicaoDisjuntor"
            Else
                ObterTagGPEDisjuntor = _
                    "GPE.IVP.IVP.LT_STIV_C3.DJ1052.Measurements.PosicaoDisjuntor"
            End If

        Case "1062"
            If UCase(subestacao) = "IVP" Then
                ObterTagGPEDisjuntor = _
                    "GPE.IVP.IVP.LT_PGR_C1.DJ1062.Measurements.PosicaoDisjuntor"
            Else
                ObterTagGPEDisjuntor = _
                    "GPE.PGR.PGR.TF_01.DJ1062.Measurements.PosicaoDisjuntor"
            End If

        Case "1112"
            If UCase(subestacao) = "IVP" Then
                ObterTagGPEDisjuntor = _
                    "GPE.IVP.IVP.LT_PGR_C2.DJ1112.Measurements.PosicaoDisjuntor"
            Else
                ObterTagGPEDisjuntor = _
                    "GPE.PGR.PGR.TF_03.DJ1112.Measurements.PosicaoDisjuntor"
            End If

        Case "1110"
            If UCase(subestacao) = "IVP" Then
                ObterTagGPEDisjuntor = _
                    "GPE.IVP.IVP.DJC_06.DJ1110.Measurements.PosicaoDisjuntor"
            Else
                ObterTagGPEDisjuntor = _
                    "GPE.PGR.PGR.DJC_05.DJ1110.Measurements.PosicaoDisjuntor"
            End If

        ' 04.01.g | IRN/PGR: mesmo numero, selecao pela SE

        Case "762"
            If UCase(subestacao) = "IRN" Then
                ObterTagGPEDisjuntor = _
                    "GPE.IRN.IRN.LT_PGR_C2.DJ762.Measurements.PosicaoDisjuntor"
            Else
                ObterTagGPEDisjuntor = _
                    "GPE.PGR.PGR.TF_02.DJ762.Measurements.PosicaoDisjuntor"
            End If

        ' 04.01.h | Numero nao mapeado

        Case Else
            ObterTagGPEDisjuntor = ""
    End Select
End Function


' ============================================================================
' 04.02 | LEITURAS: SE IVP
' Potencia ativa e religamento tripolar.
' ============================================================================

Function ObterTagGPEPotenciaAtiva(ByVal circuito)

    Select Case UCase(Trim(circuito))
        Case "C1"
            ObterTagGPEPotenciaAtiva = _
                "Abs(GPE.IVP.IVP.LT_PGR_C1.SC_MED.Measurements.P)"

        Case "C2"
            ObterTagGPEPotenciaAtiva = _
                "Abs(GPE.IVP.IVP.LT_PGR_C2.SC_MED.Measurements.P)"

        Case Else
            ObterTagGPEPotenciaAtiva = ""
    End Select
End Function

Function ObterTagGPEReligamentoTripolar(ByVal circuito)

    Select Case UCase(Trim(circuito))
        Case "C1"
            ObterTagGPEReligamentoTripolar = _
                "GPE.IVP.IVP.LT_PGR_C1.PT_LT.Protections.[79].Measurements.TRI"

        Case "C2"
            ObterTagGPEReligamentoTripolar = _
                "GPE.IVP.IVP.LT_PGR_C2.PT_LT.Protections.[79].Measurements.TRI"

        Case Else
            ObterTagGPEReligamentoTripolar = ""
    End Select
End Function


' ============================================================================
' 04.03 | LEITURAS: SE PGR
' PSD de transformadores e diferenca angular.
' ============================================================================

Function ObterTagGPEPSD(ByVal numeroTransformador)

    Select Case CStr(numeroTransformador)
        Case "1"
            ObterTagGPEPSD = _
                "GPE.PGR.PGR.TF_01.SC_UCD_AT.Measurements.SCR_AUX_HB"

        Case "2"
            ObterTagGPEPSD = _
                "GPE.PGR.PGR.TF_02.SC_UCD_AT.Measurements.SCR_AUX_HB"

        Case "3"
            ObterTagGPEPSD = _
                "GPE.PGR.PGR.TF_03.SC_UCD_AT.Measurements.SCR_AUX_HB"

        Case Else
            ObterTagGPEPSD = ""
    End Select
End Function

Function ObterTagGPEDiferencaAngular(ByVal circuito)

    Select Case UCase(Trim(circuito))
        Case "C1"
            ObterTagGPEDiferencaAngular = _
                "GPE.PGR.PGR.LT_IVP_C1.SC_MED.Measurements.SC_MED_DIF_ANG"

        Case "C2"
            ObterTagGPEDiferencaAngular = _
                "GPE.PGR.PGR.LT_IVP_C2.SC_MED.Measurements.SC_MED_DIF_ANG"

        Case Else
            ObterTagGPEDiferencaAngular = ""
    End Select
End Function


' ============================================================================
' 04.04 | LEITURAS: SE IRN
' Potencia ativa da LT PGR C1 no terminal IRN.
' ============================================================================

Function ObterTagGPEPotenciaAtivaIRN(ByVal circuito)

    If UCase(Trim(circuito)) = "C1" Then
        ObterTagGPEPotenciaAtivaIRN = _
            "Abs(GPE.IRN.IRN.LT_PGR_C1.SC_MED.Measurements.SC_MED_P)"
    Else
        ObterTagGPEPotenciaAtivaIRN = ""
    End If
End Function


' ============================================================================
' 04.05 | LEITURAS COMPARTILHADAS: TENSAO
' Subgrupos internos na ordem original: IVP, IRN, PGR.
' ============================================================================

Function ObterTagGPE(ByVal equipamento, ByVal local, ByVal valorNominal)

    ' ------------------------------------------------------------------------
    ' 04.05.a | IVP: tensao nominal 525 kV e barras BR10BA/BR10BB
    ' ------------------------------------------------------------------------
    If InStr(UCase(local), "IVAIPOR") > 0 And _
       UCase(Replace(Trim(valorNominal), " ", "")) = "525KV" Then

        Select Case UCase(Trim(equipamento))
            Case "BARRA BR10BA"
                ObterTagGPE = "GPE.IVP.IVP.BR10BA.Terminal1.V_FBC"

            Case "BARRA BR10BB"
                ObterTagGPE = "GPE.IVP.IVP.BR10BB.Terminal1.V_FBC"

            Case Else
                ObterTagGPE = ""
        End Select

        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 04.05.b | IRN: barras BR07BA/BR07BB
    ' ------------------------------------------------------------------------
    If InStr(UCase(local), "IRATI NORTE") > 0 Or _
       InStr(UCase(local), "SE IRN") > 0 Then

        Select Case UCase(Trim(equipamento))
            Case "BARRA BR07BA"
                ObterTagGPE = "GPE.IRN.IRN.BR07BA.Terminal1.V_FB"

            Case "BARRA BR07BB"
                ObterTagGPE = "GPE.IRN.IRN.BR07BB.Terminal1.V_FB"

            Case Else
                ObterTagGPE = ""
        End Select

        Exit Function
    End If

    ' ------------------------------------------------------------------------
    ' 04.05.c | PGR: barras BR07BA/BR07BB/BR10BA/BR10BB
    ' ------------------------------------------------------------------------
    If InStr(UCase(local), "PONTA GROSSA") > 0 Then

        Select Case UCase(Trim(equipamento))
            Case "BARRA BR07BA"
                ObterTagGPE = "GPE.PGR.PGR.BR07BA.Terminal1.V_FB"

            Case "BARRA BR07BB"
                ObterTagGPE = "GPE.PGR.PGR.BR07BB.Terminal1.V_FB"

            Case "BARRA BR10BA"
                ObterTagGPE = "GPE.PGR.PGR.BR10BA.Terminal1.V_FB"

            Case "BARRA BR10BB"
                ObterTagGPE = "GPE.PGR.PGR.BR10BB.Terminal1.V_FB"

            Case Else
                ObterTagGPE = ""
        End Select

        Exit Function
    End If

    ObterTagGPE = ""
End Function


' ============================================================================
' 04.06 | LEITURAS COMPARTILHADAS: N-1
' Reconhecimento da subestacao preservado; em PGR, selecionar BR07/BR10.
' ============================================================================

Function ObterTagCLCN1(ByVal local, ByVal equipamento)

    ' ------------------------------------------------------------------------
    ' 04.06 | N-1: ordem de reconhecimento IVP -> IRN -> PGR
    ' ------------------------------------------------------------------------
    If InStr(UCase(local), "IVAIPOR") > 0 Then
        ObterTagCLCN1 = "CLC.IVP.CLC.IVP.N1.BR10.N1_Atendido"

    ElseIf InStr(UCase(local), "IRATI NORTE") > 0 Or _
           InStr(UCase(local), "SE IRN") > 0 Then
        ObterTagCLCN1 = "CLC.IRN.CLC.IRN.N1.BR07.N1_Atendido"

    ElseIf InStr(UCase(local), "PONTA GROSSA") > 0 Or _
           InStr(UCase(local), "SE PGR") > 0 Then
        Select Case UCase(Trim(equipamento))
            Case "BARRA BR07BA", "BARRA BR07BB"
                ObterTagCLCN1 = "CLC.PGR.CLC.PGR.N1.BR07.N1_Atendido"

            Case "BARRA BR10BA", "BARRA BR10BB"
                ObterTagCLCN1 = "CLC.PGR.CLC.PGR.N1.BR10.N1_Atendido"

            Case Else
                ObterTagCLCN1 = ""
        End Select

    Else
        ObterTagCLCN1 = ""
    End If
End Function


' ============================================================================
' 05 | INFRAESTRUTURA: PARSER JSON
' Classe aspJSON mantida sem alteracoes executaveis.
' ============================================================================

Class aspJSON
    Public data
    Private p_JSONstring
    Private aj_in_string, aj_in_escape, aj_i_tmp, aj_char_tmp
    Private aj_s_tmp, aj_line_tmp, aj_line, aj_lines
    Private aj_currentlevel, aj_currentkey, aj_currentvalue
    Private aj_newlabel, aj_XmlHttp, aj_RegExp, aj_colonfound

    Private Sub Class_Initialize()
        Set data = Collection()

        Set aj_RegExp = New RegExp
        aj_RegExp.Pattern = "\s{0,}(\S{1}[\s,\S]*\S{1})\s{0,}"
        aj_RegExp.Global = False
        aj_RegExp.IgnoreCase = True
        aj_RegExp.Multiline = True
    End Sub

    Private Sub Class_Terminate()
        Set data = Nothing
        Set aj_RegExp = Nothing
    End Sub

    Public Function Collection()
        Set Collection = CreateObject("Scripting.Dictionary")
    End Function

    Public Sub LoadJSONString(ByVal JSONString)
        p_JSONstring = CleanUpJSONstring(JSONString)
        aj_lines = Split(p_JSONstring, Chr(13) & Chr(10))

        Dim level(99)
        aj_currentlevel = 1
        Set level(aj_currentlevel) = data

        For Each aj_line In aj_lines
            aj_currentkey = ""
            aj_currentvalue = ""

            If InStr(aj_line, ":") > 0 Then
                aj_in_string = False
                aj_in_escape = False
                aj_colonfound = False

                For aj_i_tmp = 1 To Len(aj_line)
                    If aj_in_escape Then
                        aj_in_escape = False
                    Else
                        Select Case Mid(aj_line, aj_i_tmp, 1)
                            Case """"
                                aj_in_string = Not aj_in_string

                            Case ":"
                                If Not aj_in_escape And Not aj_in_string Then
                                    aj_currentkey = Left(aj_line, aj_i_tmp - 1)
                                    aj_currentvalue = Mid(aj_line, aj_i_tmp + 1)
                                    aj_colonfound = True
                                    Exit For
                                End If

                            Case "\"
                                aj_in_escape = True
                        End Select
                    End If
                Next

                If aj_colonfound Then
                    aj_currentkey = aj_Strip( _
                        aj_JSONDecode(aj_currentkey), _
                        """" _
                    )

                    If Not level(aj_currentlevel).Exists(aj_currentkey) Then
                        level(aj_currentlevel).Add aj_currentkey, ""
                    End If
                End If
            End If

            If Right(aj_line, 1) = "{" Or Right(aj_line, 1) = "[" Then
                If Len(aj_currentkey) = 0 Then
                    aj_currentkey = level(aj_currentlevel).Count
                End If

                Set level(aj_currentlevel).Item(aj_currentkey) = Collection()
                Set level(aj_currentlevel + 1) = _
                    level(aj_currentlevel).Item(aj_currentkey)

                aj_currentlevel = aj_currentlevel + 1
                aj_currentkey = ""

            ElseIf Right(aj_line, 1) = "}" Or _
                   Right(aj_line, 1) = "]" Or _
                   Right(aj_line, 2) = "}," Or _
                   Right(aj_line, 2) = "]," Then

                aj_currentlevel = aj_currentlevel - 1

            ElseIf Len(Trim(aj_line)) > 0 Then
                If Len(aj_currentvalue) = 0 Then
                    aj_currentvalue = aj_line
                End If

                aj_currentvalue = getJSONValue(aj_currentvalue)

                If Len(aj_currentkey) = 0 Then
                    aj_currentkey = level(aj_currentlevel).Count
                End If

                level(aj_currentlevel).Item(aj_currentkey) = aj_currentvalue
            End If
        Next
    End Sub

    Private Function CleanUpJSONstring(aj_originalstring)
        aj_originalstring = Replace(aj_originalstring, Chr(13) & Chr(10), "")
        aj_originalstring = Mid(aj_originalstring, 2, Len(aj_originalstring) - 2)

        aj_in_string = False
        aj_in_escape = False
        aj_s_tmp = ""

        For aj_i_tmp = 1 To Len(aj_originalstring)
            aj_char_tmp = Mid(aj_originalstring, aj_i_tmp, 1)

            If aj_in_escape Then
                aj_in_escape = False
                aj_s_tmp = aj_s_tmp & aj_char_tmp
            Else
                Select Case aj_char_tmp
                    Case "\"
                        aj_s_tmp = aj_s_tmp & aj_char_tmp
                        aj_in_escape = True

                    Case """"
                        aj_s_tmp = aj_s_tmp & aj_char_tmp
                        aj_in_string = Not aj_in_string

                    Case "{", "["
                        aj_s_tmp = aj_s_tmp & aj_char_tmp & _
                            aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10))

                    Case "}", "]"
                        aj_s_tmp = aj_s_tmp & _
                            aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10)) & _
                            aj_char_tmp

                    Case ","
                        aj_s_tmp = aj_s_tmp & aj_char_tmp & _
                            aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10))

                    Case Else
                        aj_s_tmp = aj_s_tmp & aj_char_tmp
                End Select
            End If
        Next

        CleanUpJSONstring = ""
        aj_s_tmp = Split(aj_s_tmp, Chr(13) & Chr(10))

        For Each aj_line_tmp In aj_s_tmp
            aj_line_tmp = Replace(Replace(aj_line_tmp, Chr(10), ""), Chr(13), "")
            CleanUpJSONstring = CleanUpJSONstring & _
                aj_Trim(aj_line_tmp) & Chr(13) & Chr(10)
        Next
    End Function

    Private Function getJSONValue(ByVal val)
        val = Trim(val)

        If Left(val, 1) = ":" Then val = Mid(val, 2)
        If Right(val, 1) = "," Then val = Left(val, Len(val) - 1)

        val = Trim(val)

        Select Case val
            Case "true"
                getJSONValue = True

            Case "false"
                getJSONValue = False

            Case "null"
                getJSONValue = Null

            Case Else
                If InStr(val, """") = 0 Then
                    If IsNumeric(val) Then
                        getJSONValue = CDbl(val)
                    Else
                        getJSONValue = val
                    End If
                Else
                    If Left(val, 1) = """" Then val = Mid(val, 2)
                    If Right(val, 1) = """" Then val = Left(val, Len(val) - 1)
                    getJSONValue = aj_JSONDecode(Trim(val))
                End If
        End Select
    End Function

    Private Function aj_JSONDecode(ByVal val)
        val = Replace(val, "\""", """")
        val = Replace(val, "\\", "\")
        val = Replace(val, "\/", "/")
        val = Replace(val, "\b", Chr(8))
        val = Replace(val, "\f", Chr(12))
        val = Replace(val, "\n", Chr(10))
        val = Replace(val, "\r", Chr(13))
        val = Replace(val, "\t", Chr(9))
        aj_JSONDecode = Trim(val)
    End Function

    Private Function aj_InlineIf(condition, returntrue, returnfalse)
        If condition Then
            aj_InlineIf = returntrue
        Else
            aj_InlineIf = returnfalse
        End If
    End Function

    Private Function aj_Strip(ByVal val, stripper)
        If Left(val, 1) = stripper Then val = Mid(val, 2)
        If Right(val, 1) = stripper Then val = Left(val, Len(val) - 1)
        aj_Strip = val
    End Function

    Private Function aj_Trim(val)
        aj_Trim = Trim(val)

        Do While Left(aj_Trim, 1) = Chr(9)
            aj_Trim = Mid(aj_Trim, 2)
        Loop

        Do While Right(aj_Trim, 1) = Chr(9)
            aj_Trim = Left(aj_Trim, Len(aj_Trim) - 1)
        Loop

        aj_Trim = Trim(aj_Trim)
    End Function
End Class


' ============================================================================
' 06 | ROTINA VAZIA
' O trecho original terminava em Sub Vazia(), sem End Sub.
' ============================================================================

Sub Vazia()
End Sub
