Attribute VB_Name = "Módulo10"
    Sub EMAIL_DIA_AUTOMATICO()
    Dim nome As String
    Dim EnviarPara As String
    Dim titulo As String
    Dim Outlook As Outlook.Application
         
    startrow = 2
    nextrow = startrow + 1
    nome = Cells(startrow, 1)
    EnviarPara = Cells(startrow, 25)
    titulo = Cells(startrow, 26) 'Inicializacao da macro.
    
    'Outlook.MailItem.SenderEmailAddress = "mesarv@manchesterinvest.com.br"
    
    Do Until ActiveSheet.Cells(startrow, 1) = ""
        nomeOld = nome
        nome = Cells(startrow, 1)
        
        Do While Cells(startrow, 1).Value = Cells(nextrow, 1).Value
                startrow = startrow + 1
                nome = Cells(nextrow, 1)
                nextrow = nextrow + 1
                EnviarPara = Cells(startrow, 25)
                titulo = Cells(startrow, 26)
            Loop
            
        With ActiveSheet
            .Range("$A$1:$AW$15001").AutoFilter Field:=1, Criteria1:=nome
            If Not .AutoFilterMode Then .UsedRange.AutoFilter
            lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
            
            Set objeto_outlook = CreateObject("Outlook.Application")
            Set Email = objeto_outlook.CreateItem(0)
            
            
            Email.To = EnviarPara
            Email.Subject = titulo
            Email.Display
            
            texto1 = ("Olá, tudo bem?<br><br><u>LEMBRETE: MATERIAL EXPRESSAMENTE PROIBIDO DE SER ENVIADO AOS CLIENTES</u><br/>") _
& "<br> Seguem os vencimentos do dia de hoje. <b> Entrar em contato com o cliente quando a operação for NÃO MESA </b>, ou seja, não existir um broker relacionado.<br>" _
& "<br><i><b>Legenda:</i></b><br/>" & "<br><i>•Financeiro saída: Valor de operação já com os ajustes das opções (aproximado).</i>" _
& "<br><i>•Operações sob custódia: O resultado final não é visto em nota pois o ativo estará em carteira.</i>" _
& "<br><i>•Booster K.O com R$0,00: Operação virou pó e ativo permanece em carteira." _
& "<br><i>•Valor Entrada Estruturada: É o valor do ativo no dia que a operação estruturada foi realizada, sendo o valor de referência para cálculo do Resultado Estruturada.<br>Observação: Valor Entrada Estruturada nem sempre é o mesmo valor do Preço Médio. Se atentar para operações sob custódia." _
            

            texto2 = "<br><br><b> OPERAÇÃO RUBI/RUBI BIDIRECIONAL </b><br/>" _
& "<br>Caso a operação tenha sido feita COM compra do ativo a venda será automática, a mesma acontecerá no leilão de fechamento a preço MOC(fechamento do fixing)." _
& "<br>Operações SOB CUSTÓDIA a venda dos ativos precisa ser feita manualmente, nesta situação entrar em contato com a Mesa RV ou vender por conta própria.<br/>" _
& "<br><b> OPERAÇÃO BOOSTER K.O </b><br/>" _
& "<br><b>Cenário 1:</b> A operação estruturada dobra os ganhos até o limite da barreira em relação ao preço de entrada.Tudo que ação subir até esse intervalo o cliente tem de ganho dobrado por conta da operação estruturada." _
& "<br><b>Cenário 2:</b> Caso em qualquer momento a ação suba mais que a barreira em relação ao preço de entrada, a operação de ganho dobrado deixa de existir e você passa a ter um limite de ganho máximo da barreira.<br/>" _
& "<br><b> OPERAÇÕES RISK </b><br/>" _
& "<br> O objetivo da operação é obter lucro com a alta do ativo sem desembolso de caixa, mas assumindo riscos na queda." _
& "<br> É uma operação apenas para cliente com perfil agressivo, pois se trata de alavancagem." _
& "<br><br><b> OPERAÇÃO PUT  </b><br/>" _
& "<br> A operação é utilizada com o objetivo de ganhar com a desvalorização do ativo." _
& "<br> O valor desembolsado no início da operação é o prêmio que consiste também no risco máximo dela." _
& "<br> Caso o ativo tenha virado <i> pó </i> , o resultado final se encontra zerado. <br/>" _

        texto3 = "<br><b> OPERAÇÕES PUT SPREAD </b><br/>" _
& "<br> Também chamada de <i> Trava de Baixa </i> possui o objetivo de ganhar com a desvalorização moderada do ativo." _
& "<br>O preço de entrada equivale ao risco máximo que essa estrutura apresenta." _
& "<br> Lembre-se que ela é bastante utilizada como proteção. <br/>" _
& "<br><i>Para falar desses cenário de risco acionar a Mesa Rv. </i> <br/>" _
& "<br><span style=""color:#FF0000""><i>Este relatório é gerencial, desenvolvido pela Mesa Rv para auxílio no controle das posições. Por regras de compliance, não podemos enviar esse tipo de material para o cliente final, por configurar confecção de relatório. Essa infração está sujeita a multas pesadas por parte da XP. </i></span style=""color:#FF0000""" _
& "<br><br><i><b>Considerações: os valores estão sendo enviados conforme o mercado neste momento, a operação é finalizada com o preço de FECHAMENTO.Pedimos atenção para este ponto e para as operações sem venda automática, pois o cliente terá de ter em conta o valor para ajuste das opções.</b></i>" _
           
           Email.HTMLBody = texto1 & RangetoHTML(Range("C1:W" & lastrow)) & texto2 & texto3 & "<br>" & Email.HTMLBody _
           
           'Email.HTMLBody = texto1 & RangetoHTML(Range("C1:W" & lastrow)) & texto2 & texto3 & "<br>" & Email.HTMLBody

            'Email.Send
            
            
            
            If .Cells.AutoFilter Then .Cells.AutoFilter
            startrow = startrow + 1
            nextrow = nextrow + 1
            EnviarPara = Cells(startrow, 25)
            titulo = Cells(startrow, 26)
        End With
        Loop
End Sub

