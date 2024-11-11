Attribute VB_Name = "M�dulo10"
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

            texto1 = ("Ol�, tudo bem?<br><br><u>LEMBRETE: MATERIAL EXPRESSAMENTE PROIBIDO DE SER ENVIADO AOS CLIENTES</u><br/>") _
& "<br> Seguem os vencimentos do dia de hoje. <b> Entrar em contato com o cliente quando a opera��o for N�O MESA </b>, ou seja, n�o existir um broker relacionado.<br>" _
& "<br><i><b>Legenda:</i></b><br/>" & "<br><i>�Financeiro sa�da: Valor de opera��o j� com os ajustes das op��es (aproximado).</i>" _
& "<br><i>�Opera��es sob cust�dia: O resultado final n�o � visto em nota pois o ativo estar� em carteira.</i>" _
& "<br><i>�Booster K.O com R$0,00: Opera��o virou p� e ativo permanece em carteira." _
& "<br><i>�Valor Entrada Estruturada: � o valor do ativo no dia que a opera��o estruturada foi realizada, sendo o valor de refer�ncia para c�lculo do Resultado Estruturada.<br>Observa��o: Valor Entrada Estruturada nem sempre � o mesmo valor do Pre�o M�dio. Se atentar para opera��es sob cust�dia." _


            texto2 = "<br><br><b> OPERA��O RUBI/RUBI BIDIRECIONAL </b><br/>" _
& "<br>Caso a opera��o tenha sido feita COM compra do ativo a venda ser� autom�tica, a mesma acontecer� no leil�o de fechamento a pre�o MOC(fechamento do fixing)." _
& "<br>Opera��es SOB CUST�DIA a venda dos ativos precisa ser feita manualmente, nesta situa��o entrar em contato com a Mesa RV ou vender por conta pr�pria.<br/>" _
& "<br><b> OPERA��O BOOSTER K.O </b><br/>" _
& "<br><b>Cen�rio 1:</b> A opera��o estruturada dobra os ganhos at� o limite da barreira em rela��o ao pre�o de entrada.Tudo que a��o subir at� esse intervalo o cliente tem de ganho dobrado por conta da opera��o estruturada." _
& "<br><b>Cen�rio 2:</b> Caso em qualquer momento a a��o suba mais que a barreira em rela��o ao pre�o de entrada, a opera��o de ganho dobrado deixa de existir e voc� passa a ter um limite de ganho m�ximo da barreira.<br/>" _
& "<br><b> OPERA��ES RISK </b><br/>" _
& "<br> O objetivo da opera��o � obter lucro com a alta do ativo sem desembolso de caixa, mas assumindo riscos na queda." _
& "<br> � uma opera��o apenas para cliente com perfil agressivo, pois se trata de alavancagem." _
& "<br><br><b> OPERA��O PUT  </b><br/>" _
& "<br> A opera��o � utilizada com o objetivo de ganhar com a desvaloriza��o do ativo." _
& "<br> O valor desembolsado no in�cio da opera��o � o pr�mio que consiste tamb�m no risco m�ximo dela." _
& "<br> Caso o ativo tenha virado <i> p� </i> , o resultado final se encontra zerado. <br/>" _

        texto3 = "<br><b> OPERA��ES PUT SPREAD </b><br/>" _
& "<br> Tamb�m chamada de <i> Trava de Baixa </i> possui o objetivo de ganhar com a desvaloriza��o moderada do ativo." _
& "<br>O pre�o de entrada equivale ao risco m�ximo que essa estrutura apresenta." _
& "<br> Lembre-se que ela � bastante utilizada como prote��o. <br/>" _
& "<br><i>Para falar desses cen�rio de risco acionar a Mesa Rv. </i> <br/>" _
& "<br><span style=""color:#FF0000""><i>Este relat�rio � gerencial, desenvolvido pela Mesa Rv para aux�lio no controle das posi��es. Por regras de compliance, n�o podemos enviar esse tipo de material para o cliente final, por configurar confec��o de relat�rio. Essa infra��o est� sujeita a multas pesadas por parte da XP. </i></span style=""color:#FF0000""" _
& "<br><br><i><b>Considera��es: os valores est�o sendo enviados conforme o mercado neste momento, a opera��o � finalizada com o pre�o de FECHAMENTO.Pedimos aten��o para este ponto e para as opera��es sem venda autom�tica, pois o cliente ter� de ter em conta o valor para ajuste das op��es.</b></i>" _

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

