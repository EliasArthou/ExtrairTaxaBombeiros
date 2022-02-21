"""
Extração de IPTUs
"""


def extrairboletos(visual):
    """

    : param caminhobanco: caminho do banco para realizar a pesquisa.
    : param resposta: opção selecionada de extração.
    : param visual: janela a ser manipulada.
    """

    import os
    import web
    import auxiliares as aux
    import sensiveis as senha
    import time
    import messagebox
    import sys
    import datetime
    from bs4 import BeautifulSoup

    codigocliente = ''
    site = None
    listaexcel = []

    resolveucaptcha = False

    # try:
    gerarboleto = not visual.var1.get()
    salvardadospdf = visual.var2.get()

    visual.acertaconfjanela(False)

    if os.path.isfile(aux.caminhoprojeto()+'/'+'Scai.WMB'):
        caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                              tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                              caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
    else:
        if os.path.isdir(aux.caminhoprojeto()):
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto())
        else:
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')])

    if len(caminhobanco) == 0:
        messagebox.msgbox('Selecione o caminho do Banco de Dados!', messagebox.MB_OK, 'Erro Banco')
        visual.manipularradio(True)
        sys.exit()

    resposta = str(visual.radio_valor.get())
    indicecliente = aux.criarinputbox('Cliente de Corte', 'Iniciar a partir de um cliente? (0 fará de todos da lista)', valorinicial='0')

    if indicecliente is not None:
        if not indicecliente.isdigit():
            messagebox.msgbox('Digite um valor válido (precisa ser numérico)!', messagebox.MB_OK, 'Opção Inválida')
            visual.manipularradio(True)
            sys.exit()
    else:
        messagebox.msgbox('Digite o inicío desejado ou deixe 0 (Zero)!', messagebox.MB_OK, 'Opção Inválida!')
        visual.manipularradio(True)
        sys.exit()

    tempoinicio = time.time()

    visual.acertaconfjanela(True)

    visual.mudartexto('labelstatus', 'Executando pesquisa no banco...')

    bd = aux.Banco(caminhobanco)
    indicecliente = str(indicecliente).zfill(4)
    if indicecliente == '0000':
        resultado = bd.consultar("SELECT * FROM [Lista CBM Completa]")
    else:
        resultado = bd.consultar("SELECT * FROM [Lista CBM Completa] WHERE Codigo >= '{codigo}'".format(codigo=indicecliente))

    bd.fecharbanco()

    pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'
    listachaves = ['Código Cliente', 'CBM', 'Área Construída', 'Utilização', 'Faixa', 'Proprietário', 'Endereço', 'Ano', 'Valor', 'Status']
    listaexcel = []
    site = web.TratarSite(senha.site, 'ExtrairBoletoIPTU')

    tempoinicio = time.time()
    for indice, linha in enumerate(resultado):
        resolveucaptcha = False
        if aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
            codigocliente = linha['codigo']
            # ==================== Parte Gráfica =======================================================
            visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
            visual.mudartexto('labelinscricao', 'Inscrição: ' + aux.left(str(linha['cbm']), 6) + '-' + aux.right(str(linha['cbm']), 1))
            visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(resultado)) + '...')
            visual.mudartexto('labelstatus', 'Extraindo boleto...')
            # Atualiza a barra de progresso das transações (Views)
            visual.configurarbarra('barraextracao', len(resultado), indice + 1)
            time.sleep(0.1)
            # ==================== Parte Gráfica =======================================================
            caminhodestino = pastadownload + '/' + codigocliente + '_' + linha['cbm'] + '.pdf'
            while not(resolveucaptcha) and aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
                if aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
                    if not os.path.isfile(caminhodestino) or not gerarboleto:
                        if site is not None:
                            site.fecharsite()
                        site = web.TratarSite(senha.site, 'ExtrairTxIncendio')
                        site.abrirnavegador()
                        if site.url != senha.site or site is None:
                            if site is not None:
                                site.fecharsite()
                            site = web.TratarSite(senha.site, senha.nomeprofile)
                            site.abrirnavegador()

                        if site is not None and site.navegador != -1:
                            # Campo de Inscrição da tela Inicial
                            inscricao = site.verificarobjetoexiste('NAME', 'num_cbmerj')
                            # Campo de dígito verificador
                            dv = site.verificarobjetoexiste('NAME', 'dv_cbmerj')
                            # Testa se tem os dois campos supracitados
                            if inscricao is not None and dv is not None:
                                # "Limpa" o campo de inscrição
                                inscricao.clear()
                                # Preenche o campo de inscrição com os dados do banco de dados (sem o dígito verificador)
                                inscricao.send_keys(aux.left(linha['cbm'], 6))
                                # "Limpa" o campo do dígito verificador
                                dv.clear()
                                # Preenche o campo do dígito verificador com os dados do banco de dados
                                dv.send_keys(aux.right(linha['cbm'], 1))
                                # Baixa a imagem do captcha para ser solucionado
                                achouimagem, baixouimagem = site.baixarimagem('ID', 'cod_seguranca', aux.caminhoprojeto() + '\\' + 'Downloads\\captcha.png')
                                # Verifica se achou e salvou a imagem do captcha
                                if achouimagem and baixouimagem:
                                    # Função que pede como entrada a caixa de texto do código de segurança e o botão de confirmar, como retorno
                                    # dá um booleano que diz se conseguiu resolver o captcha e a mensagem de erro (qualquer uma), caso exista.
                                    resolveucaptcha, mensagemerro = site.resolvercaptcha('NAME', 'txt_Seguranca', 'XPATH', '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/form/input[5]')
                                    # Testa se resolveu o captcha e se deu algum erro
                                    if resolveucaptcha and len(mensagemerro) == 0:
                                        # Verifica se está na página de geração de boletos
                                        if site.url == senha.site:
                                            # Área de retorno dos campos de informação do imposto
                                            # ==================================================================================================================================
                                            site.retornartabela(2)
                                            area = site.verificarobjetoexiste('XPATH',
                                                                              '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[2]/td[2]')
                                            if area is None:
                                                area = ''
                                            else:
                                                area = area.text
                                            utilizacao = site.verificarobjetoexiste('XPATH',
                                                                                    '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[3]/td[2]')
                                            if utilizacao is None:
                                                utilizacao = ''
                                            else:
                                                utilizacao = utilizacao.text

                                            faixa = site.verificarobjetoexiste('XPATH',
                                                                               '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[4]/td[2]')
                                            if faixa is None:
                                                faixa = ''
                                            else:
                                                faixa = faixa.text

                                            proprietario = site.verificarobjetoexiste('XPATH',
                                                                                      '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[5]/td[2]')
                                            if proprietario is None:
                                                proprietario = ''
                                            else:
                                                proprietario = proprietario.text

                                            endereco = site.verificarobjetoexiste('XPATH',
                                                                                  '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[7]/td[2]')
                                            if endereco is None:
                                                endereco = ''
                                            else:
                                                endereco = endereco.text

                                            ano = site.verificarobjetoexiste('XPATH',
                                                                             '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[3]/td[1]/b')
                                            if ano is None:
                                                ano = ''
                                            else:
                                                ano = ano.text

                                            # valor = site.verificarobjetoexiste('XPATH',
                                            #                                    '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[3]/td[2]')

                                            valor = site.verificarobjetoexiste('NAME', 'taxa[Total]')

                                            if valor is None:
                                                valor = ''
                                            else:
                                                valor = valor.text

                                            # ==================================================================================================================================

                                            # Armazena os dados para serem adicionados no excel com o resultado
                                            if float(valor) == 0:
                                                dadoscbm = [codigocliente, linha['cbm'], area, utilizacao, faixa, proprietario, endereco, ano, '0,00', 'Pago ou Isento']
                                            else:
                                                dadoscbm = [codigocliente, linha['cbm'], area, utilizacao, faixa, proprietario, endereco, ano, valor, 'Com Guia']

                                            listaexcel.append(dict(zip(listachaves, dadoscbm)))

                                            # Começa a gerar os boletos (se marcado essa opção)
                                            if gerarboleto:
                                                # Verifica quantos botões tem na tela (dependendo do valor só permite cota única)
                                                botoes = site.verificarobjetoexiste('CLASSNAME', 'botao', itemunico=False)
                                                # Verifica se é para gerar um boleto único ou parcelado
                                                match resposta:
                                                    case '1':
                                                        # Boleto Único
                                                        inicio = 3
                                                        fim = 3

                                                    case '2':
                                                        # Boletos Parcelados
                                                        inicio = 4
                                                        fim = 9
                                                    case _:
                                                        inicio = 0
                                                        fim = 0

                                                # Looping para pressionar um (boleto único) ou diversos botões (parcelado)
                                                for indicebotao in range(inicio, fim):
                                                    # Botão de gerar boleto
                                                    botao = site.verificarobjetoexiste('XPATH',
                                                                                       '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[' + str(indicebotao) + ']/td[5]/form/input[33]')
                                                    # Verifica se o botão existe
                                                    if botao is not None:
                                                        # Clique no botão
                                                        if getattr(sys, 'frozen', False):
                                                            botao.click()
                                                        else:
                                                            site.navegador.execute_script("arguments[0].click()", botao)

                                                        # Verifica a quantidade de abas
                                                        if site.num_abas() > 1:
                                                            # Vai para a última aba
                                                            site.irparaaba(site.num_abas())
                                                            # Verifica se está na tela do boleto
                                                            if site.navegador.current_url == senha.telaboleto:
                                                                # Botão de impressão
                                                                imprimir = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_ImprimirDARM',
                                                                                                      iraoobjeto=True)
                                                                # Verifica se o botão de impressão existe
                                                                if imprimir is not None:
                                                                    # Muda o status na tela
                                                                    visual.mudartexto('labelstatus', 'Salvando Boleto...')
                                                                    # Espera o download
                                                                    site.esperadownloads(pastadownload, 10)
                                                                    # Verifica o último arquivo baixado
                                                                    baixado = aux.ultimoarquivo(pastadownload, '.pdf')
                                                                    # Verifica se o arquivo baixado vem do site
                                                                    if 'DARM_' not in baixado:
                                                                        baixado = ''

                                                                    # Verifica se o arquivo baixado vem do site para renomear, adicionar o cabeçalho
                                                                    # e caminhar
                                                                    if len(baixado) > 0:
                                                                        caminhodestino = aux.to_raw(caminhodestino)
                                                                        # Adiciona o código do cliente ao cabeçalho
                                                                        aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)
                                                        # Verifica se o site está na memória
                                                        if site is not None:
                                                            # Fecha o site
                                                            site.fecharsite()
                                    # Condição se o captcha não for resolvido ou tiver uma mensagem de erro.
                                    else:
                                        # Verifica se o problema não foi o captcha
                                        if resolveucaptcha:
                                            # Linha de erro para carregar no excel
                                            dadosiptu = [codigocliente, linha['cbm'], '', '', '', '', '', '', '', mensagemerro]
                                            # Adiciona o item com o cabeçalho
                                            listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                            # Fecha o site
                                            site.fecharsite()
        else:
            # Mensagem de horário inválido para gerar boleto
            visual.acertaconfjanela(False)
            messagebox.msgbox('Impossível gerar boletos depois das 22:00!', messagebox.MB_OK, 'Horário Inválido')
            # Sai do Looping
            break

    # finally:
    # Verifica se o browser está aberto
    if site is not None:
        # Fecha o browser
        site.fecharsite()

    # Verifica se tem itens para salvar no Excel
    if len(listaexcel) > 0:
        # Atualiza o status na tela de usuário
        visual.mudartexto('labelstatus', 'Salvando lista...')
        # Escreve no arquivo a lista e salva o Excel
        aux.escreverlistaexcelog('Log_' + aux.acertardataatual() + '.xlsx', listaexcel)

    tempofim = time.time()

    hours, rem = divmod(tempofim-tempoinicio, 3600)
    minutes, seconds = divmod(rem, 60)

    visual.manipularradio(True)
    visual.acertaconfjanela(False)

    messagebox.msgbox(
        f'O tempo decorrido foi de: {"{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), int(seconds))}',
        messagebox.MB_OK, 'Tempo Decorrido')

