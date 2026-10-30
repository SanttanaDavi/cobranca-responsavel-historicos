# -*- coding: utf-8 -*-
# Edição pública configurável. Autoria: David Santana.
# Não contém as configurações da instalação original.
import configuracao as config
config.exigir_autorizacao()

"""Rotina avulsa de cobrança: Selenium, confirmação por etapa e retomada local."""
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
import hashlib
import json
import os
import re
import shutil
import tempfile
import unicodedata


caminho_driver = config.caminho("DRIVER")
arquivo_planilha = config.caminho("PLANILHA")
arquivo_log_erros = config.caminho("LOG")
USUARIO = config.valor("USUARIO")
SENHA = config.valor("SENHA")
LETRA_RESPONSAVEL = config.valor("RESPONSAVEL")
RESPONSAVEL_ESPERADO = config.valor("RESPONSAVEL")
TIPO_HISTORICO = config.valor("TIPO_HISTORICO")
TEMPO_ESPERA = 15
DOMINIO_TASY = config.host_sistema()
navegador = None
aba_principal = None
pausado = False


def log(texto):
    print(f'[{datetime.now():%H:%M:%S}] {texto}', flush=True)


def normalizar(texto):
    texto = unicodedata.normalize('NFKD', str(texto or ''))
    return re.sub(r'[^a-z0-9]', '', ''.join(c for c in texto if not unicodedata.combining(c)).lower())


def texto_celula(valor):
    if valor is None:
        return ''
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def registrar_erro_arquivo(numero, texto_erro):
    
    with open(arquivo_log_erros, 'a', encoding='utf-8') as arquivo:
        arquivo.write(f'[{datetime.now().isoformat(timespec="seconds")}] Título {numero}: {texto_erro}\n')
        arquivo.flush()
        os.fsync(arquivo.fileno())


def visiveis(raiz, por, seletor):
    encontrados = []
    for elemento in raiz.find_elements(por, seletor):
        try:
            if elemento.is_displayed():
                encontrados.append(elemento)
        except StaleElementReferenceException:
            continue
    return encontrados


def unico(raiz, por, seletor):
    elementos = visiveis(raiz, por, seletor)
    if len(elementos) > 1:
        raise RuntimeError('Mais de um elemento corresponde ao campo solicitado. Nenhuma escolha automática foi feita.')
    return elementos[0] if elementos else False


def marcar_etapa(texto):
    if navegador is not None:
        navegador._etapa_cobranca = texto
    log('ETAPA: ' + texto)


def item_menu_cobranca(texto):
    
    rotulos = visiveis(navegador, By.CSS_SELECTOR, '.wpopupmenu__label')
    itens = [e for e in rotulos if normalizar(e.text) == normalizar(texto)
             or normalizar(e.get_attribute('aria-label')) == normalizar(texto)]
    if not itens:
        itens = [e for e in visiveis(navegador, By.CSS_SELECTOR, '[role="menuitem"]')
                 if normalizar(e.text) == normalizar(texto)
                 or normalizar(e.get_attribute('aria-label')) == normalizar(texto)]
    itens = list({e.id: e for e in itens}.values())
    if len(itens) > 1:
        raise RuntimeError('Mais de um item visível corresponde ao menu ' + texto + '.')
    return itens[0] if itens else False


def diagnosticar_operacao_cobranca():
    """Registra controles e rótulos de menu, sem valores dos campos ou imagens."""
    try:
        menus = sorted({e.text.strip() for e in visiveis(navegador, By.CSS_SELECTOR, '.wpopupmenu__label')
                        if e.text.strip()})
        formularios = []
        for janela in dialogos():
            campos = [{'tag': e.tag_name, 'nome': e.get_attribute('name'),
                       'tipo': e.get_attribute('type'), 'visivel': e.is_displayed(),
                       'habilitado': e.is_enabled()}
                      for e in janela.find_elements(By.CSS_SELECTOR, 'input, textarea, select')]
            botoes = [e.text.strip() for e in visiveis(janela, By.TAG_NAME, 'button')
                      if e.text.strip() in ('Alterar', 'Gerar', 'Cancelar', 'Filtrar', 'OK', 'Ok')]
            formularios.append({'campos': campos, 'botoes': botoes,
                'listas': len(janela.find_elements(By.CSS_SELECTOR, 'tasy-listbox, tasy-wlistbox'))})
        return json.dumps({'menus': menus, 'formularios': formularios,
            'opcoes_visiveis': len(visiveis(navegador, By.CSS_SELECTOR, '[role="option"]'))}, ensure_ascii=False)
    except Exception as erro:
        return 'Diagnóstico indisponível: ' + type(erro).__name__


def aguardar(funcao, tempo=TEMPO_ESPERA):
    try:
        return WebDriverWait(navegador, tempo, poll_frequency=0.2,
            ignored_exceptions=(StaleElementReferenceException,)).until(lambda _: funcao())
    except TimeoutException as erro:
        etapa = getattr(navegador, '_etapa_cobranca', 'carregamento do Tasy')
        mensagem = 'Tempo limite na etapa: ' + etapa + '.'
        if tempo >= 10:
            mensagem += ' Diagnóstico: ' + diagnosticar_operacao_cobranca()
        raise TimeoutException(mensagem) from erro


def aguardar_foco():
    """Confere a aba e o foco antes das ações, sem trocar repetidamente de janela."""
    global pausado
    def ativa():
        global pausado
        if navegador.current_window_handle != aba_principal:
            raise RuntimeError('A aba controlada mudou. Não enviei comandos para outra janela.')
        correto = navegador.execute_script(
            'return location.hostname === arguments[0] && document.hasFocus();', DOMINIO_TASY)
        if not correto and not pausado:
            log('PAUSADO: volte à aba do Tasy para continuar. Ctrl+C preserva as etapas já registradas.')
            pausado = True
        if correto and pausado:
            log('Tasy ativo novamente; continuando.')
            pausado = False
        return correto
    WebDriverWait(navegador, 86400, poll_frequency=0.5).until(lambda _: ativa())


def clicar(elemento):
    aguardar_foco()
    bloqueios = elemento.find_elements(By.XPATH,
        "ancestor-or-self::*[@aria-disabled='true' or (self::button and @disabled) "
        "or contains(concat(' ', normalize-space(@class), ' '), ' disabled ')]")
    if bloqueios or not elemento.is_displayed() or not elemento.is_enabled():
        raise RuntimeError('O controle está indisponível; nenhum clique foi enviado.')
    elemento.click()


def preencher(elemento, texto):
    aguardar_foco()
    esperado = str(texto).replace('\r\n', '\n')
    elemento.click()
    
    elemento.send_keys(Keys.CONTROL, 'a', Keys.NULL, Keys.BACKSPACE, esperado)
    recebido = str(elemento.get_property('value') or '').replace('\r\n', '\n')
    if recebido != esperado:
        raise RuntimeError('O texto integral do campo não corresponde ao texto solicitado. Nenhum botão de gravação foi acionado.')


def dialogos():
    return visiveis(navegador, By.CSS_SELECTOR, 'div.ngdialog')


def dialogo(titulo):
    encontrados = []
    for item in dialogos():
        if any(normalizar(linha) == normalizar(titulo) for linha in item.text.splitlines()):
            encontrados.append(item)
    if len(encontrados) > 1:
        raise RuntimeError('Mais de uma janela com o mesmo título está aberta.')
    return encontrados[0] if encontrados else False


def botao(raiz, texto):
    return unico(raiz, By.XPATH, './/button[normalize-space(.)=' + json.dumps(texto, ensure_ascii=False) + ']')


def fechar_cancelando(janela):
    controle = botao(janela, 'Cancelar')
    if not controle:
        raise RuntimeError('Não encontrei Cancelar nesta janela; não fechei uma confirmação desconhecida.')
    clicar(controle)
    def fechada():
        try:
            return not janela.is_displayed()
        except StaleElementReferenceException:
            return True
    aguardar(fechada)


def verificar_erro_tasy():
    erros = []
    for janela in dialogos():
        texto = janela.text.strip()
        normal = normalizar(texto)
        if 'houveumerro' in normal or 'naopermit' in normal or 'naoepermit' in normal:
            erros.append(texto)
    if erros:
        raise RuntimeError('Tasy: ' + ' | '.join(erros))


def abrir_funcao():
    
    if visiveis(navegador, By.CSS_SELECTOR, 'div.slick-row'):
        return
    aba = unico(navegador, By.XPATH,
        '//span[contains(@class,"w-feature-app__name") and normalize-space(.)="Cobrança"]')
    if not aba:
        log('Selecione o perfil Cobrança ou Contas a Receber e deixe a função Cobrança disponível no menu.')
        aba = aguardar(lambda: unico(navegador, By.XPATH,
            '//span[contains(@class,"w-feature-app__name") and normalize-space(.)="Cobrança"]'), 180)
    clicar(aba)


def filtro_cobranca_em_abertura():
    """Reconhece exclusivamente a camada de fundo do filtro do Tasy."""
    camadas = visiveis(navegador, By.CSS_SELECTOR, '.ngdialog-overlay')
    return bool(camadas) and all('hideFilter' in (e.get_attribute('ng-click') or '') for e in camadas)


def dialogos_bloqueadores():
    """Reconhece o filtro por seu campo real e pelos controles do formulário."""
    bloqueadores = []
    for janela in dialogos():
        campos = janela.find_elements(By.CSS_SELECTOR, 'input[name="NR_TITULO"]')
        filtrar = janela.find_elements(By.XPATH, './/button[normalize-space(.)="Filtrar"]')
        grupos_filtro = janela.find_elements(By.CSS_SELECTOR, '[w-detail-parent-type="filter"]')
        if not (campos and (filtrar or grupos_filtro or filtro_cobranca_em_abertura())):
            bloqueadores.append(janela)
    return bloqueadores


def diagnosticar_pesquisa_titulo():
    """Informa somente estrutura de controles; não lê valores, HTML completo ou imagens."""
    controles = []
    for elemento in navegador.find_elements(By.CSS_SELECTOR, 'input[name="NR_TITULO"]'):
        try:
            controles.append({'nome': elemento.get_attribute('name'),
                'visivel': elemento.is_displayed(), 'habilitado': elemento.is_enabled(),
                'somente_leitura': elemento.get_attribute('aria-readonly')})
        except StaleElementReferenceException:
            controles.append({'controle_redesenhado': True})
    botoes = []
    conhecidos = {'Filtrar', 'Cancelar', 'Ok', 'OK', 'Alterar', 'Gerar'}
    for elemento in visiveis(navegador, By.TAG_NAME, 'button'):
        texto = elemento.text.strip()
        if texto in conhecidos:
            botoes.append(texto)
    return json.dumps({'campos_titulo': controles, 'botoes_reconhecidos': botoes,
        'janelas_visiveis': len(dialogos()), 'janelas_bloqueadoras': len(dialogos_bloqueadores()),
        'camada_do_filtro': filtro_cobranca_em_abertura()}, ensure_ascii=False)


def pesquisar_titulo(numero):
    def campo_disponivel():
        
        
        if dialogos_bloqueadores():
            return False
        campo = unico(navegador, By.CSS_SELECTOR, 'input[name="NR_TITULO"]')
        if campo and campo.is_enabled() and campo.get_attribute('aria-readonly') != 'true':
            return campo
        return False

    def aguardar_campo():
        try:
            return aguardar(campo_disponivel)
        except TimeoutException as erro:
            raise RuntimeError('Não confirmei o campo Número do título após aguardar o formulário. '
                'Diagnóstico: ' + diagnosticar_pesquisa_titulo()) from erro

    aguardar_foco()
    try:
        primeiro_acesso = not getattr(navegador, '_filtro_cobranca_confirmado', False)
        campo = aguardar(campo_disponivel, 3 if primeiro_acesso else 0.4)
        log('Filtro reconhecido: aproveitando o campo Número do título.')
    except TimeoutException:
        if filtro_cobranca_em_abertura() or dialogos():
            log('Aguardando a renderização e a identificação da janela antes de pesquisar.')
            campo = aguardar_campo()
        else:
            if visiveis(navegador, By.CSS_SELECTOR, '.ngdialog-overlay'):
                raise RuntimeError('Camada não identificada antes da pesquisa. Diagnóstico: '
                    + diagnosticar_pesquisa_titulo())
            funil = aguardar(lambda: unico(navegador, By.CSS_SELECTOR,
                'tasy-wlabel.filter-icon[ng-click*="toggleFilter"]'))
            try:
                if filtro_cobranca_em_abertura() or dialogos():
                    log('A janela abriu durante a transição; aguardando seu campo.')
                else:
                    clicar(funil)
                    log('Filtro solicitado pelo ícone do painel.')
            except ElementClickInterceptedException:
                if not filtro_cobranca_em_abertura() and not dialogos():
                    raise
                log('O clique coincidiu com a abertura da janela; aguardando a identificação do formulário.')
            campo = aguardar_campo()

    navegador._filtro_cobranca_confirmado = True
    preencher(campo, numero)
    raizes = campo.find_elements(By.XPATH,
        'ancestor::*[.//button[normalize-space(.)="Filtrar"]][1]')
    if len(raizes) != 1:
        raise RuntimeError('Não identifiquei o botão Filtrar no mesmo formulário do título. Diagnóstico: '
            + diagnosticar_pesquisa_titulo())
    filtrar = aguardar(lambda: botao(raizes[0], 'Filtrar'))
    clicar(filtrar)
    aguardar(lambda: not visiveis(navegador, By.CSS_SELECTOR, '.ngdialog-overlay'))
    log(f'Filtro aplicado ao título {numero}; aguardando a linha exata na grade.')
    try:
        encontrada = aguardar(lambda: linha_titulo(numero))
        log(f'Título {numero} localizado na coluna real da grade.')
        return encontrada
    except TimeoutException as erro:
        raise RuntimeError('O filtro foi aplicado, mas a linha exata não foi confirmada. Diagnóstico da grade: '
            + diagnosticar_grade_titulo(numero)) from erro


def cabecalho_corresponde(elemento, legendas):
    nomes = {normalizar(n) for n in legendas}
    identificador = elemento.get_attribute('id') or ''
    if nomes.intersection({'titulo', 'numerodotitulo', 'nrtitulo'}) and identificador.endswith('NR_TITULO'):
        return True
    textos = [elemento.get_attribute('textContent'), elemento.get_attribute('title'),
              elemento.get_attribute('aria-label')]
    return any(normalizar(t) in nomes for t in textos if t)


def capturar_grade_cobranca(numero='', legendas=()):
    
    estado = navegador.execute_script('// Consulta somente leitura: não altera o DOM, não dispara eventos nem faz cliques.\nconst numero = String(arguments[0] || \'\');\nconst aliasesExtras = arguments[1] || [];\nconst norm = s => String(s || \'\').normalize(\'NFD\').replace(/[\\u0300-\\u036f]/g, \'\')\n  .toLowerCase().replace(/[^a-z0-9]/g, \'\');\nconst visivel = e => !!e && !!e.getClientRects().length &&\n  getComputedStyle(e).display !== \'none\' && getComputedStyle(e).visibility !== \'hidden\';\nconst texto = e => e ? String(e.innerText || e.textContent || \'\').trim() : \'\';\nconst corresponde = (e, nomes) => {\n  const ns = nomes.map(norm);\n  if (ns.some(n => [\'titulo\', \'numerodotitulo\', \'nrtitulo\'].includes(n)) &&\n      String(e.id || \'\').endsWith(\'NR_TITULO\')) return true;\n  return [e.textContent, e.getAttribute(\'title\'), e.getAttribute(\'aria-label\')]\n    .some(t => t && ns.includes(norm(t)));\n};\nconst hs = [...document.querySelectorAll(\'[role="columnheader"]\')]\n  .filter(e => visivel(e) && corresponde(e, [\'Título\', \'Número do título\', \'NR_TITULO\']));\nif (hs.length !== 1) return {grade: null, linha: null, erro: \'Cabeçalho do título ausente ou ambíguo.\'};\nconst grade = hs[0].closest(\'[role="grid"]\');\nif (!grade) return {grade: null, linha: null, erro: \'Grade do título não identificada.\'};\nconst headers = [...grade.querySelectorAll(\'[role="columnheader"]\')];\nconst indice = aliases => {\n  const encontrados = headers.map((e, i) => ({e, i})).filter(x => corresponde(x.e, aliases));\n  if (encontrados.length !== 1) return null;\n  const h = encontrados[0];\n  const n = Number(h.e.getAttribute(\'aria-colindex\'));\n  return Number.isInteger(n) && n > 0 ? n : h.i + 1;\n};\nconst idxTitulo = indice([\'Título\', \'Número do título\', \'NR_TITULO\']);\nconst linhas = [...grade.querySelectorAll(\'div.slick-row\')].filter(visivel);\nconst encontrados = [];\nif (numero && idxTitulo) for (const linha of linhas) {\n  const cs = [...linha.querySelectorAll(\'[role="gridcell"][aria-colindex="\' + idxTitulo + \'"]\')];\n  if (cs.length === 1 && texto(cs[0]) === numero) encontrados.push({linha, celula: cs[0]});\n}\nconst retorno = {grade, linha: null, celula: null, indice_extra: aliasesExtras.length ? indice(aliasesExtras) : null,\n  quantidade_titulos: encontrados.length, outra_linha: false, selecionados: null, status: [],\n  responsavel: null, historico_texto: null, historico_tipo: null, erro: null};\nif (encontrados.length > 1) {retorno.erro = \'O título aparece em mais de uma linha.\'; return retorno;}\nif (encontrados.length === 0) return retorno;\nconst escolhido = encontrados[0];\nretorno.linha = escolhido.linha;\nretorno.celula = escolhido.celula;\nconst r = escolhido.linha.getBoundingClientRect();\nconst rowIndex = escolhido.linha.getAttribute(\'aria-rowindex\');\nconst partes = linhas.filter(e => {\n  const rect = e.getBoundingClientRect();\n  const ri = e.getAttribute(\'aria-rowindex\');\n  return Math.abs(rect.top - r.top) <= 1 && Math.abs(rect.height - r.height) <= 1 &&\n    !(rowIndex && ri && rowIndex !== ri);\n});\nretorno.outra_linha = partes.length !== linhas.length;\nconst celulas = aliases => {\n  const i = indice(aliases);\n  if (!i) return [];\n  return [...new Set(partes.flatMap(e => [...e.querySelectorAll(\n    \'[role="gridcell"][aria-colindex="\' + i + \'"]\')]).filter(visivel))];\n};\nconst valor = aliases => { const cs = celulas(aliases); return cs.length === 1 ? texto(cs[0]) : null; };\nconst status = celulas([\'Status\']);\nif (status.length === 1) retorno.status = [...new Set([...status[0].querySelectorAll(\n  \'.legend-image-cell[role="img"][aria-label]\')].filter(visivel).map(e => e.getAttribute(\'aria-label\').trim()))];\nretorno.responsavel = valor([\'Responsável cobrança\', \'Responsável pela cobrança\', \'Responsável\']);\nretorno.historico_texto = valor([\'Histórico complementar\', \'Último histórico complementar\']);\nretorno.historico_tipo = valor([\'Tipo histórico\', \'Último histórico\']);\nconst xpath = document.evaluate(\'//*[starts-with(normalize-space(text()), "Selecionados:")]\',\n  document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);\nconst contagens = [];\nfor (let i = 0; i < xpath.snapshotLength; i++) {\n  const e = xpath.snapshotItem(i);\n  if (!visivel(e)) continue;\n  const m = texto(e).match(/^Selecionados:\\s*(\\d+)$/);\n  if (m) contagens.push(Number(m[1]));\n}\nif (contagens.length) retorno.selecionados = Math.max(...contagens);\nreturn retorno;\n', str(numero), list(legendas))
    if not isinstance(estado, dict):
        raise RuntimeError('Não recebi a estrutura da grade de cobrança.')
    if estado.get('quantidade_titulos', 0) > 1:
        raise RuntimeError('O título aparece em mais de uma linha. Nenhuma alteração foi solicitada.')
    return estado


def campo_dialogo(titulo, seletor):
    janela = dialogo(titulo)
    return unico(janela, By.CSS_SELECTOR, seletor) if janela else False


def botao_dialogo_habilitado(titulo, legenda):
    janela = dialogo(titulo)
    controle = botao(janela, legenda) if janela else False
    return controle if controle and controle.is_enabled() else False


def grade_cobranca():
    return capturar_grade_cobranca().get('grade') or False


def partes_linha(linha):
    """Reúne painéis congelados da mesma linha lógica, sem juntar linhas distintas."""
    grades = linha.find_elements(By.XPATH, './ancestor::*[@role="grid"][1]')
    if len(grades) != 1:
        return [linha]
    referencia = linha.rect
    indice = linha.get_attribute('aria-rowindex')
    partes = []
    for outra in visiveis(grades[0], By.CSS_SELECTOR, 'div.slick-row'):
        retangulo = outra.rect
        if abs(retangulo['y'] - referencia['y']) > 1 or abs(retangulo['height'] - referencia['height']) > 1:
            continue
        outro_indice = outra.get_attribute('aria-rowindex')
        if indice and outro_indice and indice != outro_indice:
            continue
        partes.append(outra)
    return partes or [linha]


def celulas_linha(linha, seletor):
    if not seletor:
        return []
    celulas = {}
    for parte in partes_linha(linha):
        for celula in parte.find_elements(By.CSS_SELECTOR, seletor):
            if celula.is_displayed():
                celulas[celula.id] = celula
    return list(celulas.values())


def diagnosticar_grade_titulo(numero):
    grade = grade_cobranca()
    raiz = grade if grade else navegador
    headers = []
    for posicao, e in enumerate(raiz.find_elements(By.CSS_SELECTOR, '[role="columnheader"]'), 1):
        headers.append({'posicao': posicao, 'id': e.get_attribute('id'),
            'legenda': (e.get_attribute('textContent') or '').strip()[:100],
            'indice': e.get_attribute('aria-colindex'), 'visivel': e.is_displayed()})
    linhas = visiveis(raiz, By.CSS_SELECTOR, 'div.slick-row')
    exatas = sum(1 for e in visiveis(raiz, By.CSS_SELECTOR, '.slick-row [role="gridcell"]')
                 if texto_celula(e.text) == numero)
    return json.dumps({'grade_identificada': bool(grade), 'cabecalhos': headers,
        'partes_de_linhas_visiveis': len(linhas), 'celulas_com_titulo_exato': exatas,
        'seletor_titulo': coluna(('Título', 'Número do título', 'NR_TITULO'))}, ensure_ascii=False)


def coluna(legendas):
    indice = capturar_grade_cobranca(legendas=legendas).get('indice_extra')
    return f'[role="gridcell"][aria-colindex="{indice}"]' if indice else None


def linha_titulo(numero):
    return capturar_grade_cobranca(numero).get('linha') or False


def status_titulo(numero):
    def ler():
        estado = capturar_grade_cobranca(numero)
        return estado if estado.get('linha') else False
    estado = aguardar(ler)
    status = estado.get('status', [])
    if len(status) != 1:
        raise RuntimeError('Status ausente ou ambíguo para o título exato. Nenhuma gravação foi solicitada.')
    return status[0]


def menu_titulo(numero, opcao, alterar=False):
    def alvo_confirmado():
        estado = capturar_grade_cobranca(numero)
        if not estado.get('celula'):
            return False
        if estado.get('outra_linha'):
            raise RuntimeError('A pesquisa contém outro título. Não abri uma alteração em múltiplos registros.')
        if estado.get('selecionados') is not None and estado['selecionados'] > 1:
            raise RuntimeError('Há mais de um registro selecionado. Nenhuma gravação foi solicitada.')
        if estado.get('status') != ['Pendente']:
            raise RuntimeError('O título não está inequivocamente Pendente antes da ação.')
        return estado['celula']
    marcar_etapa(f'confirmar título {numero}, status e seleção antes do botão direito')
    celula = aguardar(alvo_confirmado)
    aguardar_foco()
    ActionChains(navegador).context_click(celula).perform()
    navegador._menu_cobranca_aberto = True
    aguardar(lambda: item_menu_cobranca('Alterar') or item_menu_cobranca('Gerar históricos'))
    aguardar(alvo_confirmado)
    if alterar:
        marcar_etapa('abrir o submenu Alterar')
        item = aguardar(lambda: item_menu_cobranca('Alterar'))
        aguardar_foco()
        ActionChains(navegador).move_to_element(item).perform()
    marcar_etapa('acionar a opção ' + opcao)
    clicar(aguardar(lambda: item_menu_cobranca(opcao)))
    navegador._menu_cobranca_aberto = False


def fechar_erro_conhecido():
    
    for janela in dialogos():
        if 'houveumerro naexecucaodaaplicacao'.replace(' ', '') in normalizar(janela.text):
            botoes = [e for e in visiveis(janela, By.TAG_NAME, 'button') if normalizar(e.text) == 'ok']
            if len(botoes) != 1:
                raise RuntimeError('O aviso de erro conhecido não possui um único OK.')
            clicar(botoes[0])
            aguardar(lambda: not any('houveumerro naexecucaodaaplicacao'.replace(' ', '')
                     in normalizar(j.text) for j in dialogos()))


def componente_lista(janela, atributo=None):
    if atributo:
        grupo = unico(janela, By.CSS_SELECTOR, f'[w-attr-name="{atributo}"]')
        if grupo:
            return grupo
    listas = visiveis(janela, By.CSS_SELECTOR, 'tasy-listbox, tasy-wlistbox')
    if len(listas) != 1:
        raise RuntimeError('Não identifiquei uma única lista no formulário. É necessário conferir seu HTML.')
    return listas[0]


def escolher_opcao(janela, atributo=None, codigo=None):
    nome_campo = 'tipo de histórico' if codigo is not None else 'responsável'
    titulo = 'Gerar históricos' if codigo is not None else 'Alterar responsável cobrança'
    def grupo_atual():
        atual = dialogo(titulo)
        return componente_lista(atual, atributo) if atual else False

    def controle():
        grupo = grupo_atual()
        return unico(grupo, By.CSS_SELECTOR, '.w-listbox__label') if grupo else False

    def codigo_normalizado(valor):
        texto = str(valor or '').strip()
        correspondencia = re.fullmatch(r'(\d+)(?:\.0+)?(?:_\d+)?', texto)
        return str(int(correspondencia.group(1))) if correspondencia else None

    marcar_etapa('abrir a lista de ' + nome_campo)
    clicar(aguardar(controle))
    codigo_pedido = codigo_normalizado(codigo) if codigo is not None else None
    def candidata():
        opcoes = visiveis(navegador, By.CSS_SELECTOR, '[role="option"]')
        if codigo is not None:
            def corresponde(opcao):
                valores = [opcao.get_attribute(a) for a in ('w-code', 'data-value', 'value')]
                return any(codigo_normalizado(v) == codigo_pedido for v in valores if v) or bool(
                    re.match(r'^' + re.escape(str(codigo)) + r'(?:\s*[-:–.]|\s|$)', opcao.text.strip()))
            itens = [e for e in opcoes if corresponde(e)]
        else:
            prefixo = normalizar(LETRA_RESPONSAVEL)
            esperado = normalizar(RESPONSAVEL_ESPERADO)
            itens = [e for e in opcoes if normalizar(e.text) == esperado] if esperado.startswith(prefixo) else [
                e for e in opcoes if normalizar(e.text).startswith(prefixo)]
        if len(itens) > 1:
            raise RuntimeError('Mais de uma opção corresponde ao ' + nome_campo + '.')
        return itens[0] if itens else False

    marcar_etapa('selecionar a opção exata de ' + nome_campo)
    opcao = aguardar(candidata)
    texto = opcao.text.strip()
    codigos_opcao = [opcao.get_attribute(a) for a in ('w-code', 'data-value', 'value')]
    clicar(opcao)
    ultimo_estado = {}
    def selecionada():
        nonlocal ultimo_estado
        grupo = grupo_atual()
        if not grupo:
            return False
        
        
        ultimo_estado = navegador.execute_script("""
            const root = arguments[0], nome = arguments[1];
            const vis = e => !!e.getClientRects().length && getComputedStyle(e).display !== 'none'
                && getComputedStyle(e).visibility !== 'hidden';
            const labels = [...root.querySelectorAll('.w-listbox__label')].filter(vis)
                .map(e => String(e.innerText || e.textContent || '').trim()).filter(Boolean);
            const campos = nome ? [...root.querySelectorAll('input')].filter(e => e.name === nome) : [];
            return {rotulos: labels, valores: campos.map(e => String(e.value || '').trim()),
                opcoes_abertas: [...document.querySelectorAll('[role="option"]')].some(vis)};
        """, grupo, atributo or '')
        rotulos = ultimo_estado.get('rotulos', [])
        rotulo_exato = any(normalizar(r) == normalizar(texto) for r in rotulos)
        if codigo is None:
            return rotulo_exato and not ultimo_estado.get('opcoes_abertas')
        valores = [v for v in ultimo_estado.get('valores', []) if v]
        numericos = [codigo_normalizado(v) for v in valores if codigo_normalizado(v) is not None]
        if any(v != codigo_pedido for v in numericos):
            return False
        codigo_confere = codigo_pedido in numericos or any(v in codigos_opcao for v in valores)
        return (codigo_confere or rotulo_exato) and not ultimo_estado.get('opcoes_abertas')

    marcar_etapa('confirmar o valor selecionado de ' + nome_campo)
    try:
        aguardar(selecionada)
    except TimeoutException as erro:
        diagnostico = ultimo_estado if codigo is not None else {'rotulo_esperado': texto}
        raise RuntimeError('A seleção de ' + nome_campo + ' não foi confirmada. Estado do controle: '
            + json.dumps(diagnostico, ensure_ascii=False)) from erro
    log('Seleção de ' + nome_campo + ' confirmada.')
    return texto


def ler_responsavel(numero):
    estado = capturar_grade_cobranca(numero)
    return estado.get('responsavel') if estado.get('linha') else None


def estado_historico_grade(numero):
    estado = capturar_grade_cobranca(numero)
    if estado.get('linha') and estado.get('historico_texto') is not None and estado.get('historico_tipo') is not None:
        return estado['historico_texto'], estado['historico_tipo']
    return None


def confirmar_historico(numero, texto, tipo_texto, avisos_antes, estado_antes):
    def confirmado():
        verificar_erro_tasy()
        
        for aviso in visiveis(navegador, By.CSS_SELECTOR, '[role="alert"], .toast-success, .notification-success, .ngdialog'):
            mensagem = normalizar(aviso.text)
            if aviso.id not in avisos_antes and 'historico' in mensagem and 'gerad' in mensagem and 'sucesso' in mensagem:
                return True
        
        atual = estado_historico_grade(numero)
        if atual and atual != estado_antes:
            return atual[0] == texto.strip() and atual[1] in (TIPO_HISTORICO, tipo_texto)
        return False
    try:
        aguardar(confirmado, 8)
        return True
    except TimeoutException:
        return False


def main():
    global navegador, aba_principal
    caminho = Path(arquivo_planilha)
    livro = load_workbook(caminho)
    folhas = [f for f in livro.worksheets if {'NR_TITULO', 'HISTORICO'}.issubset(
        {texto_celula(c.value) for c in f[1]})]
    if len(folhas) != 1:
        raise RuntimeError('É necessária uma única aba com NR_TITULO e HISTORICO na primeira linha.')
    folha = folhas[0]
    cabecalhos = [texto_celula(c.value) for c in folha[1]]
    if len([v for v in cabecalhos if v]) != len(set(v for v in cabecalhos if v)):
        raise RuntimeError('Existem cabeçalhos repetidos na planilha.')
    colunas = {texto: i + 1 for i, texto in enumerate(cabecalhos) if texto}
    for nome in ('STATUS', 'ERRO', 'STATUS_RESPONSAVEL', 'STATUS_HISTORICO', 'RESPONSAVEL_ALVO',
                 'TIPO_HISTORICO_ALVO', 'HASH_HISTORICO', 'ATUALIZADO_EM'):
        if nome not in colunas:
            colunas[nome] = folha.max_column + 1
            folha.cell(1, colunas[nome], nome)
    marca = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    pasta_backup = caminho.parent / 'Bak'
    pasta_backup.mkdir(exist_ok=True)
    backup = pasta_backup / (caminho.name + '.bak')
    shutil.copy2(caminho, backup)

    def ler(linha, coluna):
        return texto_celula(folha.cell(linha, colunas[coluna]).value)

    def salvar(linha=None, **valores):
        if linha is not None:
            valores['ATUALIZADO_EM'] = datetime.now().isoformat(timespec='seconds')
            for chave, valor in valores.items():
                folha.cell(linha, colunas[chave], valor)
        temporario = None
        try:
            with tempfile.NamedTemporaryFile(dir=caminho.parent, suffix='.xlsx', delete=False) as arquivo:
                temporario = Path(arquivo.name)
            livro.save(temporario)
            
            os.replace(temporario, caminho)
        finally:
            if temporario is not None and temporario.exists():
                temporario.unlink()

    salvar()
    options = webdriver.EdgeOptions()
    options.add_argument('--start-maximized')
    navegador = webdriver.Edge(service=Service(caminho_driver), options=options)
    aba_principal = navegador.current_window_handle
    navegador.get(config.url_sistema())
    preencher(aguardar(lambda: unico(navegador, By.ID, 'loginUsername')), USUARIO)
    campo_senha = aguardar(lambda: unico(navegador, By.ID, 'loginPassword'))
    preencher(campo_senha, SENHA)
    campo_senha.send_keys(Keys.ENTER)
    aguardar(lambda: not visiveis(navegador, By.ID, 'loginPassword'), 60)
    abrir_funcao()
    resultados = []
    import csv
    relatorio = caminho.with_name('Alt responsavel - resultado.csv')
    with relatorio.open('w', encoding='utf-8-sig', newline='') as arquivo:
        csv.writer(arquivo, delimiter=';').writerow(['Título', 'Resultado', 'Motivo', 'Etapa responsável', 'Etapa histórico'])

    def registrar_resultado(registro):
        with relatorio.open('a', encoding='utf-8-sig', newline='') as arquivo:
            csv.writer(arquivo, delimiter=';').writerow(registro + (
                ler(linha, 'STATUS_RESPONSAVEL'), ler(linha, 'STATUS_HISTORICO')))
            arquivo.flush()
            os.fsync(arquivo.fileno())
        resultados.append(registro)

    titulos_execucao = set()
    for linha in range(2, folha.max_row + 1):
        numero = re.sub(r'\.0$', '', ler(linha, 'NR_TITULO'))
        texto = ler(linha, 'HISTORICO')
        if not numero:
            continue
        estado = ler(linha, 'STATUS')
        if estado == 'OK':
            registrar_resultado((numero, 'JÁ CONCLUÍDO', 'Preservado conforme status da planilha.'))
            continue
        if numero in titulos_execucao:
            salvar(linha, STATUS='REVISAR_DUPLICADO', ERRO='Título repetido nesta execução; não repeti alterações/históricos.')
            registrar_resultado((numero, 'REVISAR_DUPLICADO', 'Título repetido na planilha.'))
            continue
        titulos_execucao.add(numero)
        
        if (ler(linha, 'STATUS_RESPONSAVEL').startswith('ENVIADO')
                or ler(linha, 'STATUS_HISTORICO').startswith('ENVIADO')):
            registrar_resultado((numero, 'REVISAR', 'Gravação anterior sem conferência. Não foi reenviada.'))
            continue
        try:
            if not numero.isdigit() or not texto or len(texto) > 4000:
                raise ValueError('Título deve ser numérico e HISTORICO deve conter entre 1 e 4.000 caracteres.')
            resumo_texto = hashlib.sha256(texto.encode('utf-8')).hexdigest()
            if ler(linha, 'HASH_HISTORICO') and ler(linha, 'HASH_HISTORICO') != resumo_texto:
                raise RuntimeError('O texto mudou após uma etapa anterior; confira o registro antes de repetir.')
            log(f'Processando título {numero}, linha {linha}.')
            aguardar_foco()
            marcar_etapa(f'pesquisar o título {numero}')
            pesquisar_titulo(numero)
            marcar_etapa('ler o status do título pesquisado')
            situacao = status_titulo(numero)
            log('Status confirmado: ' + situacao)
            if situacao != 'Pendente':
                motivo = 'Status encontrado: ' + situacao
                salvar(linha, STATUS='IGNORADO_STATUS', ERRO=motivo)
                registrar_resultado((numero, 'IGNORADO_STATUS', motivo))
                continue
            if ler(linha, 'STATUS_RESPONSAVEL') != 'CONFIRMADO':
                
                marcar_etapa('localizar a coluna do responsável atual')
                anterior = ler_responsavel(numero)
                if anterior is None:
                    raise RuntimeError('Não identifiquei a coluna Responsável na grade. Falta seu HTML para conferir a gravação; não alterei este título.')
                menu_titulo(numero, 'Responsável cobrança', alterar=True)
                marcar_etapa('aguardar a janela Alterar responsável cobrança')
                janela = aguardar(lambda: dialogo('Alterar responsável cobrança'))
                alvo = escolher_opcao(janela)
                salvar(linha, STATUS='EM_ANDAMENTO', STATUS_RESPONSAVEL='ENVIADO_NAO_CONFERIDO',
                       RESPONSAVEL_ALVO=alvo, TIPO_HISTORICO_ALVO=TIPO_HISTORICO, HASH_HISTORICO=resumo_texto, ERRO='')
                marcar_etapa('enviar a alteração de responsável uma única vez')
                clicar(aguardar(lambda: botao(janela, 'Alterar')))
                aguardar(lambda: not dialogo('Alterar responsável cobrança'))
                verificar_erro_tasy()
                marcar_etapa('confirmar o responsável gravado na grade')
                aguardar(lambda: normalizar(ler_responsavel(numero)) == normalizar(alvo))
                salvar(linha, STATUS_RESPONSAVEL='CONFIRMADO')
            else:
                alvo = ler(linha, 'RESPONSAVEL_ALVO')
                if (not alvo or not normalizar(alvo).startswith(normalizar(LETRA_RESPONSAVEL))
                        or normalizar(ler_responsavel(numero)) != normalizar(alvo)):
                    raise RuntimeError('O responsável gravado anteriormente não pôde ser reconfirmado.')
            if ler(linha, 'STATUS_HISTORICO') != 'CONFIRMADO':
                menu_titulo(numero, 'Gerar históricos')
                marcar_etapa('aguardar a janela Gerar históricos')
                janela = aguardar(lambda: dialogo('Gerar históricos'))
                tipo_texto = escolher_opcao(janela, 'TIPO_HISTORICO', TIPO_HISTORICO)
                marcar_etapa('preencher o histórico complementar da planilha')
                campo_texto = aguardar(lambda: campo_dialogo('Gerar históricos', 'textarea[name="DS_HISTORICO"]'))
                preencher(campo_texto, texto)
                log(f'Histórico complementar preenchido e conferido: {len(texto)} caracteres.')
                avisos_antes = {e.id for e in visiveis(navegador, By.CSS_SELECTOR,
                    '[role="alert"], .toast-success, .notification-success, .ngdialog')}
                estado_antes = estado_historico_grade(numero)
                marcar_etapa('aguardar Gerar habilitado após o preenchimento')
                botao_gerar = aguardar(lambda: botao_dialogo_habilitado('Gerar históricos', 'Gerar'))
                salvar(linha, STATUS_HISTORICO='ENVIADO_NAO_CONFERIDO')
                marcar_etapa('enviar o histórico uma única vez')
                clicar(botao_gerar)
                marcar_etapa('confirmar a gravação do histórico no Tasy')
                if not confirmar_historico(numero, texto, tipo_texto, avisos_antes, estado_antes):
                    raise RuntimeError('Histórico enviado, mas não confirmado. Não será reenviado automaticamente. Falta conferir o retorno do Tasy.')
                salvar(linha, STATUS_HISTORICO='CONFIRMADO')
            salvar(linha, STATUS='OK', ERRO='')
            registrar_resultado((numero, 'OK', 'Responsável e histórico confirmados.'))
        except PermissionError:
            log('Acesso negado ou planilha bloqueada. Execução interrompida sem novas tentativas de gravação.')
            raise
        except Exception as erro:
            motivo = f'{type(erro).__name__}: {erro} | Etapa: {getattr(navegador, "_etapa_cobranca", "não identificada")}'
            registrar_erro_arquivo(numero, motivo)
            salvar(linha, STATUS='REVISAR' if ler(linha, 'STATUS_RESPONSAVEL').startswith('ENVIADO')
                   or ler(linha, 'STATUS_HISTORICO').startswith('ENVIADO') else 'ERRO', ERRO=motivo)
            registrar_resultado((numero, 'FALHA', motivo))
            log(f'Título {numero}: {motivo}')
            try:
                fechar_erro_conhecido()
                for titulo in ('Alterar responsável cobrança', 'Gerar históricos'):
                    janela = dialogo(titulo)
                    if janela:
                        fechar_cancelando(janela)
                if dialogos_bloqueadores():
                    raise RuntimeError('Uma janela desconhecida impede a próxima pesquisa.')
                
                if getattr(navegador, '_menu_cobranca_aberto', False):
                    aguardar_foco()
                    navegador.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                    aguardar(lambda: not visiveis(navegador, By.CSS_SELECTOR,
                        '.wpopupmenu [role="menuitem"], [role="menu"] [role="menuitem"]'), 3)
                    navegador._menu_cobranca_aberto = False
            except PermissionError:
                raise
            except Exception as erro_recuperacao:
                detalhe = f'Falha na recuperação após o título {numero}: {type(erro_recuperacao).__name__}: {erro_recuperacao}'
                registrar_erro_arquivo(numero, detalhe)
                salvar(linha, ERRO=motivo + ' | ' + detalhe)
                registrar_resultado((numero, 'RECUPERACAO_INTERROMPIDA', detalhe))
                log(detalhe)
                log('Execução encerrada com os resultados e as etapas preservados na planilha e no relatório.')
                break
            log('Ocorrência registrada. Seguindo para o próximo título.')
    log(f'Relatório: {relatorio}')
    log(f'Backup único da planilha: {backup}')
    for numero, resultado, motivo in resultados:
        log(f'{numero}: {resultado} — {motivo}')


def manter_tela_ativa():
    """Solicita tela e sistema ativos enquanto esta thread estiver em execução."""
    import ctypes
    if os.name != 'nt':
        raise RuntimeError('A manutenção da tela ativa desta versão requer Windows.')
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    funcao = kernel.SetThreadExecutionState
    funcao.argtypes = [ctypes.c_uint32]
    funcao.restype = ctypes.c_uint32
    es_continuous = 0x80000000
    es_system_required = 0x00000001
    es_display_required = 0x00000002
    anterior = funcao(es_continuous | es_system_required | es_display_required)
    if not anterior:
        raise RuntimeError('O Windows não confirmou a solicitação de tela e sistema ativos. A rotina não foi iniciada.')
    log('Tela e sistema mantidos ativos durante a execução.')
    return kernel, funcao, anterior


def restaurar_estado_energia(estado):
    """Libera a solicitação temporária e preserva o estado anterior desta thread."""
    if estado is None:
        return
    kernel, funcao, anterior = estado
    
    restaurar = 0x80000000 | (anterior & 0x00000043)
    if funcao(restaurar):
        log('Solicitação de tela ativa liberada; comportamento normal de energia restaurado.')
    else:
        log('O Windows não confirmou a liberação; a solicitação também termina ao encerrar esta thread.')


if __name__ == '__main__':
    estado_energia = None
    try:
        estado_energia = manter_tela_ativa()
        main()
    except KeyboardInterrupt:
        log('Execução interrompida pelo usuário. Consulte as etapas registradas na planilha antes de retomar.')
    except Exception as erro:
        log(f'Execução encerrada: {type(erro).__name__}: {erro}')
    finally:
        try:
            if navegador is not None:
                try:
                    navegador.quit()
                except KeyboardInterrupt:
                    log('Fechamento interrompido pelo usuário; os registros já gravados foram preservados.')
                except Exception as erro:
                    log(f'Não foi possível fechar a sessão do navegador: {erro}')
        finally:
            restaurar_estado_energia(estado_energia)
