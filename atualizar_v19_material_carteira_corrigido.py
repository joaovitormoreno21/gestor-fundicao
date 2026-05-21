# =============================================================================
#  GESTOR FUNDIÇÃO — Atualização Direta via SQL Server
#
#  Como usar: clique duas vezes neste arquivo
#  O que faz:
#    1. Conecta direto ao SQL Server (sem abrir o Excel)
#    2. Executa as queries (Acompanhamento + Gerenciamento + Vfusão)
#    3. Salva o BaseDados.xlsx atualizado, preservando a aba Apoio
#    4. Gera o dashboard HTML em index.html
#
#  Dependências instaladas automaticamente na primeira execução:
#    pyodbc, pandas, openpyxl
# =============================================================================

import subprocess, sys, os, traceback
from pathlib import Path
from datetime import date, datetime

# ── Utilidades de tela/log ────────────────────────────────────────────────────
def log(msg=''):
    """Mostra mensagens imediatamente no console, inclusive em duplo clique."""
    print(msg, flush=True)

def pausar(msg='\nPressione Enter para fechar...'):
    try:
        input(msg)
    except Exception:
        pass

def perguntar(msg, padrao='n'):
    """Pergunta com valor padrão para evitar travar por dúvida do usuário."""
    resp = input(msg).strip().lower()
    return resp if resp else padrao

# ── Instala dependências automaticamente na primeira execução ──────────────────────────────────────
def instalar(pkg):
    log(f'  Instalando {pkg}. Isso pode levar alguns minutos na primeira execução...')
    try:
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install', pkg,
            '--disable-pip-version-check'
        ])
        log(f'  ✓ {pkg} instalado')
    except subprocess.CalledProcessError:
        # Tenta com --break-system-packages (necessário em alguns ambientes Linux)
        try:
            subprocess.check_call([
                sys.executable, '-m', 'pip', 'install', pkg,
                '--disable-pip-version-check', '--break-system-packages'
            ])
            log(f'  ✓ {pkg} instalado')
        except Exception as e:
            log(f'  ⚠ Erro ao instalar {pkg}: {e}')
            log(f'  Por favor, instale manualmente: pip install {pkg}')
            pausar('\nPressione Enter para sair...')
            sys.exit(1)

log('Verificando dependências...')
for pkg in ['pyodbc', 'pandas', 'openpyxl']:
    try:
        __import__(pkg)
        log(f'  ✓ {pkg} OK')
    except ImportError:
        instalar(pkg)

import pyodbc
import pandas as pd
import json

VERSAO_SCRIPT = '2026-05-19-v19-MATERIAL-CARTEIRA-CORRIGIDO'

# ── Caminhos ──────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
EXCEL_PATH = BASE_DIR / 'dados' / 'BaseDados.xlsx'
OUTPUT_DIR = BASE_DIR / 'output'
HTML_PATH  = BASE_DIR / 'index.html'
OUTPUT_DIR.mkdir(exist_ok=True)
(BASE_DIR / 'dados').mkdir(exist_ok=True)

# ── Configurações do banco ────────────────────────────────────────────────────
SQL_SERVER = 'srv-moreno-prd'
SQL_DB     = 'dbwel'
SQL_USER   = 'consulta.prd'
SQL_PASS   = 'consulta'

# ── Queries (exatamente as mesmas do Excel) ───────────────────────────────────
QUERY_ACOMP = """
SELECT
    v.empresa,
    CAST(v.id_ordem_servico AS VARCHAR(50)) + '-' + CAST(v.item AS VARCHAR(50)) AS chave,
    v.id_ordem_servico, v.item,
    v.codigo_peca, v.seq_peca, v.qtde_planejada,
    v.cliente, v.prz_contratual,
    v.projeto_prev,            v.projeto_real,
    v.fabr_modelo_prev,        v.fabr_modelo_real,
    v.insp_modelo_prev,        v.insp_modelo_real,
    v.ficha_processo_prev,     v.ficha_processo_real,
    v.modelacao_prev,          v.modelacao_real,
    v.lib_processo_eng_prev,   v.lib_processo_eng_real,
    v.moldagem_prev,           v.moldagem_real,
    v.fusao_prev,              v.fusao_real,
    v.desmoldagem_prev,        v.desmoldagem_real,
    v.normalizacao_corte_prev, v.normalizacao_corte_real,
    v.normalizacao_prev,       v.normalizacao_real,
    v.tempera_prev,            v.tempera_real,
    v.revenimento_prev,        v.revenimento_real,
    v.alivio_tensoes_prev,     v.alivio_tensoes_real,
    v.acabamento_prev,         v.acabamento_real,
    v.usinagem_prev,           v.usinagem_real,
    v.lib_fundido_prev,        v.lib_fundido_real,
    v.lt_total_prev,           v.lt_total_real,  v.lt_total_dif,
    v.descr_peca,              v.desenho
FROM dbwel.dbo.v_fund_acompanhamento v
WHERE v.lib_fundido_real IS NULL
"""

QUERY_GER = """
SELECT
    g.descr_resumida,
    g.ovenda_sap,
    g.itemov_sap,
    g.oprod_sap,
    g.mega_filial_doc_v,
    g.mega_serie_doc_v,
    g.mega_doc_v,
    g.mega_item_doc_v,
    CAST(g.id_ordem_servico AS VARCHAR(50)) + '-' + CAST(g.item AS VARCHAR(50)) AS chave,
    g.id_ordem_servico,
    g.item,
    g.qtde,
    g.dt_abertura,
    g.dt_envio_email,
    g.nro_ped_compra,
    g.cliente,
    g.nome_fantasia,
    g.codigo_peca_planej,
    g.seq_peca_planej,
    g.numero_corrida,
    g.descricao,
    g.id_cond_fornec,
    g.desenho_mod_com,
    g.desenho_desbaste,
    g.desenho_mecanica,
    g.material_comercial,
    g.material_fusao,
    g.nr_roteiro,
    g.pie,
    g.kic_data,
    g.data_entrada,
    g.ultima_inspecao_mod,
    g.status_modelo,
    g.tratamentos_term,
    g.prz_entrega,
    g.prz_fundicao,
    g.prz_contratual,
    g.dt_fatur,
    g.nro_nf,
    g.semana_fusao,
    g.preteste,
    g.data_pre_teste,
    g.tem_certificado,
    g.comp_qumica,
    g.propr_mecanica,
    g.teste_impacto,
    g.dureza,
    g.metalografia,
    g.status_notificacao,
    g.tem_rnc,
    g.numero_controle,
    g.reservada,
    g.data_planejamento,
    g.data_programacao,
    g.data_fusao,
    g.data_exp_fund,
    g.peso_liquido,
    g.peso_bruto,
    g.dias_abertura_hoje,
    g.dias_abertura_prz_contratual,
    g.dias_abertura_prz_fundicao,
    g.dias_abertura_insp_mod,
    g.dias_abertura_fusao,
    g.dias_abertura_t_termico,
    g.dias_fusao_t_termico,
    g.dias_abertura_faturamento,
    g.dias_t_termico_prz_fund,
    g.dias_t_termico_prz_contrato,
    g.dias_abertura_liberacao
FROM dbwel.dbo.v_gerencia_os_gestare g
WHERE g.semana_fusao >= 202515
"""

QUERY_FUSAO = """
SELECT
    f.empresa,
    f.id_ordem_servico,
    f.item,
    f.codigo_peca,
    f.seq_peca,
    f.descr_peca,
    f.grupo,
    f.descr_grupo,
    f.material,
    f.material_comercial,
    f.ano_semana,
    f.carga_prog_forno,
    f.temp_vazada,
    f.qtde_planejada,
    f.qtde_programada,
    f.qtde_fundida,
    f.peso_liquido,
    f.peso_bruto,
    f.peca_perdida,
    f.nome_rg,
    f.observacao,
    f.id_area_neg,
    f.descr_area_neg
FROM dbwel.dbo.v_fusao f
WHERE f.ano_semana >= '202401'
"""

# =============================================================================
#  PARTE 1 — Conectar e buscar dados
# =============================================================================

def conectar():
    drivers = [d for d in pyodbc.drivers() if 'SQL Server' in d]
    if not drivers:
        log('\n[ERRO] Nenhum driver ODBC do SQL Server encontrado.')
        log('  Baixe e instale em: https://aka.ms/downloadmsodbcsql')
        pausar()
        sys.exit(1)
    drivers.sort(reverse=True)
    driver = drivers[0]
    log(f'  Driver ODBC: {driver}')
    conn_str = (
        f'DRIVER={{{driver}}};'
        f'SERVER={SQL_SERVER};'
        f'DATABASE={SQL_DB};'
        f'UID={SQL_USER};'
        f'PWD={SQL_PASS};'
        f'TrustServerCertificate=yes;'
        f'Connection Timeout=30;'
    )
    try:
        log('  Tentando conectar ao banco...')
        conn = pyodbc.connect(conn_str, timeout=30)
        return conn
    except pyodbc.Error as e:
        log(f'\n[ERRO] Falha na conexão:\n  {e}')
        log('\n  Verifique:')
        log('  • Você está na rede da empresa (ou VPN ativa)')
        log(f'  • Servidor: {SQL_SERVER}')
        log(f'  • Usuário:  {SQL_USER}')
        pausar()
        sys.exit(1)

def buscar_dados():
    log('\n[1/4] Conectando ao SQL Server...')
    conn = conectar()
    log('  ✓ Conectado com sucesso!')

    log('[2/5] Buscando Carteira/Acompanhamento (v_fund_acompanhamento)...')
    acomp = pd.read_sql(QUERY_ACOMP, conn)
    log(f'  ✓ {len(acomp)} registros em produção/carteira')

    log('[3/5] Buscando Histórico/Gerenciamento (v_gerencia_os_gestare)...')
    ger = pd.read_sql(QUERY_GER, conn)
    log(f'  ✓ {len(ger)} registros históricos')

    log('[4/5] Buscando Fusão (v_fusao)...')
    fusao = pd.read_sql(QUERY_FUSAO, conn)
    log(f'  ✓ {len(fusao)} registros de fusão')

    conn.close()
    return acomp, ger, fusao

# =============================================================================
#  PARTE 2 — Processar dados
# =============================================================================

# ─────────────────────────────────────────────────────────────────────────────
#  FUNÇÕES AUXILIARES PARA TRATAMENTOS TÉRMICOS
# ─────────────────────────────────────────────────────────────────────────────

def todos_tt_concluidos_ou_nao_aplicaveis(row):
    """
    Verifica se TODOS os TTs que foram planejados (Prev.) já foram executados (Real).
    OU se não eram aplicáveis (Prev. vazio = não precisa fazer).
    
    Retorna True se:
    - Todos os TTs planejados foram executados OU
    - Não havia TTs planejados
    """
    tts = [
        ('alivio_tensoes_prev', 'alivio_tensoes_real'),
        ('revenimento_prev', 'revenimento_real'),
        ('tempera_prev', 'tempera_real'),
        ('normalizacao_prev', 'normalizacao_real'),
        ('normalizacao_corte_prev', 'normalizacao_corte_real')
    ]
    
    for prev, real in tts:
        # Se FOI planejado (Prev. preenchido)...
        if pd.notna(row.get(prev)):
            # ... mas NÃO foi executado (Real vazio) → FALSO
            if pd.isna(row.get(real)):
                return False
    
    return True  # Todos os TTs planejados foram feitos OU não havia TTs

def algum_tt_real_preenchido(row):
    """Verifica se algum TT já foi executado (Real preenchido)."""
    tts_real = [
        'normalizacao_corte_real', 'normalizacao_real',
        'tempera_real', 'revenimento_real', 'alivio_tensoes_real'
    ]
    return any(pd.notna(row.get(tt)) for tt in tts_real)

def desmoldagem_feita_e_tt_planejado(row):
    """Desmoldagem feita E existe algum TT planejado (Prev. preenchido)."""
    if pd.isna(row.get('desmoldagem_real')):
        return False
    
    tts_prev = [
        'normalizacao_corte_prev', 'normalizacao_prev',
        'tempera_prev', 'revenimento_prev', 'alivio_tensoes_prev'
    ]
    return any(pd.notna(row.get(tt)) for tt in tts_prev)

def desmoldagem_feita_e_nenhum_tt_planejado(row):
    """Desmoldagem feita E NENHUM TT foi planejado (todos Prev. vazios)."""
    if pd.isna(row.get('desmoldagem_real')):
        return False
    
    tts_prev = [
        'normalizacao_corte_prev', 'normalizacao_prev',
        'tempera_prev', 'revenimento_prev', 'alivio_tensoes_prev'
    ]
    return all(pd.isna(row.get(tt)) for tt in tts_prev)

def obter_menor_data_tt_pendente(row):
    """
    Retorna a MENOR data entre os TTs que:
    - Têm data prevista (Prev. preenchido) E
    - NÃO têm data real (Real vazio) = estão pendentes
    
    Retorna pd.Timestamp('2099-12-31') se não há TTs pendentes.
    """
    INFINITO = pd.Timestamp('2099-12-31')
    
    tts = [
        ('normalizacao_corte_prev', 'normalizacao_corte_real'),
        ('normalizacao_prev', 'normalizacao_real'),
        ('tempera_prev', 'tempera_real'),
        ('revenimento_prev', 'revenimento_real'),
        ('alivio_tensoes_prev', 'alivio_tensoes_real')
    ]
    
    datas_pendentes = []
    
    for prev, real in tts:
        # Se tem Prev. mas não tem Real → está pendente
        if pd.notna(row.get(prev)) and pd.isna(row.get(real)):
            datas_pendentes.append(row.get(prev))
    
    if not datas_pendentes:
        return INFINITO  # Nenhum TT pendente
    
    return min(datas_pendentes)

# ─────────────────────────────────────────────────────────────────────────────
#  FUNÇÃO PRINCIPAL: CALCULAR ETAPA ATUAL
# ─────────────────────────────────────────────────────────────────────────────

def calcular_etapa_atual(row):
    """
    Determina qual a PRÓXIMA etapa que precisa ser executada.
    Verifica as etapas DE TRÁS PARA FRENTE (do fim para o início).
    
    Lógica baseada na fórmula Excel do arquivo Acompanhamento_Fundição.
    
    Returns:
        str: Nome da próxima etapa (ex: 'Fusão', 'Tratamento Térmico', 'Finalizado')
    """
    
    # 1. Se Liberação Fundido foi feita → FINALIZADO
    if pd.notna(row.get('lib_fundido_real')):
        return 'Finalizado'
    
    # 2. Se Usinagem foi feita → próxima é Liberação Fundido
    if pd.notna(row.get('usinagem_real')):
        return 'Liberação Fundido'
    
    # 3. Se Acabamento foi feito → próxima é Usinagem
    if pd.notna(row.get('acabamento_real')):
        return 'Usinagem'
    
    # 4. LÓGICA COMPLEXA: verificar se vai para Acabamento ou Tratamento Térmico
    # Se Desmoldagem foi feita E todos TTs planejados foram concluídos → Acabamento
    if pd.notna(row.get('desmoldagem_real')) and todos_tt_concluidos_ou_nao_aplicaveis(row):
        return 'Acabamento'
    
    # 5. Se algum TT foi executado → está em Tratamento Térmico
    if algum_tt_real_preenchido(row):
        return 'Tratamento Térmico'
    
    # 6. Se Desmoldagem feita E TT planejado (mas não executado) → TT
    if desmoldagem_feita_e_tt_planejado(row):
        return 'Tratamento Térmico'
    
    # 7. Se Desmoldagem feita mas NENHUM TT planejado → pula para Acabamento
    if desmoldagem_feita_e_nenhum_tt_planejado(row):
        return 'Acabamento'
    
    # 8. Se Fusão foi feita → próxima é Desmoldagem
    if pd.notna(row.get('fusao_real')):
        return 'Desmoldagem'
    
    # 9. Se Moldagem foi feita → próxima é Fusão
    if pd.notna(row.get('moldagem_real')):
        return 'Fusão'
    
    # 10. Se Liberação Processo Eng foi feita → próxima é Moldagem
    if pd.notna(row.get('lib_processo_eng_real')):
        return 'Moldagem'
    
    # 11. Se Modelação Processo foi feita → próxima é Liberação Processo Eng
    if pd.notna(row.get('modelacao_real')):
        return 'Liberação Processo Eng'
    
    # 12. Se Ficha Processo foi feita → próxima é Modelação Processo
    if pd.notna(row.get('ficha_processo_real')):
        return 'Modelação Processo'
    
    # 13. Se Inspeção Modelo foi feita → próxima é Ficha Processo
    if pd.notna(row.get('insp_modelo_real')):
        return 'Ficha Processo'
    
    # 14. Se Fabricação Modelo foi feita → próxima é Inspeção Modelo
    if pd.notna(row.get('fabr_modelo_real')):
        return 'Inspeção Modelo'
    
    # 15. Se Projeto Eng foi feito → próxima é Fabricação Modelo
    if pd.notna(row.get('projeto_real')):
        return 'Fabricação Modelo'
    
    # 16. Se nada foi feito mas tem Projeto Eng planejado → próxima é Projeto Eng
    if pd.notna(row.get('projeto_prev')):
        return 'Projeto Eng'
    
    # 17. Caso contrário, planejamento não foi preenchido
    return 'Planejamento Não Preenchido'

# ─────────────────────────────────────────────────────────────────────────────
#  FUNÇÃO: CALCULAR STATUS
# ─────────────────────────────────────────────────────────────────────────────

def obter_data_prevista_da_etapa(row, etapa):
    """
    Retorna a data prevista correspondente à etapa atual.
    Para Tratamento Térmico, retorna a MENOR data entre os TTs pendentes.
    """
    
    mapeamento = {
        'Projeto Eng': 'projeto_prev',
        'Fabricação Modelo': 'fabr_modelo_prev',
        'Inspeção Modelo': 'insp_modelo_prev',
        'Ficha Processo': 'ficha_processo_prev',
        'Modelação Processo': 'modelacao_prev',
        'Liberação Processo Eng': 'lib_processo_eng_prev',
        'Moldagem': 'moldagem_prev',
        'Fusão': 'fusao_prev',
        'Desmoldagem': 'desmoldagem_prev',
        'Acabamento': 'acabamento_prev',
        'Usinagem': 'usinagem_prev',
        'Liberação Fundido': 'lib_fundido_prev',
    }
    
    # Caso especial: Tratamento Térmico
    if etapa == 'Tratamento Térmico':
        return obter_menor_data_tt_pendente(row)
    
    # Etapas normais
    campo = mapeamento.get(etapa)
    if campo:
        val = row.get(campo)
        if pd.notna(val):
            return val
    
    return None

def obter_prazo_referencia(row):
    """
    Define qual prazo usar como referência:
    - Se Usinagem|Previsto preenchido → Prazo Contratual
    - Senão → Prazo Fundição
    - Se Prazo Fundição vazio → usa Prazo Contratual
    """
    if pd.notna(row.get('usinagem_prev')):
        return row.get('prz_contratual')
    else:
        prz_fund = row.get('prz_fundicao')
        prz_cont = row.get('prz_contratual')
        # Se não tem fundição, usa contratual
        if pd.isna(prz_fund) and pd.notna(prz_cont):
            return prz_cont
        return prz_fund

def calcular_status(row, etapa_atual, hoje):
    """
    Calcula o status baseado na data prevista da etapa atual e prazos de referência.
    
    Lógica baseada na fórmula LET do Excel.
    
    Args:
        row: linha do DataFrame
        etapa_atual: string retornada por calcular_etapa_atual()
        hoje: pd.Timestamp de hoje
    
    Returns:
        str: 'Finalizado' | 'No Prazo' | 'Potencial Atraso' | 'Atrasado'
    """
    
    INFINITO = pd.Timestamp('2099-12-31')
    
    # 1. Se finalizado → status = Finalizado
    if etapa_atual == 'Finalizado':
        return 'Finalizado'
    
    # 2. Busca a data prevista da etapa atual
    data_prevista = obter_data_prevista_da_etapa(row, etapa_atual)
    
    # 3. Define o prazo de referência
    prazo_referencia = obter_prazo_referencia(row)
    
    # 4. Se não tem data prevista OU data prevista = infinito → No Prazo
    if data_prevista is None or data_prevista == INFINITO:
        return 'No Prazo'
    
    # 5. Se data prevista ainda não passou → No Prazo
    if data_prevista >= hoje:
        return 'No Prazo'
    
    # 6. Data prevista JÁ PASSOU (< hoje)
    # Verifica o prazo de referência:
    if pd.notna(prazo_referencia):
        if prazo_referencia < hoje:
            return 'Atrasado'        # Prazo de referência também passou
        else:
            return 'Potencial Atraso' # Prazo de referência ainda não passou
    else:
        # Se não tem prazo de referência, considera apenas a data prevista
        return 'Atrasado'

# ─────────────────────────────────────────────────────────────────────────────
#  MANTER FUNÇÃO ANTIGA PARA COMPATIBILIDADE (não usada mais)
# ─────────────────────────────────────────────────────────────────────────────

def etapa_atual(row):
    """DEPRECATED: Usar calcular_etapa_atual() com lógica completa do Excel."""
    return calcular_etapa_atual(row)

def classifica_grupo(mat):
    if not mat or str(mat) in ('nan','None',''): return 'Outros'
    m = str(mat).upper()
    if 'FM-' in m:
        return 'Ferro Nodular' if 'N' in m else 'Ferro Cinzento'
    if 'INOX' in m or 'CA6' in m or 'ITF-15' in m: return 'Aço Inox'
    if 'XT' in m:   return 'Aço Manganês'
    if '50 MN' in m or 'MN N' in m: return 'Aço Manganês'
    if any(x in m for x in ['WC','4330','4335','4140','8640','WC6','WC9']): return 'Aço Ligado'
    return 'Aço Carbono'

def fmt_date(v):
    try:
        return pd.Timestamp(v).strftime('%d/%m/%Y') if pd.notna(v) else '-'
    except Exception:
        return '-'

def short(s, n=40):
    if not s or str(s) in ('nan','None',''): return ''
    s = str(s).strip()
    return s[:n] + '…' if len(s) > n else s

def normaliza_grupo(v):
    """Regra do Excel: grupos 6 e 9 aparecem como 06/09."""
    if pd.isna(v):
        return ''
    txt = str(v).strip()
    try:
        num = int(float(txt.replace(',', '.')))
        if num in (6, 9):
            return '06/09'
        return str(num)
    except Exception:
        return txt

def carregar_apoio():
    """
    Lê a aba Apoio do BaseDados.xlsx, se existir.
    Usa: coluna A = Grupo, coluna B = Desc_Grupo, coluna E = fator para metal_liq_calc.
    """
    fator_map = {}
    desc_map = {}
    if not EXCEL_PATH.exists():
        log('  ⚠ BaseDados.xlsx não encontrado para ler a aba Apoio. Usando fallback.')
        return fator_map, desc_map
    try:
        apoio = pd.read_excel(EXCEL_PATH, sheet_name='Apoio')
        if apoio.shape[1] >= 2:
            col_grupo = apoio.columns[0]
            col_desc = apoio.columns[1]
            for _, r in apoio.iterrows():
                g = normaliza_grupo(r.get(col_grupo))
                if g:
                    desc_map[g] = str(r.get(col_desc, '')).strip() if pd.notna(r.get(col_desc)) else ''
        if apoio.shape[1] >= 5:
            col_grupo = apoio.columns[0]
            col_fator = apoio.columns[4]
            for _, r in apoio.iterrows():
                g = normaliza_grupo(r.get(col_grupo))
                try:
                    fator_map[g] = float(r.get(col_fator))
                except Exception:
                    pass
        log(f'  ✓ Apoio carregado: {len(desc_map)} grupos, {len(fator_map)} fatores')
    except Exception as e:
        log(f'  ⚠ Não consegui ler a aba Apoio: {e}')
    return fator_map, desc_map

def numero(v, padrao=0):
    try:
        if pd.isna(v):
            return padrao
        return float(v)
    except Exception:
        return padrao

def processar(acomp, ger, fusao):
    """
    REGRAS ATUAIS:
    - dados_Gerenciamento = histórico de produção e base dos CARDS/GRÁFICOS.
    - dados_Acomp = carteira/em produção e base da TABELA.
    - metal_liq_calc, Grupo e Desc_Grupo são criados em dados_Gerenciamento.
    """
    fator_map, desc_map = carregar_apoio()

    # -------------------- Gerenciamento / histórico --------------------
    ger_proc = ger.copy()
    for col in ['prz_contratual', 'prz_fundicao', 'data_fusao', 'dt_abertura']:
        if col in ger_proc.columns:
            ger_proc[col] = pd.to_datetime(ger_proc[col], errors='coerce')

    # Lookup da v_fusao por codigo_peca_planej -> codigo_peca
    fusao_ref = fusao.copy()
    if 'ano_semana' in fusao_ref.columns:
        fusao_ref['ano_semana_num'] = pd.to_numeric(fusao_ref['ano_semana'], errors='coerce')
        fusao_ref = fusao_ref.sort_values('ano_semana_num').drop_duplicates('codigo_peca', keep='last')
    else:
        fusao_ref = fusao_ref.drop_duplicates('codigo_peca', keep='last')

    cols_fusao = [c for c in ['codigo_peca', 'grupo', 'descr_grupo'] if c in fusao_ref.columns]
    if 'codigo_peca_planej' in ger_proc.columns and cols_fusao:
        ger_proc = ger_proc.merge(
            fusao_ref[cols_fusao].rename(columns={'grupo': 'Grupo_lookup', 'descr_grupo': 'Desc_Grupo_lookup'}),
            left_on='codigo_peca_planej',
            right_on='codigo_peca',
            how='left'
        )
    else:
        ger_proc['Grupo_lookup'] = ''
        ger_proc['Desc_Grupo_lookup'] = ''

    ger_proc['Grupo'] = ger_proc['Grupo_lookup'].apply(normaliza_grupo)
    ger_proc['Desc_Grupo'] = ger_proc.apply(
        lambda r: desc_map.get(r.get('Grupo'), '')
                  or (str(r.get('Desc_Grupo_lookup', '')).strip() if pd.notna(r.get('Desc_Grupo_lookup')) else '')
                  or r.get('Grupo')
                  or 'Sem grupo',
        axis=1
    )

    def calc_metal_ger(r):
        pb = numero(r.get('peso_bruto'), 0)
        if pb > 0:
            return pb
        pl = numero(r.get('peso_liquido'), 0)
        fator = fator_map.get(r.get('Grupo'), 1)
        return pl * fator if pl else 0

    ger_proc['metal_liq_calc'] = ger_proc.apply(calc_metal_ger, axis=1)

    # -------------------- Material Fusão Corrigida --------------------
    # Buscar mapeamento da aba Apoio (colunas I:J = material_comercial -> material_fusao)
    apoio_mat_map = {}
    if EXCEL_PATH.exists():
        try:
            apoio_df = pd.read_excel(EXCEL_PATH, sheet_name='Apoio')
            # Colunas I e J são índices 8 e 9 (0-indexed)
            if apoio_df.shape[1] >= 10:
                col_comercial = apoio_df.columns[8]  # Coluna I
                col_fusao_map = apoio_df.columns[9]  # Coluna J
                for _, row in apoio_df.iterrows():
                    comercial = row[col_comercial]
                    fusao = row[col_fusao_map]
                    if pd.notna(comercial) and pd.notna(fusao):
                        apoio_mat_map[str(comercial).strip()] = str(fusao).strip()
                log(f'  ✓ {len(apoio_mat_map)} mapeamentos de material carregados da aba Apoio')
        except Exception as e:
            log(f'  ⚠ Não foi possível ler mapeamento de material da aba Apoio: {e}')
    
    # Aplicar regra: se material_fusao vazio, busca material_comercial no mapeamento
    materiais_nao_encontrados = set()
    
    def corrigir_material(row):
        mat_fusao = row.get('material_fusao')
        mat_comercial = row.get('material_comercial')
        
        # Se material_fusao já está preenchido, usa ele
        if pd.notna(mat_fusao) and str(mat_fusao).strip():
            return str(mat_fusao).strip()
        
        # Senão, tenta buscar material_comercial no mapeamento
        if pd.notna(mat_comercial):
            mat_com_str = str(mat_comercial).strip()
            if mat_com_str in apoio_mat_map:
                return apoio_mat_map[mat_com_str]
            else:
                # Não encontrou no mapeamento
                materiais_nao_encontrados.add(mat_com_str)
        
        return ''
    
    ger_proc['material_fusao_corrigida'] = ger_proc.apply(corrigir_material, axis=1)
    
    # Avisar sobre materiais não encontrados
    if materiais_nao_encontrados:
        log('\n⚠ MATERIAIS NÃO ENCONTRADOS NO MAPEAMENTO:')
        for mat in sorted(materiais_nao_encontrados):
            log(f'  - {mat}')
        log(f'\nTotal: {len(materiais_nao_encontrados)} materiais sem mapeamento')
        log('Adicione esses materiais na aba Apoio (colunas I e J) para corrigi-los.\n')

    # Remove colunas técnicas de lookup para não poluir o Excel
    for c in ['codigo_peca', 'Grupo_lookup', 'Desc_Grupo_lookup']:
        if c in ger_proc.columns:
            ger_proc.drop(columns=[c], inplace=True)

    ger_records = []
    for _, r in ger_proc.iterrows():
        try:
            semana = r.get('semana_fusao')
            data_fusao = r.get('data_fusao')
            fundido = pd.notna(data_fusao)
            ger_records.append({
                'os': int(r['id_ordem_servico']) if pd.notna(r.get('id_ordem_servico')) else '',
                'item': int(r['item']) if pd.notna(r.get('item')) else '',
                'chave': str(r.get('chave') or f"{r.get('id_ordem_servico','')}-{r.get('item','')}"),
                'cliente': short(r.get('nome_fantasia') or r.get('cliente') or ''),
                'descricao': short(r.get('descricao') or '', 65),
                'material': short(str(r.get('material_fusao_corrigida') or r.get('material_fusao') or r.get('material_comercial') or ''), 24),
                'semana': int(semana) if pd.notna(semana) else 0,
                'ano': str(int(semana))[:4] if pd.notna(semana) else '',
                'data_fusao': fmt_date(data_fusao),
                'fundido': bool(fundido),
                'situacao': 'Fundido' if fundido else 'Planejado',
                'peso_liq': round(numero(r.get('peso_liquido'), 0), 0),
                'peso_brt': round(numero(r.get('peso_bruto'), 0), 0),
                'metal_liq_calc': round(numero(r.get('metal_liq_calc'), 0), 0),
                'grupo': str(r.get('Grupo') or ''),
                'desc_grupo': short(str(r.get('Desc_Grupo') or 'Sem grupo'), 45),
            })
        except Exception:
            pass

    # -------------------- Acompanhamento / tabela carteira --------------------
    acomp_tab = acomp.copy()
    acomp_tab['etapa_atual'] = acomp_tab.apply(etapa_atual, axis=1)

    # BUSCAR PRAZOS DE dados_Gerenciamento (ger_proc)
    # Preparar dados de gerenciamento para merge
    ger_prazos = ger_proc[['id_ordem_servico', 'item', 'chave', 'prz_fundicao', 'prz_contratual']].copy()
    
    # Aplicar fallback: se prz_fundicao vazio, usar prz_contratual
    ger_prazos['prz_fundicao'] = ger_prazos.apply(
        lambda r: r['prz_contratual'] if pd.isna(r['prz_fundicao']) else r['prz_fundicao'],
        axis=1
    )
    
    # Fazer merge com acomp_tab para adicionar as colunas de prazo
    # Tentativa 1: merge por id_ordem_servico e item
    acomp_tab = acomp_tab.merge(
        ger_prazos[['id_ordem_servico', 'item', 'prz_fundicao', 'prz_contratual']],
        on=['id_ordem_servico', 'item'],
        how='left',
        suffixes=('', '_ger')
    )
    
    # Tentativa 2: para os que não encontraram, tentar por chave
    mask_sem_prazo = acomp_tab['prz_fundicao'].isna()
    if mask_sem_prazo.any() and 'chave' in acomp_tab.columns:
        acomp_com_chave = acomp_tab[mask_sem_prazo].copy()
        prazos_por_chave = ger_prazos[['chave', 'prz_fundicao', 'prz_contratual']].drop_duplicates('chave')
        acomp_com_chave = acomp_com_chave.drop(columns=['prz_fundicao', 'prz_contratual'], errors='ignore')
        acomp_com_chave = acomp_com_chave.merge(
            prazos_por_chave,
            on='chave',
            how='left'
        )
        # Atualizar os registros que encontraram por chave
        for col in ['prz_fundicao', 'prz_contratual']:
            if col in acomp_com_chave.columns:
                acomp_tab.loc[mask_sem_prazo, col] = acomp_com_chave[col].values

    # Enriquecimento leve da tabela com histórico
    ger_ref = ger_proc.copy()
    sort_cols = [c for c in ['data_fusao', 'dt_abertura'] if c in ger_ref.columns]
    if sort_cols:
        ger_ref = ger_ref.sort_values(sort_cols).drop_duplicates(['id_ordem_servico', 'item'], keep='last')
    else:
        ger_ref = ger_ref.drop_duplicates(['id_ordem_servico', 'item'], keep='last')

    cols_ger_ref = [c for c in [
        'id_ordem_servico','item','chave','cliente','nome_fantasia','descricao',
        'material_fusao_corrigida','material_fusao','material_comercial','semana_fusao','data_fusao'
    ] if c in ger_ref.columns]

    tabela = acomp_tab.merge(
        ger_ref[cols_ger_ref],
        on=['id_ordem_servico','item'],
        how='left',
        suffixes=('_acomp','_ger')
    )

    for col in ['prz_contratual', 'prz_fundicao', 'data_fusao']:
        if col in tabela.columns:
            tabela[col] = pd.to_datetime(tabela[col], errors='coerce')

    hoje = pd.Timestamp(date.today())
    if 'prz_contratual' in tabela.columns:
        tabela['dias_prazo'] = (tabela['prz_contratual'] - hoje).dt.days
    else:
        tabela['dias_prazo'] = None

    tabela_records = []
    for _, r in tabela.iterrows():
        try:
            semana = r.get('semana_fusao')
            tabela_records.append({
                'os': int(r['id_ordem_servico']) if pd.notna(r.get('id_ordem_servico')) else '',
                'item': int(r['item']) if pd.notna(r.get('item')) else '',
                'chave': str(r.get('chave') or f"{r.get('id_ordem_servico','')}-{r.get('item','')}"),
                'cliente': short(r.get('nome_fantasia') or r.get('cliente_ger') or r.get('cliente') or ''),
                'descricao': short(r.get('descricao') or r.get('descr_peca') or '', 65),
                'material': short(str(r.get('material_fusao_corrigida') or r.get('material_fusao') or r.get('material_comercial') or ''), 24),
                'semana': int(semana) if pd.notna(semana) else 0,
                'ano': str(int(semana))[:4] if pd.notna(semana) else '',
                'prz_fund': fmt_date(r.get('prz_fundicao')),
                'prz_cont': fmt_date(r.get('prz_contratual')),
                'dias': int(r['dias_prazo']) if pd.notna(r.get('dias_prazo')) else None,
                'etapa': str(r.get('etapa_atual', '')) if pd.notna(r.get('etapa_atual')) else '',
                'data_fusao': fmt_date(r.get('data_fusao')),
            })
        except Exception:
            pass

    # Semana atual no formato AAAASS para regra visual do gráfico
    iso = date.today().isocalendar()
    semana_atual = int(f"{iso.year}{iso.week:02d}")

    # Gera dados para aba Acompanhamento
    acomp_records = []
    hoje = pd.Timestamp(date.today())
    
    for _, r in tabela.iterrows():
        try:
            # USA AS NOVAS FUNÇÕES COM LÓGICA COMPLETA DO EXCEL
            etapa_calculada = calcular_etapa_atual(r)
            status_calculado = calcular_status(r, etapa_calculada, hoje)
            
            # LÓGICA DE PRAZO: se não tem fundição, usa contratual
            prz_fund = r.get('prz_fundicao')
            prz_cont = r.get('prz_contratual')
            if pd.isna(prz_fund) and pd.notna(prz_cont):
                prz_fund = prz_cont
            
            # LÓGICA DE ETAPAS: mostra "-" para etapas não aplicáveis
            # Etapa é aplicável se tem _prev OU _real preenchido
            def etapa_status(prev_campo, real_campo):
                prev_val = r.get(prev_campo)
                real_val = r.get(real_campo)
                # Se não tem nem prev nem real = não aplicável
                if pd.isna(prev_val) and pd.isna(real_val):
                    return '-'
                # Se tem real = concluído
                if pd.notna(real_val):
                    return 'OK'
                # Se tem prev mas não real = pendente
                return '!'
            
            acomp_records.append({
                'os': int(r['id_ordem_servico']) if pd.notna(r.get('id_ordem_servico')) else '',
                'item': int(r['item']) if pd.notna(r.get('item')) else '',
                'chave': str(r.get('chave') or f"{r.get('id_ordem_servico','')}-{r.get('item','')}"),
                'cliente': short(r.get('nome_fantasia') or r.get('cliente_ger') or r.get('cliente') or ''),
                'descricao': short(r.get('descricao') or r.get('descr_peca') or '', 50),
                'semana': int(r.get('semana_fusao')) if pd.notna(r.get('semana_fusao')) else 0,
                'desc_grupo': str(r.get('Desc_Grupo') or ''),
                'etapa': etapa_calculada,
                'status': status_calculado,
                'prz_fund': fmt_date(prz_fund),  # Usa o prazo com fallback
                'prz_cont': fmt_date(prz_cont),
                'projeto': etapa_status('projeto_prev', 'projeto_real'),
                'modelo': etapa_status('fabr_modelo_prev', 'fabr_modelo_real'),
                'moldagem': etapa_status('moldagem_prev', 'moldagem_real'),
                'fusao': etapa_status('fusao_prev', 'fusao_real'),
                'desmoldagem': etapa_status('desmoldagem_prev', 'desmoldagem_real'),
                'normalizacao_corte': etapa_status('normalizacao_corte_prev', 'normalizacao_corte_real'),
                'normalizacao': etapa_status('normalizacao_prev', 'normalizacao_real'),
                'tempera': etapa_status('tempera_prev', 'tempera_real'),
                'revenimento': etapa_status('revenimento_prev', 'revenimento_real'),
                'alivio_tensoes': etapa_status('alivio_tensoes_prev', 'alivio_tensoes_real'),
                'acabamento': etapa_status('acabamento_prev', 'acabamento_real'),
                'usinagem': etapa_status('usinagem_prev', 'usinagem_real'),
                'lib_fundido': etapa_status('lib_fundido_prev', 'lib_fundido_real'),
            })
        except Exception:
            pass

    return {
        'ger_records': ger_records,
        'tabela_records': tabela_records,
        'acomp_records': acomp_records,
        'today': str(date.today()),
        'total': len(ger_records),
        'semana_atual': semana_atual,
        'ger_processado': ger_proc,
        'acomp_original': acomp_tab,  # Usa acomp_tab que tem as colunas de prazo
    }

# =============================================================================
#  PARTE 3 — Salvar Excel (backup)
# =============================================================================

def limpar_para_excel(df):
    """
    Limpa valores problemáticos antes de salvar no Excel.
    Algumas descrições/material podem vir do banco com caracteres invisíveis
    que o openpyxl não aceita em células.
    """
    import re
    illegal = re.compile(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]')

    def limpa_valor(v):
        if pd.isna(v):
            return ''
        if isinstance(v, pd.Timestamp):
            return v.strftime('%d/%m/%Y')
        texto = str(v)
        texto = illegal.sub('', texto)
        return texto

    # Compatível com pandas novo e antigo
    if hasattr(df, 'map'):
        return df.map(limpa_valor)
    return df.applymap(limpa_valor)

def salvar_excel(acomp_original, ger_processado, fusao):
    log('\n[Backup] Salvando BaseDados.xlsx...')
    try:
        acomp_excel = limpar_para_excel(acomp_original)
        ger_excel = limpar_para_excel(ger_processado)
        fusao_excel = limpar_para_excel(fusao)
        apoio_excel = None
        if EXCEL_PATH.exists():
            try:
                apoio_excel = pd.read_excel(EXCEL_PATH, sheet_name='Apoio')
            except Exception:
                apoio_excel = None

        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as w:
            acomp_excel.to_excel(w, sheet_name='dados_Acomp', index=False)
            ger_excel.to_excel(w, sheet_name='dados_Gerenciamento', index=False)
            fusao_excel.to_excel(w, sheet_name='Vfusão', index=False)
            if apoio_excel is not None:
                apoio_excel.to_excel(w, sheet_name='Apoio', index=False)

        log(f'  ✓ Salvo em: {EXCEL_PATH}')
    except Exception as e:
        log(f'  ⚠ Não foi possível salvar: {e}')
        log('  O dashboard HTML será gerado normalmente mesmo sem o backup Excel.')

# =============================================================================
#  PARTE 4 — Gerar HTML
# =============================================================================

def gerar_html(data):
    today_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    data_json = json.dumps(data, ensure_ascii=False, default=str)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gestor Fundição — {today_str}</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f5f4f0;color:#1a1a18;font-size:14px}}
header{{background:#1a1a18;color:#f5f4f0;padding:.9rem 1.5rem;display:flex;align-items:center;justify-content:space-between}}
header h1{{font-size:15px;font-weight:500;display:flex;align-items:center;gap:8px}}
.hsub{{font-size:11px;color:#888780;margin-top:3px}}
.container{{max-width:1440px;margin:0 auto;padding:1.25rem}}
.kpi-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:1.25rem}}
.kpi{{background:#fff;border:0.5px solid rgba(0,0,0,.12);border-radius:8px;padding:.85rem 1rem}}
.klabel{{font-size:10px;color:#5f5e5a;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px}}
.kval{{font-size:26px;font-weight:600;line-height:1}}
.ksub{{font-size:11px;color:#888780;margin-top:4px}}
.ok .kval{{color:#3b6d11}}.info .kval{{color:#185fa5}}.warn .kval{{color:#ba7517}}
.row2{{display:grid;grid-template-columns:2fr 1fr;gap:10px;margin-bottom:10px}}
.card{{background:#fff;border:0.5px solid rgba(0,0,0,.12);border-radius:12px;padding:1rem 1.25rem}}
.ctitle{{font-size:11px;font-weight:600;color:#5f5e5a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:.75rem}}
.cwrap{{position:relative;width:100%;height:260px}}
.cwrap.tall{{height:310px}}
.tabs{{display:flex;gap:4px;margin-bottom:1rem;border-bottom:1px solid rgba(0,0,0,.12)}}
.tab{{padding:0.7rem 1.2rem;font-size:13px;font-weight:500;color:#5f5e5a;background:none;border:none;cursor:pointer;border-bottom:2px solid transparent;transition:all .2s}}
.tab:hover{{color:#1a1a18;background:rgba(0,0,0,.03)}}
.tab.active{{color:#185fa5;border-bottom-color:#185fa5}}
.tab-content{{display:none}}
.tab-content.active{{display:block}}
.filters{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:1rem;align-items:center}}
.filters select,.filters input{{border:0.5px solid rgba(0,0,0,.22);border-radius:8px;background:#fff;color:#1a1a18;padding:6px 10px;font-size:12px;outline:none}}
.filters select:focus,.filters input:focus{{border-color:#185fa5}}
.btn{{background:none;border:0.5px solid rgba(0,0,0,.22);border-radius:8px;padding:6px 12px;font-size:12px;color:#5f5e5a;cursor:pointer}}
.btn:hover{{background:#f1efe8}}
.twrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:#f1efe8;color:#5f5e5a;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;padding:6px 8px;text-align:left;border-bottom:0.5px solid rgba(0,0,0,.12);cursor:pointer;user-select:none;line-height:1.3}}
th:hover{{color:#1a1a18}}
td{{padding:7px 10px;border-bottom:0.5px solid rgba(0,0,0,.08);vertical-align:middle}}
tr:hover td{{background:#f8f7f4}}
.pag{{display:flex;align-items:center;gap:8px;margin-top:10px;justify-content:flex-end;font-size:12px;color:#5f5e5a}}
.pag button{{background:none;border:0.5px solid rgba(0,0,0,.22);border-radius:8px;padding:4px 12px;font-size:13px;cursor:pointer}}
.pag button:disabled{{opacity:.35;cursor:default}}
.legrow{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;font-size:11px;color:#5f5e5a}}
.ldot{{width:9px;height:9px;border-radius:2px;display:inline-block;vertical-align:middle;margin-right:3px}}
@media(max-width:900px){{
  header{{flex-wrap:wrap;gap:0.5rem}}
  .kpi-grid{{grid-template-columns:1fr}}
  .row2{{grid-template-columns:1fr}}
  .tabs{{overflow-x:auto;white-space:nowrap}}
  .filters{{flex-direction:column;align-items:stretch}}
  .filters select,.filters input{{width:100%}}
}}
@media(max-width:600px){{
  .container{{padding:0.75rem}}
  header{{padding:0.75rem}}
  .kval{{font-size:22px}}
  table{{font-size:11px}}
  th,td{{padding:6px 8px}}
}}
</style>
</head>
<body>
<header>
  <div>
    <h1><i class="ti ti-flame" style="color:#ef9f27;font-size:20px;" aria-hidden="true"></i>Gestor Fundição Moreno</h1>
    <div class="hsub">Dados atualizados em: <strong style="color:#d3d1c7">{today_str}</strong></div>
  </div>
</header>

<div class="container">
  <div class="tabs">
    <button class="tab active" onclick="switchTab('geral')">Geral</button>
    <button class="tab" onclick="switchTab('acomp')">Acompanhamento de Produção</button>
  </div>

  <!-- ABA GERAL -->
  <div class="tab-content active" id="tab-geral">
  <div class="kpi-grid">
    <div class="kpi info"><div class="klabel">Total de itens</div><div class="kval" id="k-total">—</div></div>
    <div class="kpi ok"><div class="klabel">Itens fundidos em T / Qnt</div><div class="kval" id="k-fund-combo">—</div></div>
    <div class="kpi warn"><div class="klabel">Itens planejados em T / Qnt</div><div class="kval" id="k-plan-combo">—</div></div>
  </div>

  <div class="filters">
    <span style="font-size:11px;color:#5f5e5a;"><i class="ti ti-filter" style="font-size:13px;vertical-align:-2px;" aria-hidden="true"></i> Filtros:</span>
    <select id="f-ano"><option value="">Todos os anos</option></select>
    <select id="f-mes"><option value="">Todos os meses</option></select>
    <select id="f-situacao"><option value="">Fundido + Planejado</option></select>
    <select id="f-grupo"><option value="">Todos os grupos</option></select>
    <select id="f-cliente"><option value="">Todos os clientes</option></select>
    <select id="f-semana"><option value="">Todas as semanas</option></select>
    <input id="f-search" type="text" placeholder="Buscar OS, peça, cliente, material..." style="min-width:220px;">
    <button class="btn" onclick="resetF()"><i class="ti ti-x" style="font-size:11px;" aria-hidden="true"></i> Limpar</button>
    <span style="margin-left:auto;font-size:11px;color:#888780;" id="clabel"></span>
  </div>

  <div class="row2">
    <div class="card">
      <div class="ctitle">Fusão (toneladas)</div>
      <div class="legrow">
        <span><span class="ldot" style="background:#185fa5"></span>Fundido</span>
        <span><span class="ldot" style="background:#ef9f27"></span>Planejado</span>
      </div>
      <div class="cwrap tall"><canvas id="cS" role="img" aria-label="Fusão por semana"></canvas></div>
    </div>
    <div class="card">
      <div class="ctitle">Grupo de material</div>
      <div class="cwrap"><canvas id="cG" role="img" aria-label="Grupo de material"></canvas></div>
      <div id="lG" style="margin-top:8px;"></div>
    </div>
  </div>

  <div class="card">
    <div class="ctitle">Itens da carteira <span style="font-size:10px;background:#f1efe8;color:#888780;border-radius:4px;padding:1px 7px;margin-left:6px;" id="tcnt">—</span></div>
    <div class="twrap">
      <table>
        <thead><tr>
          <th onclick="srt('os')">OS</th><th onclick="srt('item')">Item</th><th onclick="srt('cliente')">Cliente</th>
          <th onclick="srt('descricao')">Descrição</th><th onclick="srt('material')">Material</th><th onclick="srt('semana')">Sem. Fusão</th>
          <th onclick="srt('prz_fund')">Prz. Fund.</th><th onclick="srt('prz_cont')">Prz. Cont.</th><th onclick="srt('etapa')">Etapa Atual</th>
        </tr></thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
    <div class="pag"><span id="pinf"></span><button id="bpv" onclick="gp(-1)">&#8249;</button><button id="bnt" onclick="gp(1)">&#8250;</button></div>
  </div>
  </div>
  <!-- Fim ABA GERAL -->

  <!-- ABA ACOMPANHAMENTO -->
  <div class="tab-content" id="tab-acomp">
    <div class="filters">
      <span style="font-size:11px;color:#5f5e5a;"><i class="ti ti-filter" style="font-size:13px;vertical-align:-2px;" aria-hidden="true"></i> Filtros Acompanhamento:</span>
      <select id="f-status-acomp"><option value="">Todos status</option></select>
      <select id="f-cliente-acomp"><option value="">Todos clientes</option></select>
      <select id="f-etapa-acomp"><option value="">Todas etapas</option></select>
      <input id="f-search-acomp" type="text" placeholder="Buscar OS, peça..." style="min-width:180px;">
      <button class="btn" onclick="resetFAcomp()"><i class="ti ti-x" style="font-size:11px;"></i> Limpar</button>
    </div>
    
    <div class="card">
      <div class="ctitle">Acompanhamento de etapas <span style="font-size:10px;background:#f1efe8;color:#888780;border-radius:4px;padding:1px 7px;margin-left:6px;" id="tcnt-acomp">—</span></div>
      <div class="twrap">
        <table>
          <thead><tr>
            <th onclick="srtA('os')">OS</th><th onclick="srtA('item')">Item</th><th onclick="srtA('cliente')">Cliente</th>
            <th onclick="srtA('descricao')">Descrição</th><th onclick="srtA('status')">Status</th><th onclick="srtA('prz_fund')">Prz.<br>Fund.</th><th onclick="srtA('prz_cont')">Prz.<br>Cont.</th>
            <th>Proj.</th><th>Mod.</th><th>Mold.</th><th>Fusão</th><th>Desm.</th>
            <th>Norm.<br>Corte</th><th>Norm.</th><th>Têmp.</th><th>Reven.</th><th>Alív.<br>Tens.</th>
            <th>Acab.</th><th>Usin.</th><th>Lib.<br>Fund.</th>
          </tr></thead>
          <tbody id="tbody-acomp"></tbody>
        </table>
      </div>
      <div class="pag"><span id="pinf-acomp"></span><button id="bpv-acomp" onclick="gpAcomp(-1)">&#8249;</button><button id="bnt-acomp" onclick="gpAcomp(1)">&#8250;</button></div>
    </div>
  </div>
  <!-- Fim ABA ACOMPANHAMENTO -->
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>
const DATA={data_json};
const GER=DATA.ger_records||[];
const TAB=DATA.tabela_records||[];
const ACOMP=DATA.acomp_records||[];
const SEMANA_ATUAL=Number(DATA.semana_atual||0);
let gerFil=[...GER], tabFil=[...TAB], acompFil=[...ACOMP];
let sk='os', sd=1, p=1, pAcomp=1, skA='os', sdA=1;
const PP=30, DEFAULT_YEAR='2026';
const fmt=n=>n!=null?Number(n).toLocaleString('pt-BR'):'-';
let chartSemana=null, chartGrupo=null;
const MESES=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

function semanaParaMes(semana){{
  if(!semana)return '';
  const s=String(semana);
  if(s.length<6)return '';
  const ano=s.slice(0,4);
  const sem=parseInt(s.slice(4));
  const mes=Math.ceil((sem*7)/30.5);
  return mes>=1&&mes<=12?`${{ano}}-${{String(mes).padStart(2,'0')}}`:'';
}}

function switchTab(name){{
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(tc=>tc.classList.remove('active'));
  document.querySelector(`.tab[onclick*="'${{name}}'"]`).classList.add('active');
  document.getElementById(`tab-${{name}}`).classList.add('active');
}}
const totalLabelPlugin={{id:'totalLabelPlugin',afterDatasetsDraw(chart){{
  const {{ctx,scales:{{x,y}}}}=chart;
  ctx.save();
  ctx.textAlign='center';
  ctx.textBaseline='middle';
  ctx.textBaseline='alphabetic';
  ctx.font='11px Segoe UI';
  ctx.fillStyle='#5f5e5a';
  chart.data.labels.forEach((label,i)=>{{
    const total=chart.data.datasets.reduce((s,ds)=>s+(Number(ds.data[i])||0),0);
    if(total>0)ctx.fillText(total.toLocaleString('pt-BR',{{maximumFractionDigits:1}}),x.getPixelForValue(i),y.getPixelForValue(total)-6);
  }});
  ctx.restore();
}}}};
function anoSemana(r){{const s=String(r.semana||'');return s.length>=4?s.slice(0,4):'';}}
function popF(){{
  const add=(id,vals,def='')=>{{const el=document.getElementById(id);[...new Set(vals.filter(Boolean))].sort().forEach(v=>{{const o=document.createElement('option');o.value=v;o.textContent=v;el.appendChild(o);}});if(def&&[...el.options].some(o=>o.value===def))el.value=def;el.addEventListener('change',applyF);}};
  add('f-ano',GER.map(anoSemana),DEFAULT_YEAR);
  
  // Meses únicos
  const mesesUnicos=[...new Set(GER.map(r=>semanaParaMes(r.semana)).filter(Boolean))];
  mesesUnicos.sort().forEach(m=>{{
    const [a,mes]=m.split('-');
    const opt=document.createElement('option');
    opt.value=m;
    opt.textContent=`${{MESES[parseInt(mes)-1]}}/${{a}}`;
    document.getElementById('f-mes').appendChild(opt);
  }});
  document.getElementById('f-mes').addEventListener('change',applyF);
  
  add('f-situacao',GER.map(r=>r.situacao));
  add('f-grupo',GER.map(r=>r.desc_grupo));
  add('f-cliente',GER.map(r=>r.cliente));
  add('f-semana',GER.map(r=>r.semana).filter(Boolean).map(String));
  document.getElementById('f-search').addEventListener('input',applyF);
  
  // Filtros Acompanhamento
  const addA=(id,vals)=>{{
    const el=document.getElementById(id);
    [...new Set(vals.filter(Boolean))].sort().forEach(v=>{{
      const o=document.createElement('option');
      o.value=v;
      o.textContent=v;
      el.appendChild(o);
    }});
    el.addEventListener('change',applyFAcomp);
  }};
  addA('f-status-acomp',ACOMP.map(r=>r.status));
  addA('f-cliente-acomp',ACOMP.map(r=>r.cliente));
  addA('f-etapa-acomp',ACOMP.map(r=>r.etapa));
  document.getElementById('f-search-acomp').addEventListener('input',applyFAcomp);
}}
function passaFiltros(r, tipo){{
  const an=document.getElementById('f-ano').value;
  const ms=document.getElementById('f-mes').value;
  const si=document.getElementById('f-situacao').value;
  const gr=document.getElementById('f-grupo').value;
  const cl=document.getElementById('f-cliente').value;
  const sm=document.getElementById('f-semana').value;
  const q=document.getElementById('f-search').value.toLowerCase();
  
  if(an&&anoSemana(r)!==an)return false;
  if(ms&&semanaParaMes(r.semana)!==ms)return false;
  if(tipo==='ger'&&si&&r.situacao!==si)return false;
  if(tipo==='ger'&&gr&&r.desc_grupo!==gr)return false;
  if(cl&&r.cliente!==cl)return false;
  if(sm&&String(r.semana)!==sm)return false;
  if(q&&!(`${{r.os}} ${{r.chave}} ${{r.descricao}} ${{r.cliente}} ${{r.material||''}} ${{r.desc_grupo||''}}`).toLowerCase().includes(q))return false;
  return true;
}}
function applyF(){{
  gerFil=GER.filter(r=>passaFiltros(r,'ger'));
  tabFil=TAB.filter(r=>passaFiltros(r,'tab'));
  acompFil=ACOMP.filter(r=>passaFiltros(r,'tab'));
  p=1;pAcomp=1;renderKpis();renderT();renderTAcomp();charts();
  document.getElementById('clabel').textContent=`${{gerFil.length}} produção · ${{tabFil.length}} carteira`;
}}
function resetF(){{
  ['f-situacao','f-grupo','f-cliente','f-semana','f-mes'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('f-ano').value=DEFAULT_YEAR;
  document.getElementById('f-search').value='';
  applyF();
}}
function renderKpis(){{
  const total=gerFil.length;
  const fund=gerFil.filter(r=>r.fundido&&r.semana);
  const plan=gerFil.filter(r=>!r.fundido&&r.semana);
  const fundPeso=fund.reduce((s,r)=>s+(Number(r.metal_liq_calc)||0),0);
  const planPeso=plan.reduce((s,r)=>s+(Number(r.metal_liq_calc)||0),0);
  document.getElementById('k-total').textContent=fmt(total);
  document.getElementById('k-fund-combo').textContent=`${{Math.round(fundPeso/1000).toLocaleString('pt-BR')}} t / ${{fund.length.toLocaleString('pt-BR')}}`;
  document.getElementById('k-plan-combo').textContent=`${{Math.round(planPeso/1000).toLocaleString('pt-BR')}} t / ${{plan.length.toLocaleString('pt-BR')}}`;
}}
function comparePrazo(a,b){{
  // Compara datas no formato dd/mm/yyyy
  const parseDate=d=>{{
    if(!d||d==='-')return new Date(9999,11,31);
    const[dia,mes,ano]=d.split('/').map(Number);
    return new Date(ano,mes-1,dia);
  }};
  const da=parseDate(a);
  const db=parseDate(b);
  return da-db;
}}
function srt(k){{if(sk===k)sd*=-1;else{{sk=k;sd=1;}}renderT();}}
function renderT(){{
  const sorted=[...tabFil].sort((a,b)=>{{
    // Se não há ordenação customizada, ordena por prazo fundição e depois cliente
    if(sk==='os'){{
      const prazoDiff=comparePrazo(a.prz_fund,b.prz_fund);
      if(prazoDiff!==0)return prazoDiff;
      return(a.cliente||'').localeCompare(b.cliente||'');
    }}
    // Ordenação customizada
    let va=a[sk],vb=b[sk];
    if(va==null)return 1;
    if(vb==null)return-1;
    return typeof va==='string'?va.localeCompare(vb)*sd:(va-vb)*sd;
  }});
  const sl=sorted.slice((p-1)*PP,p*PP);
  document.getElementById('tbody').innerHTML=sl.map(r=>`<tr><td><strong>${{r.os}}</strong></td><td style="color:#888780">${{r.item}}</td><td>${{r.cliente}}</td><td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${{r.descricao}}">${{r.descricao}}</td><td style="font-size:11px;color:#888780;">${{r.material}}</td><td style="text-align:center;font-weight:500;">${{r.semana||'-'}}</td><td style="font-size:11px;">${{r.prz_fund}}</td><td style="font-size:11px;">${{r.prz_cont}}</td><td style="font-size:11px;color:#5f5e5a;">${{r.etapa}}</td></tr>`).join('');
  const pages=Math.max(1,Math.ceil(tabFil.length/PP));document.getElementById('tcnt').textContent=tabFil.length;document.getElementById('pinf').textContent=`Pág. ${{p}} de ${{pages}}`;document.getElementById('bpv').disabled=p<=1;document.getElementById('bnt').disabled=p>=pages;
}}
function gp(d){{p+=d;renderT();}}
function srtA(k){{if(skA===k)sdA*=-1;else{{skA=k;sdA=1;}}renderTAcomp();}}
function renderTAcomp(){{
  const sorted=[...acompFil].sort((a,b)=>{{
    // Se não há ordenação customizada, ordena por prazo fundição e depois cliente
    if(skA==='os'){{
      const prazoDiff=comparePrazo(a.prz_fund,b.prz_fund);
      if(prazoDiff!==0)return prazoDiff;
      return(a.cliente||'').localeCompare(b.cliente||'');
    }}
    // Ordenação customizada
    let va=a[skA],vb=b[skA];
    if(va==null)return 1;
    if(vb==null)return-1;
    return typeof va==='string'?va.localeCompare(vb)*sdA:(va-vb)*sdA;
  }});
  const sl=sorted.slice((pAcomp-1)*PP,pAcomp*PP);
  const ck=v=>{{
    if(v==='OK')return`<span style="color:#3b6d11;font-weight:700">✓</span>`;
    if(v==='!')return`<span style="color:#e24b4a;font-weight:700">!</span>`;
    return`<span style="color:#888780">—</span>`;  // "-" para não aplicável
  }};
  const statusBadge=s=>{{
    if(s==='Finalizado')return'<span style="background:#d1f4e0;color:#1e5f3a;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">✓ FINALIZADO</span>';
    if(s==='NO PRAZO')return'<span style="background:#e6f1fb;color:#0c447c;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">NO PRAZO</span>';
    if(s==='No Prazo')return'<span style="background:#e6f1fb;color:#0c447c;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">NO PRAZO</span>';
    if(s==='POT. ATRASO')return'<span style="background:#faeeda;color:#633806;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">POT. ATRASO</span>';
    if(s==='Potencial Atraso')return'<span style="background:#faeeda;color:#633806;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">POT. ATRASO</span>';
    if(s==='Atrasado')return'<span style="background:#fcebeb;color:#791f1f;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">ATRASADO</span>';
    if(s==='Fundido')return'<span style="background:#eaf3de;color:#27500a;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">FUNDIDO</span>';
    if(s==='Liberado')return'<span style="background:#d1f4e0;color:#1e5f3a;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">LIBERADO</span>';
    if(s==='Planejado')return'<span style="background:#f1efe8;color:#5f5e5a;padding:2px 6px;border-radius:4px;font-size:10px;">PLANEJADO</span>';
    return`<span style="background:#f1efe8;color:#5f5e5a;padding:2px 6px;border-radius:4px;font-size:10px;">${{s}}</span>`;
  }};
  document.getElementById('tbody-acomp').innerHTML=sl.map(r=>`<tr>
    <td><strong>${{r.os}}</strong></td><td style="color:#888780">${{r.item}}</td><td>${{r.cliente}}</td>
    <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${{r.descricao}}">${{r.descricao}}</td>
    <td>${{statusBadge(r.status)}}</td>
    <td style="font-size:11px;">${{r.prz_fund}}</td>
    <td style="font-size:11px;">${{r.prz_cont}}</td>
    <td style="text-align:center;">${{ck(r.projeto)}}</td><td style="text-align:center;">${{ck(r.modelo)}}</td>
    <td style="text-align:center;">${{ck(r.moldagem)}}</td><td style="text-align:center;">${{ck(r.fusao)}}</td>
    <td style="text-align:center;">${{ck(r.desmoldagem)}}</td>
    <td style="text-align:center;">${{ck(r.normalizacao_corte)}}</td><td style="text-align:center;">${{ck(r.normalizacao)}}</td>
    <td style="text-align:center;">${{ck(r.tempera)}}</td><td style="text-align:center;">${{ck(r.revenimento)}}</td>
    <td style="text-align:center;">${{ck(r.alivio_tensoes)}}</td>
    <td style="text-align:center;">${{ck(r.acabamento)}}</td><td style="text-align:center;">${{ck(r.usinagem)}}</td>
    <td style="text-align:center;">${{ck(r.lib_fundido)}}</td>
  </tr>`).join('');
  const pages=Math.max(1,Math.ceil(acompFil.length/PP));
  document.getElementById('tcnt-acomp').textContent=acompFil.length;
  document.getElementById('pinf-acomp').textContent=`Pág. ${{pAcomp}} de ${{pages}}`;
  document.getElementById('bpv-acomp').disabled=pAcomp<=1;
  document.getElementById('bnt-acomp').disabled=pAcomp>=pages;
}}
function gpAcomp(d){{pAcomp+=d;renderTAcomp();}}
function applyFAcomp(){{
  const st=document.getElementById('f-status-acomp').value;
  const cl=document.getElementById('f-cliente-acomp').value;
  const et=document.getElementById('f-etapa-acomp').value;
  const q=document.getElementById('f-search-acomp').value.toLowerCase();
  acompFil=ACOMP.filter(r=>{{
    if(st&&r.status!==st)return false;
    if(cl&&r.cliente!==cl)return false;
    if(et&&r.etapa!==et)return false;
    if(q&&!(`${{r.os}} ${{r.descricao}} ${{r.cliente}}`).toLowerCase().includes(q))return false;
    return true;
  }});
  pAcomp=1;renderTAcomp();
}}
function resetFAcomp(){{
  ['f-status-acomp','f-cliente-acomp','f-etapa-acomp'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('f-search-acomp').value='';
  applyFAcomp();
}}
function destroyCharts(){{if(chartSemana)chartSemana.destroy();if(chartGrupo)chartGrupo.destroy();chartSemana=chartGrupo=null;}}
function charts(){{
  destroyCharts();
  const bySemana={{}};
  gerFil.forEach(r=>{{const k=String(r.semana||'');if(!k)return;if(!bySemana[k])bySemana[k]={{semana:Number(k),fundido:0,planejado:0}};const ton=(Number(r.metal_liq_calc)||0)/1000;if(r.fundido)bySemana[k].fundido+=ton;else bySemana[k].planejado+=ton;}});
  const semanas=Object.values(bySemana).sort((a,b)=>a.semana-b.semana);
  const labels=semanas.map(s=>String(s.semana).slice(4));
  chartSemana=new Chart(document.getElementById('cS'),{{type:'bar',plugins:[totalLabelPlugin],data:{{labels,datasets:[{{label:'Fundido',data:semanas.map(s=>Math.round(s.fundido)),backgroundColor:'#185fa5',borderWidth:0}},{{label:'Planejado',data:semanas.map(s=>Math.round(s.planejado)),backgroundColor:'#ef9f27',borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:false,layout:{{padding:{{top:22}}}},onClick:(e,activeEls)=>{{
    if(activeEls.length>0){{
      const idx=activeEls[0].index;
      const semanaClicada=semanas[idx].semana;
      document.getElementById('f-semana').value=String(semanaClicada);
      applyF();
      document.querySelector('.tab[onclick*="geral"]').scrollIntoView({{behavior:'smooth'}});
    }}
  }},plugins:{{legend:{{display:true}},tooltip:{{callbacks:{{label:c=>`${{c.dataset.label}}: ${{c.parsed.y}} t`}}}}}},scales:{{x:{{stacked:true,ticks:{{font:{{size:10}},autoSkip:false,maxRotation:45,color:'#888780'}},grid:{{display:false}}}},y:{{stacked:true,ticks:{{font:{{size:10}},color:'#888780',callback:v=>v+'t'}},grid:{{color:'rgba(0,0,0,0.05)'}}}}}}}}}});
  const grupos={{}};let total=0;gerFil.forEach(r=>{{const k=r.desc_grupo||'Sem grupo';const v=Number(r.metal_liq_calc)||0;grupos[k]=(grupos[k]||0)+v;total+=v;}});
  const ge=Object.entries(grupos).map(([key,val])=>({{key,val}})).sort((a,b)=>b.val-a.val);
  const gc=['#185fa5','#3b6d11','#ba7517','#d4537e','#1d9e75','#d85a30','#634ab7','#5f5e5a','#639922','#e24b4a'];
  chartGrupo=new Chart(document.getElementById('cG'),{{type:'doughnut',data:{{labels:ge.map(e=>e.key),datasets:[{{data:ge.map(e=>Math.round(e.val/1000)),backgroundColor:gc,borderWidth:2,borderColor:'#fff'}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>`${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} t (${{total?Math.round((ge[c.dataIndex].val/total)*100):0}}%)`}}}}}},cutout:'58%'}}}});
  document.getElementById('lG').innerHTML='<div style="display:flex;flex-wrap:wrap;gap:5px;">'+ge.map((e,i)=>`<span style="display:flex;align-items:center;gap:3px;font-size:10px;color:#5f5e5a"><span style="width:8px;height:8px;border-radius:2px;background:${{gc[i%gc.length]}}"></span>${{e.key}} (${{total?Math.round(e.val/total*100):0}}%)</span>`).join('')+'</div>';
}}
popF();applyF();
</script>
</body>
</html>"""

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    log(f'  ✓ Dashboard salvo em: {HTML_PATH}')


# =============================================================================
#  EXECUÇÃO PRINCIPAL
# =============================================================================
if __name__ == '__main__':
    log()
    log('╔══════════════════════════════════════════════════╗')
    log('║    GESTOR FUNDIÇÃO — Atualização de Dados        ║')
    log('╚══════════════════════════════════════════════════╝')
    log()

    try:
        log('[INÍCIO] Verificando conexão com banco de dados...')
        acomp, ger, fusao = buscar_dados()

        log('\n[5/5] Processando carteira e gerando colunas calculadas...')
        data = processar(acomp, ger, fusao)
        ger_processado = data.get('ger_processado')
        acomp_original = data.get('acomp_original')

        resp = perguntar('\nSalvar backup no BaseDados.xlsx? (s/n) [n]: ', 'n')
        if resp == 's':
            salvar_excel(acomp_original, ger_processado, fusao)

        log('\nGerando dashboard HTML...')
        data.pop('ger_processado', None)
        data.pop('acomp_original', None)
        gerar_html(data)

        resp2 = perguntar('\nAbrir o dashboard agora? (s/n) [s]: ', 's')
        if resp2 == 's':
            import webbrowser
            webbrowser.open(str(HTML_PATH))

    except KeyboardInterrupt:
        log('\n\n[CANCELADO] Você pressionou Ctrl+C.')
    except Exception as e:
        log('\n╔════════════════════════════════════════════════════════════╗')
        log('║  [ERRO] O script encontrou um problema                     ║')
        log('╚════════════════════════════════════════════════════════════╝')
        log(f'\nErro: {str(e)}')
        log('\n--- Detalhes técnicos (copie e envie se precisar de ajuda) ---')
        log(traceback.format_exc())
        log('--- Fim dos detalhes técnicos ---')

    log('\n✓ Script finalizado.')
    pausar()
