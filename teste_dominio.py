import time
import os
import sys
import glob
import ctypes
import pandas as pd
from pywinauto.keyboard import send_keys

ARQUIVO_LOG = "acumuladores_processados.txt"

# --- SISTEMA DE SEGURANÇA ---
def verificar_seguranca():
    """Verifica se a tecla ESC (código 0x1B) está sendo pressionada no momento."""
    # O bit 0x8000 garante que a tecla está fisicamente pressionada no instante da checagem
    if ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000:
        print("\n\n" + "="*60)
        print("[!] PARADA DE EMERGÊNCIA ACIONADA!")
        print("A tecla ESC foi detetada. O robô foi abortado imediatamente.")
        print("O seu teclado e rato estão livres.")
        print("="*60)
        os._exit(1) # Corta o processo instantaneamente pela raiz

def pausa_segura(segundos):
    """Substitui o time.sleep() padrão para vigiar o ESC a cada 0.1 segundos durante as esperas."""
    fim = time.time() + segundos
    while time.time() < fim:
        verificar_seguranca()
        time.sleep(0.1)

# --- FUNÇÕES BASE ---
def pausar_e_sair():
    input("\nPressione ENTER para fechar esta janela...")
    sys.exit(1)

def teclar(comando, espera_depois=0.6, pausa_interna=0.20, com_espacos=False):
    """Envia teclas com cadência controlada e trava de segurança blindada."""
    verificar_seguranca()
    send_keys(comando, pause=pausa_interna, with_spaces=com_espacos)
    pausa_segura(espera_depois)

def teclar_multiplo(comando, vezes, espera_depois=0.35):
    """Envia uma tecla repetidas vezes (1 por 1)."""
    for _ in range(vezes):
        teclar(comando, espera_depois=espera_depois)

def extrair_inteiro_str(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().replace('.0', '')
    return texto if texto.isdigit() else ""

def carregar_e_classificar_planilha():
    arquivos_excel = [f for f in glob.glob("*.xlsx") if not f.startswith("~$")]
    
    if not arquivos_excel:
        print("[ERRO] Nenhuma planilha (.xlsx) foi encontrada na pasta.")
        pausar_e_sair()
        
    arquivo_atual = max(arquivos_excel, key=os.path.getmtime)
    print(f"\n[*] Planilha detectada: '{arquivo_atual}'")
    
    df_teste = pd.read_excel(arquivo_atual, nrows=0)
    
    if 'Codigo' in df_teste.columns and 'Geral_Faturamento' in df_teste.columns:
        print("[*] Status: A planilha já está limpa e formatada. Iniciando leitura...")
        return pd.read_excel(arquivo_atual)
        
    print("[*] Status: Planilha BRUTA detectada. Formatando e limpando os dados agora...")
    
    try:
        df_raw = pd.read_excel(arquivo_atual, sheet_name='Acumuladores Simples Nacional', header=None)
    except Exception as e:
        print(f"[ERRO] A planilha não possui a aba 'Acumuladores Simples Nacional' ou formato desconhecido.\nDetalhe: {e}")
        pausar_e_sair()
        
    data = df_raw.iloc[7:].copy()

    col_mapping = {
        0: 'Codigo', 1: 'Nome',
        8: 'Geral_Faturamento', 9: 'Geral_Receita_Bruta', 10: 'Geral_Devolucao', 11: 'Geral_Retorno_Remessa',
        18: 'Notas_Gerar_Parcelas', 20: 'Notas_Digitacao_Estoque', 26: 'Opcoes_Gerar_Bens_Patrimonio',
        29: 'Lanc_Entrada', 30: 'Lanc_Saida', 31: 'Lanc_Servico', 37: 'Lanc_Outros',
        38: 'NotasII_Info_Complementares', 39: 'NotasII_Observacoes', 40: 'NotasII_Proc_Referenciado',
        41: 'NotasII_Doc_Arrecadacao', 42: 'NotasII_NF_Referenciada',
        44: 'Imposto_1', 46: 'Imposto_2', 48: 'Imposto_3', 50: 'Imposto_4',
        68: 'Estadual_Codigo_Municipio',
        76: 'CFOP_1', 77: 'CFOP_2', 78: 'CFOP_3',
        85: 'Contab_Conta_Compensacao', 86: 'Contab_Conta_Debito', 87: 'Contab_Conta_Credito'
    }

    df_clean = data[list(col_mapping.keys())].copy()
    df_clean.columns = list(col_mapping.values())

    num_cols = [
        'Imposto_1', 'Imposto_2', 'Imposto_3', 'Imposto_4',
        'CFOP_1', 'CFOP_2', 'CFOP_3',
        'Contab_Conta_Debito', 'Contab_Conta_Credito'
    ]
    
    for col in num_cols:
        df_clean[col] = df_clean[col].apply(
            lambda x: '' if pd.isna(x) or not str(x).strip().replace('.0', '').isdigit()
            else int(float(x))
        )

    for col in df_clean.columns:
        if col not in ['Codigo', 'Nome'] + num_cols:
            df_clean[col] = df_clean[col].apply(lambda x: 'Sim' if str(x).strip() == 'Sim' else 'Não')

    df_clean['Codigo'] = df_clean['Codigo'].astype(int)
    df_clean['Nome'] = df_clean['Nome'].astype(str).str.strip()
    
    arquivo_limpo = "Planilha_Formatada_Pelo_Robo.xlsx"
    df_clean.to_excel(arquivo_limpo, index=False)
    print(f"[*] Cópia estruturada salva como '{arquivo_limpo}' com sucesso!\n")
    
    return df_clean

# --- 1. LEITURA DA PLANILHA E DO HISTÓRICO TXT ---
df_completo = carregar_e_classificar_planilha()

processados = []
if os.path.exists(ARQUIVO_LOG):
    with open(ARQUIVO_LOG, "r") as f:
        processados = [linha.strip() for linha in f.readlines() if linha.strip()]
        
if processados:
    print(f"\n[*] Encontrados {len(processados)} acumuladores já processados no arquivo '{ARQUIVO_LOG}'. Eles serão ignorados.")

df_completo['Codigo_Str'] = df_completo['Codigo'].astype(str)
df_filtrado = df_completo[~df_completo['Codigo_Str'].isin(processados)]

if df_filtrado.empty:
    print("\n[✓] Todos os acumuladores da planilha já foram processados!")
    pausar_e_sair()

# --- 2. CONTAGEM REGRESSIVA PARA VOCÊ MUDAR DE TELA ---
print("\n" + "="*50)
print(" PREPARAÇÃO PARA DIGITAÇÃO")
print("="*50)
print("1. Abra a tela de Acumuladores no Domínio.")
print("2. Deixe o cursor a piscar no campo 'Código'.")
print("3. Pressione a tecla 'ESC' a qualquer momento para abortar.")

for i in range(10, 0, -1):
    print(f"Iniciando em {i} segundos... Mude para o Domínio! ", end='\r')
    pausa_segura(1.0)

print("\n\n[ROBÔ TRABALHANDO] Por favor, não mexa no teclado nem no rato!")

# --- 3. LOOP DE PREENCHIMENTO ---
try:
    for index, linha in df_filtrado.iterrows():
        if pd.isna(linha['Codigo']):
            continue
            
        codigo = str(int(linha['Codigo']))
        nome = str(linha['Nome']).strip()

        print(f"\n[Processando] Acumulador: {codigo} - {nome}")

        cols_imp = ['Imposto_1', 'Imposto_2', 'Imposto_3', 'Imposto_4']
        impostos = [extrair_inteiro_str(linha[c]) for c in cols_imp if extrair_inteiro_str(linha[c]) != ""]

        cols_cfop = ['CFOP_1', 'CFOP_2', 'CFOP_3']
        cfops = [extrair_inteiro_str(linha[c]) for c in cols_cfop if extrair_inteiro_str(linha[c]) != ""]

        conta_debito = extrair_inteiro_str(linha['Contab_Conta_Debito'])
        conta_credito = extrair_inteiro_str(linha['Contab_Conta_Credito'])

        # --- CABEÇALHO ---
        teclar(codigo, espera_depois=0.8, pausa_interna=0.10)
        teclar("{TAB}", espera_depois=0.8)  
        teclar(nome, espera_depois=0.8, pausa_interna=0.05, com_espacos=True)
        teclar_multiplo("{TAB}", 4, espera_depois=0.4)
        teclar(nome, espera_depois=0.8, pausa_interna=0.05, com_espacos=True)
        teclar("{TAB}", espera_depois=0.8)

        # --- ABA GERAL ---
        cols_geral = ['Geral_Faturamento', 'Geral_Receita_Bruta', 'Geral_Devolucao', 'Geral_Retorno_Remessa']
        pos_geral_ativas = [i + 1 for i, c in enumerate(cols_geral) if str(linha[c]).strip() == "Sim"]
        if pos_geral_ativas:
            ultima_pos_geral = max(pos_geral_ativas)
            for pos_atual in range(1, ultima_pos_geral + 1):
                teclar("{TAB}", espera_depois=0.4)
                if pos_atual in pos_geral_ativas:
                    teclar("{SPACE}", espera_depois=0.4)
            teclar_multiplo("+{TAB}", ultima_pos_geral, espera_depois=0.4)

        # --- ABA NOTAS > GERAL E OPÇÕES ---
        teclar("{RIGHT}", espera_depois=0.8)  
        teclar("{TAB}", espera_depois=0.6)    
        teclar_multiplo("{LEFT}", 3, espera_depois=0.3)

        pos_notas_geral = []
        if str(linha['Notas_Gerar_Parcelas']).strip() == "Sim":
            pos_notas_geral.append(1)
        if str(linha['Notas_Digitacao_Estoque']).strip() == "Sim":
            pos_notas_geral.append(3)
        
        if pos_notas_geral:
            ult_pos_notas = max(pos_notas_geral)
            for i in range(1, ult_pos_notas + 1):
                teclar("{TAB}", espera_depois=0.4)
                if i in pos_notas_geral:
                    teclar("{SPACE}", espera_depois=0.4)
            teclar_multiplo("+{TAB}", ult_pos_notas, espera_depois=0.4)

        if str(linha['Opcoes_Gerar_Bens_Patrimonio']).strip() == "Sim":
            teclar("{RIGHT}", espera_depois=0.6)  
            teclar("{TAB}", espera_depois=0.5)    
            teclar("{SPACE}", espera_depois=0.6) 
            teclar("+{TAB}", espera_depois=0.6)   
            teclar_multiplo("{RIGHT}", 2, espera_depois=0.4)  
        else:
            teclar_multiplo("{RIGHT}", 3, espera_depois=0.4)  

        # --- ABA NOTAS > LANÇAMENTOS ---
        teclar("{TAB}", espera_depois=0.6)
        teclar("{SPACE}", espera_depois=1.0)  

        mapa_lancamentos = [(1, 'Lanc_Entrada'), (2, 'Lanc_Saida'), (3, 'Lanc_Servico'), (9, 'Lanc_Outros')]
        posicoes_ativas = [pos for pos, col_name in mapa_lancamentos if str(linha[col_name]).strip() == "Sim"]
        ultima_pos = max(posicoes_ativas) if posicoes_ativas else 1

        for pos_atual in range(1, ultima_pos + 1):
            teclar("{TAB}", espera_depois=0.4)
            if pos_atual in posicoes_ativas:
                teclar("{SPACE}", espera_depois=0.4)

        teclar_multiplo("+{TAB}", ultima_pos + 1, espera_depois=0.4)
        teclar_multiplo("{LEFT}", 3, espera_depois=0.3)
        teclar("+{TAB}", espera_depois=0.5)

        # --- ABA NOTAS II ---
        teclar("{RIGHT}", espera_depois=0.8)  
        teclar("{TAB}", espera_depois=0.6)
        if str(linha['NotasII_Info_Complementares']).strip() == "Sim":
            teclar("{SPACE}", espera_depois=1.0)

        cols_notas_ii = ['NotasII_Observacoes', 'NotasII_Proc_Referenciado', 'NotasII_Doc_Arrecadacao', 'NotasII_NF_Referenciada']
        pos_n2_ativas = [i + 1 for i, c in enumerate(cols_notas_ii) if str(linha[c]).strip() == "Sim"]
        
        if pos_n2_ativas:
            ult_n2 = max(pos_n2_ativas)
            for i in range(1, ult_n2 + 1):
                teclar("{TAB}", espera_depois=0.4)
                if i in pos_n2_ativas:
                    teclar("{SPACE}", espera_depois=0.4)
            teclar_multiplo("+{TAB}", ult_n2 + 1, espera_depois=0.4)
        else:
            teclar("+{TAB}", espera_depois=0.4)

        # --- ABA IMPOSTOS E ESTADUAL ---
        teclar("{RIGHT}", espera_depois=0.8)  
        if len(impostos) > 0:
            teclar("{TAB}", espera_depois=0.6)    
            teclar("{SPACE}", espera_depois=0.8)  
            for idx_imp, imp in enumerate(impostos):
                teclar(imp, espera_depois=1.0, pausa_interna=0.10)
                if idx_imp < len(impostos) - 1:
                    teclar("{TAB}", espera_depois=0.5)
                    teclar("{TAB}", espera_depois=0.5)
                    teclar("{SPACE}", espera_depois=0.8)
                else:
                    qtd_shift_tabs = (2 * len(impostos)) - 1
                    teclar_multiplo("+{TAB}", qtd_shift_tabs, espera_depois=0.45)

        sped_habilitado = any(imp in ['28', '26'] for imp in impostos)
        setas_ate_estadual = 2 if sped_habilitado else 1
        setas_ate_cfop = 4 if sped_habilitado else 3

        if str(linha['Estadual_Codigo_Municipio']).strip() == "Sim":
            teclar_multiplo("{RIGHT}", setas_ate_estadual, espera_depois=0.4)  
            teclar("{TAB}", espera_depois=0.6)
            teclar("{SPACE}", espera_depois=0.6)
            teclar("+{TAB}", espera_depois=0.6)                                
            teclar_multiplo("{RIGHT}", 2, espera_depois=0.4)                   
        else:
            teclar_multiplo("{RIGHT}", setas_ate_cfop, espera_depois=0.4)      

        # --- ABA CFOP/CFPS ---
        if len(cfops) > 0:
            teclar("{TAB}", espera_depois=0.6)    
            teclar("{SPACE}", espera_depois=0.8)  
            teclar("{TAB}", espera_depois=0.5)    
            teclar("{TAB}", espera_depois=0.5)    
            teclar("{TAB}", espera_depois=0.5)    
            teclar("{SPACE}", espera_depois=0.8)  

            for idx_cfop, cfop in enumerate(cfops):
                teclar(cfop, espera_depois=1.2, pausa_interna=0.10)
                if idx_cfop < len(cfops) - 1:
                    teclar("{TAB}", espera_depois=0.6)
                    teclar("{SPACE}", espera_depois=0.8)
                else:
                    qtd_shift_cfop = len(cfops) + 2
                    teclar_multiplo("+{TAB}", qtd_shift_cfop, espera_depois=0.45)

        # --- ABA CONTABILIDADE ---
        teclar("{RIGHT}", espera_depois=0.8)  
        teclar("{TAB}", espera_depois=0.6)    
        teclar("{LEFT}", espera_depois=0.4)

        if str(linha['Contab_Conta_Compensacao']).strip() == "Sim":
            teclar_multiplo("{TAB}", 2, espera_depois=0.5)
            teclar("{SPACE}", espera_depois=0.6)
            teclar_multiplo("+{TAB}", 2, espera_depois=0.5)

        teclar("{RIGHT}", espera_depois=0.8)  
        teclar("{TAB}", espera_depois=0.6)
        if conta_debito != "":
            teclar(conta_debito, espera_depois=1.2, pausa_interna=0.10)

        teclar("{TAB}", espera_depois=0.8)
        if conta_credito != "":
            teclar(conta_credito, espera_depois=1.2, pausa_interna=0.10)

        # --- SALVAR E REGISTAR NO TXT ---
        teclar("{TAB}", espera_depois=1.2)
        teclar_multiplo("{TAB}", 19, espera_depois=0.30)
        
        teclar("{SPACE}", espera_depois=0.1) 
        pausa_segura(4.0) 
        
        with open(ARQUIVO_LOG, "a") as f:
            f.write(f"{codigo}\n")
            
        print(f"[✓] Acumulador {codigo} gravado e salvo no histórico!")

except Exception as e:
    print("\n" + "="*60)
    print("[ERRO INESPERADO] O ROBÔ PERDEU O CONTROLO DO SISTEMA")
    print(f"Detalhe Técnico: {str(e)}")
    pausar_e_sair()

print("\n" + "="*50)
print("[✓] PROCESSO FINALIZADO COM SUCESSO!")
print("Todos os acumuladores da planilha foram preenchidos.")
print("="*50)
pausar_e_sair()