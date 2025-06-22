import pandas as pd
import openpyxl
import os
from pathlib import Path
from datetime import datetime, timedelta
import logging

# Configuração básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constantes
BASE_DIR = Path(r"C:\Users\rmbotelho\Documents\Plano PM v2")
ORIGINAL_FILE = BASE_DIR / "ih08.xlsx"
FILTERED_FILE = BASE_DIR / "ih08_filtrada.xlsx"
IP03_FILE = BASE_DIR / "ip03.xlsx"
MERGED_FILE = BASE_DIR / "ih08_ip03_merged.xlsx"
SEM_PLANO_FILE = BASE_DIR / "equipamentos_sem_plano.xlsx"  

# Configuração dos intervalos

def tratamento_dados():
    """Processa o arquivo original e cria versão filtrada."""
    try:
        if not ORIGINAL_FILE.exists():
            raise FileNotFoundError(f"Arquivo original não encontrado: {ORIGINAL_FILE}")
        
        df = pd.read_excel(ORIGINAL_FILE, sheet_name=0)
        
        # Filtros combinados
        mask = (
            (df['Status sistema'] == 'DEPS ECLI') &
            (~df['Equipamento'].str.startswith('GEBRA', na=False))
        )
        df_filtrado = df[mask]
        
        if df_filtrado.empty:
            logging.warning("Nenhum registro encontrado após filtragem.")
            return False
        
        df_filtrado.to_excel(FILTERED_FILE, index=False)
        logging.info(f"Planilha filtrada salva em: {FILTERED_FILE}")
        
        try:
            ORIGINAL_FILE.unlink()
            logging.info(f"Arquivo original removido: {ORIGINAL_FILE}")
        except Exception as e:
            logging.error(f"Erro ao remover arquivo original: {e}")
        
        return True
    
    except Exception as e:
        logging.error(f"Falha no tratamento de dados: {e}")
        return False

def tratamento_e_merge():
    """Processa, combina os dados e calcula dias de atraso."""
    try:
        # Verificar existência dos arquivos
        for file in [FILTERED_FILE, IP03_FILE]:
            if not file.exists():
                raise FileNotFoundError(f"Arquivo não encontrado: {file}")
        
        # Ler dados
        df_ih08 = pd.read_excel(FILTERED_FILE)
        df_ip03 = pd.read_excel(IP03_FILE)
        
        # Passo 1: Extrair os 10 primeiros caracteres do 'Plano manut.' no IP03
        df_ip03['Equipamento_Base'] = df_ip03['Plano manut.'].str[:10]
        
        # Passo 2: Merge usando a coluna base
        df_merged = df_ih08.merge(
            df_ip03,
            how='left',
            left_on='Equipamento',          # Coluna original do IH08
            right_on='Equipamento_Base',    # Coluna derivada do IP03
            suffixes=('_ih08', '_ip03'),
            indicator=True
        )
        
        # Passo 3: Remover coluna auxiliar (opcional)
        df_merged.drop('Equipamento_Base', axis=1, inplace=True, errors='ignore')
        
        # Atualizar datas e calcular atraso
        df_merged['Dt.criação'] = df_merged['Modificado em_ip03'].combine_first(df_merged['Dt.criação'])
        hoje = pd.Timestamp.now().normalize()
        df_merged['Dias_atras'] = (hoje - df_merged['Dt.criação']).dt.days
        
        # Salvar arquivo merged
        df_merged.to_excel(MERGED_FILE, index=False)
        logging.info(f"Arquivo de merge salvo em: {MERGED_FILE}")
        
        return True
    
    except Exception as e:
        logging.error(f"Erro durante o merge: {e}")
        return False

def gerar_tabela_sem_plano():
    """Gera tabela de equipamentos sem plano de manutenção."""
    try:
        if not MERGED_FILE.exists():
            raise FileNotFoundError(f"Arquivo merged não encontrado: {MERGED_FILE}")
        
        df_merged = pd.read_excel(MERGED_FILE)
        
        # Identificar equipamentos sem correspondência no IP03
        equipamentos_sem_plano = df_merged[df_merged['_merge'] == 'left_only']
        
        if not equipamentos_sem_plano.empty:
            # Selecionar apenas colunas relevantes
            cols = ['Equipamento', 'Denominação', 'Dt.criação', 'Status sistema']
            equipamentos_sem_plano[cols].to_excel(SEM_PLANO_FILE, index=False)
            logging.info(f"Arquivo com equipamentos sem plano salvo em: {SEM_PLANO_FILE}")
            return True
        else:
            logging.info("Todos os equipamentos possuem plano de manutenção")
            return True
            
    except Exception as e:
        logging.error(f"Erro ao gerar tabela sem plano: {e}")
        return False

if __name__ == "__main__":
    logging.info("Iniciando processamento de dados...")
    
    if tratamento_dados():
        logging.info("Processamento inicial concluído. Iniciando merge...")
        if tratamento_e_merge():
            logging.info("Merge concluído com sucesso. Buscando equipamentos sem plano...")
            if gerar_tabela_sem_plano():
                logging.info("Processamento de equipamentos sem plano concluído.")
        else:
            logging.error("Falha durante o merge de dados.")
    else:
        logging.error("Processamento inicial falhou. Operação abortada.")