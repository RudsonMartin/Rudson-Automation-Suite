# processo_sap.py
import pandas as pd
import subprocess
import logging
from pathlib import Path

# Configurações de caminho
BASE_DIR = Path(r"C:\Users\rmbotelho\Documents\Plano PM v2")
SEM_PLANO_FILE = BASE_DIR / "equipamentos_sem_plano.xlsx"
VBS_SCRIPT_PATH = BASE_DIR / "vbs" / "FLAGRAFAEL.vbs"
MAMOMETRO_VBS = BASE_DIR / "vbs" / "manometro-teste.vbs"
FILTRO_VBS = BASE_DIR / "vbs" / "filtroteste.vbs"

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(BASE_DIR / "processamento.log"),
        logging.StreamHandler()
    ]
)

def limparequipamentto():
    try:
        arquivo_para_excluir =[
            SEM_PLANO_FILE
        ]
        for arquivo in arquivo_para_excluir:
         if arquivo.exists():
             arquivo.unlink()
             logging.info(f"Arquivo excluído: {arquivo.name}")
                
        return True
        
    except Exception as e:
        logging.error(f"Falha ao excluir arquivo: {str(e)}")
        return False

def processar_equipamentos():  # Nome corrigido
    """Executa scripts SAP conforme regras de negócio"""
    try:
        df = pd.read_excel(SEM_PLANO_FILE)
        equipamentos = df['Equipamento'].tolist()

        for equip in equipamentos:
            try:  # Tratamento de erro por equipamento
                if equip.startswith('PMBRA'):
                    saniDias = "30"
                    subprocess.run(f'cscript.exe "{VBS_SCRIPT_PATH}" "{equip}" "{saniDias}"', shell=True, check=True)
                    subprocess.run(f'cscript.exe "{MAMOMETRO_VBS}" "{equip}"', shell=True, check=True)
                    subprocess.run(f'cscript.exe "{FILTRO_VBS}" "{equip}"', shell=True, check=True)
                elif equip.startswith('CHBRA'):
                    saniDias = "20"
                    subprocess.run(f'cscript.exe "{VBS_SCRIPT_PATH}" "{equip}" "{saniDias}"', shell=True, check=True)
                    subprocess.run(f'cscript.exe "{MAMOMETRO_VBS}" "{equip}"', shell=True, check=True)
                elif equip.startswith('SNBRA'):
                    saniDias = "7"
                    subprocess.run(f'cscript.exe "{VBS_SCRIPT_PATH}" "{equip}" "{saniDias}"', shell=True, check=True)
                    subprocess.run(f'cscript.exe "{FILTRO_VBS}" "{equip}"', shell=True, check=True)
                else:
                    logging.warning(f"Equipamento {equip} não possui regra definida.")
                
            except subprocess.CalledProcessError as e:
                logging.error(f"Falha no equipamento {equip}: {str(e)}")
                continue  # Pula para o próximo equipamento

        return True

    except Exception as e:
        logging.error(f"Erro geral no processamento SAP: {str(e)}", exc_info=True)
        return False

def main():
    logging.info("Início do programa.")
    processar_equipamentos()
    limparequipamentto()
    logging.info("Fim do programa.")

if __name__ == "__main__":
    main()

