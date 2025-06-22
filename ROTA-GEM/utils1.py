import pandas as pd
import os
from datetime import date, timedelta
from openpyxl import load_workbook


def getData(path):
    CIDADES_DIAS = {
        "SEGUNDA": ["CEILANDIA", "SAMAMBAIA", "BRAZLANDIA", "AGUAS LINDAS DE GOIAS"],
        "TERÇA": [
            "OCTOGONAL", "SUDOESTE", "ASA SUL", "SIA", "SAAN", "ZONA CIVICO-ADMINISTRATIVA",
            "SETOR DE GARAGENS OFICIAIS", "ZONA INDUSTRIAL", "LAGO SUL", "JARDIM BOTANICO",
            "VILA PLANALTO", "SETOR MILITAR URBANO", "SIG", "CRUZEIRO NOVO", "SETOR POLICIAL",
            "AEROPORTO", "VILA DA TELEBRASILIA", "SAO SEBASTIAO", "SAO SEBASTIÃO", "SOFN"
        ],
        "QUARTA": [
            "GAMA", "NOVO GAMA", "SANTA MARIA", "RIACHO FUNDO II", "RECANTO DAS EMAS",
            "ALEXANIA", "VALPARAISO DE GOIAS", "ABADIANIA", "SANTO ANTONIO DE GOIAS",
            "SANTO ANTONIO DO DESCOBERTO", "LUZIANIA"
        ],
        "QUINTA": [
            "ASA NORTE", "PLANALTINA", "SOBRADINHO", "LAGO NORTE", "PARANOA", "ITAPOA",
            "NOROESTE", "LAGO OESTE", "UNB DARCY RIBEIRO", "SAM", "VILA PLANALTO", "SOFN",
            "CAFE SEM TROCO", "TAQUARI", "FERCAL", "GRANJA DO TORTO", "SETOR DE GARAGENS OFICIAIS", "VARJÃO"
        ],
        "SEXTA": ["TAGUATINGA", "AGUAS CLARAS", "VICENTE PIRES", "26 DE SETEMBRO", "PARK WAY", "RIACHO FUNDO I"],
        "SABADO": ["GUARA", "GUARA II", "NUCLEO BANDEIRANTE", "CRUZEIRO", "CANDANGOLANDIA", "SETOR HIPICO", "ESTRUTURAL"],
    }

    def get_dia(cidade):
        for dia in CIDADES_DIAS:
            if cidade in CIDADES_DIAS[dia]:
                return dia
        return "NENHUM"

    df = pd.read_excel(path)
    df["ROTA"] = [get_dia(cidade) for cidade in df["Cidade"]]
    df.loc[df["ROTA"] == "NENHUM", "ROTA"] = [
        get_dia(bairro) for bairro in df.loc[df["ROTA"] == "NENHUM", "Bairro"]
    ]
    df.loc[df["Centro"] != 1, "ROTA"] = "NENHUM"

    df = df.fillna("")

    df_desc = df[df["Descrição Tipo Solic."].isin(["Instalação", "Recolhimento", "Troca Comercial", "Troca técnica"])]
    df_status = df_desc[df_desc["Status"].isin(["FEC", "BLQ"])]
    df_status_entrega = df_status[df_status["Status da Entrega"].isin(["", "REP_TECNIC"])]
    df_deps = df_status_entrega[~df_status_entrega["Depós."].isin(["GAM", "RMK", "DVT"])]

    df_install = df_deps[df_deps["Descrição Tipo Solic."].isin(["Instalação", "Troca Comercial"])]
    df_install_3 = df_install[df_install["Niv.Apr.Atual"] == 3]
    df_install_2 = df_install[df_install["Niv.Apr.Atual"] == 2]

    df_recolhimento = df_deps[df_deps["Descrição Tipo Solic."].isin(["Recolhimento"])]
    df_recolhimento_3 = df_recolhimento[df_recolhimento["Niv.Apr.Atual"] == 3]
    df_recolhimento_2 = df_recolhimento[df_recolhimento["Niv.Apr.Atual"] == 2]
    df_recolhimento_1 = df_recolhimento[df_recolhimento["Niv.Apr.Atual"] == 1]

    df_troca = df_deps[df_deps["Descrição Tipo Solic."].isin(["Troca técnica"])]
    df_troca_1 = df_troca[df_troca["Niv.Apr.Atual"] == 1]
    df_troca_2 = df_troca[df_troca["Niv.Apr.Atual"] == 3]

    # 👇 incluir tudo de instalação (2 e 3) e recolhimento (1, 2 e 3)
    df_final = pd.concat([
        df_install_2, df_install_3,
        df_recolhimento_1, df_recolhimento_2, df_recolhimento_3,
        df_troca_1, df_troca_2
    ])

    df_final.rename(columns={
        "Aprovador 3": "aprovador3",
        "Bairro": "Bairro",
        "Cidade": "Cidade",
        "Cód.Cliente": "Cliente",
        "Desc.Canal": "descCanal",
        "Desc.Sub Canal": "descSubCanal",
        "Descrição Tipo Solic.": "DescTipoSolic",
        "Dt.Aprovação 1": "dtAprovacao1",
        "Dt.Aprovação 2": "dtAprovacao2",
        "Dt.Aprovação 3": "dtAprovacao3",
        "Dt.Criação": "dtCriacao",
        "Endereço": "Endereco",
        "Niv.Apr.Atual": "nivAprAtual",
        "Nº Solicitação": "Solicitacao",
        "Nome Fantasia": "NomeFantasia",
        "ROTA": "Rota",
        "SLA": "sla",
        "Status": "Status",
        "Status da Entrega": "statusEntrega",
        "Desc.Coord.": "DescCoord",
    }, inplace=True)

    # 🟢 Print de debug no lugar certo, dentro da função
    print("Total geral:", len(df))
    print("Com Descrição Tipo:", len(df_desc))
    print("Com Status FEC ou BLQ:", len(df_status))
    print("Com Status entrega '' ou REP_TECNIC:", len(df_status_entrega))
    print("Sem depósito GAM, RMK, DVT:", len(df_deps))
    print("Instalação N2:", len(df_deps[(df_deps["Descrição Tipo Solic."] == "Instalação") & (df_deps["Niv.Apr.Atual"] == 2)]))
    print("Recolhimento N1:", len(df_deps[(df_deps["Descrição Tipo Solic."] == "Recolhimento") & (df_deps["Niv.Apr.Atual"] == 1)]))

    df_final["Cliente"] = df_final["Cliente"].astype("Int64").astype(str)

    df_final = df_final[[
        "Solicitacao", "Cliente", "NomeFantasia", "DescCoord", "Rota",
        "Depós.", "DescTipoSolic", "Bairro", "Cidade", "Nº Equip.",
        "Nº.Equip.Instalar", "Texto Breve Material", "Endereco", "Status",
        "Centro", "nivAprAtual"
    ]]

    return df_final, df_install_2, df_recolhimento_1, df_recolhimento_2, df_recolhimento_3
    




def updateSolicitacoes():
    data_hoje = date.today()
    data_inicial = data_hoje - timedelta(days=180)
    
    vbs_script = r"C:\Users\rudso\Downloads\ROTA_DASH (1)\ROTA_DASH\listar_solicitacoes_c.vbs"
    
    param1 = data_inicial.strftime('%d%m%Y')
    param2 = data_hoje.strftime('%d%m%Y')

    comando = f'cscript "{vbs_script}" {param1} {param2}'

    os.system(comando)
