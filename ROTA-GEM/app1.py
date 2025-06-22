import dash_bootstrap_components as dbc
import dash
from dash import Dash, dcc, html, Input, Output, dash_table, callback, State, ctx
import pandas as pd
from utils1 import getData, updateSolicitacoes
import sys
import os
from io import BytesIO

df_solicitacoes, df_install_2, df_recolhimento_1,df_recolhimento_2,df_recolhimento_3 = getData(r"C:\Users\rudso\Downloads\ROTA_DASH (1)\ListarSolic20251.XLSX")

deleted_rows = []
df_deleted = pd.DataFrame()

df_solicitacoes_C1 = pd.DataFrame()
df_solicitacoes_C2 = pd.DataFrame()
df_solicitacoes_C3 = pd.DataFrame()
show_table = True

table_styles = dict(
    style_table={
        'maxHeight': '500px',
        'overflowY': 'auto',
        'borderRadius': '8px'
    },
    style_header={
        'backgroundColor': '#2c3e50',
        'color': 'white'
    },
    style_cell={
        'padding': '10px',
        'minWidth': '120px',
        'textAlign': 'left',
        'border': 'none'
    },
    style_data_conditional=
[
    {
        'if': {'column_id': 'nivAprAtual', 'filter_query': '{nivAprAtual} = 1'},
        'backgroundColor': '#e74c3c',  # vermelho
        'color': 'white'
    },
    {
        'if': {'column_id': 'nivAprAtual', 'filter_query': '{nivAprAtual} = 2'},
        'backgroundColor': '#f1c40f',  # amarelo
        'color': 'black'
    },
    {
        'if': {'column_id': 'nivAprAtual', 'filter_query': '{nivAprAtual} = 3'},
        'backgroundColor': '#2ecc71', 
        'color': 'white'
    }
]
)


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


app = Dash(
    __name__,
    assets_folder=resource_path("assets"),
    external_stylesheets=[dbc.themes.SIMPLEX],
)

app.layout = html.Div(
    [
        html.Div(
            className="app-header",
            children=[
                html.H1("ROTA - GEM", className="app-header-title"),
                html.Hr(),
                html.Div(
                    className="container",
                    children=[
                        html.Div(
                            id="botao-altera-tela",
                            children=[
                                dbc.Button("Caminhões", id="separa-rota", n_clicks=0),
                            ],
                        ),
                        dcc.Download(id="download-dataframe-xlsx"),
                        html.Hr(),
                        html.Div(
                            id="tab-solicitacoes-div",
                            children=[
                                html.Div(
                                    className="app-header-content",
                                    children=[
                                        dbc.Button(
                                            "Atualizar Solicitações",
                                            id="atualiza-solicitacoes",
                                            n_clicks=0,
                                        ),
                                        dcc.Checklist(
                                            [
                                                "SEGUNDA",
                                                "TERÇA",
                                                "QUARTA",
                                                "QUINTA",
                                                "SEXTA",
                                                "SABADO",
                                                "NENHUM",
                                            ],
                                            inline=True,
                                            id="filtro-rota-dia",
                                        ),
                                        html.Div(id="quantidade-filtrada"),
                                        html.Div(id="update-solicitacoes"),
                                    ],
                                ),
                                html.Div(id="separar-rota-div"),
                                html.Hr(),
                                html.Div(
                                    id="buttons-caminhao-q-nivel-financeiro-div",
                                    children=[
                                        html.Div(
                                            id="buttons-caminhao-div",
                                            children=[
                                                dbc.Button(
                                                    "Mover C1",
                                                    id="move-C1",
                                                    className="buttons",
                                                    n_clicks=0,
                                                ),
                                                dbc.Button(
                                                    "Mover C2",
                                                    id="move-C2",
                                                    className="buttons",
                                                    n_clicks=0,
                                                ),
                                                dbc.Button(
                                                    "Mover C3",
                                                    id="move-C3",
                                                    className="buttons",
                                                    n_clicks=0,
                                                ),
                                            ],
                                        ),
                                        html.Div(id="instalacao-div"),
                                        html.Div(id="recolhimento-div"),
                                        html.Div(
                                            id="colapsse-area",
                                            children=[
                                                dbc.Button(
                                                    "Mostrar",
                                                    id="collapse-button",
                                                    className="mb-3",
                                                    color="primary",
                                                    n_clicks=0,
                                                ),
                                                dbc.Collapse(
                                                    html.Div(id="collapse-content"),
                                                    id="collapse",
                                                    is_open=False,
                                                ),
                                            ],
                                        ),
                                    ],
                                ),
                                html.Hr(),
                                html.Div(
                                    id="tabela-solicitacoes-div",
                                    children=[
                                        dash_table.DataTable(
                                            id="tabela-solicitacoes",
                                            columns=[                                           
                                            {"name": "Solicitação", "id": "Solicitacao"},
                                            {"name": "Nível", "id": "nivAprAtual", "presentation": "markdown"},
                                            {"name": "DescTipoSolic", "id": "DescTipoSolic"},
                                            {"name": "Depós.", "id": "Depós."},
                                            {"name": "Cidade", "id": "Cidade" },
                                            {"name": "Bairro","id":"Bairro" },
                                            {"name": "Endereco", "id": "Endereco"},
                                            {"name": "NomeFantasia", "id": "NomeFantasia"},
                                            {"name": "Nº Equip.", "id": "Nº Equip."},
                                            {"name": "Nº.Equip.Instalar", "id": "Nº.Equip.Instalar"},
                                            {"name": "Texto Breve Material", "id": "Texto Breve Material"}
                                            ],

                                            data=df_solicitacoes.to_dict("records"),
                                            filter_action="native",
                                            sort_action="native",
                                            sort_mode="multi",
                                            row_deletable=True,
                                            editable=True,
                                            row_selectable="multi",
                                            selected_rows=[],
                                            filter_options={
                                                "placeholder_text": " ",
                                                "case": "insensitive",
                                            },
                                            **table_styles,
                                            derived_virtual_data=[],
                                            # export_format="xlsx",
                                            # export_headers="display",
                                        ),
                                    ],
                                ),
                                html.Hr(),
                                dash_table.DataTable(
                                    id="tabela-solicitacoes-deletadas",
                                    columns=[
                                    {"name": "Solicitação", "id": "Solicitacao"},
                                    {"name": "Nível", "id": "nivAprAtual", "presentation": "markdown"},
                                    {"name": "DescTipoSolic", "id": "DescTipoSolic"},
                                    {"name": "Depós.", "id": "Depós."},
                                    {"name": "Endereco", "id": "Endereco"},
                                    {"name": "Cidade", "id": "Cidade" },
                                    {"name": "Bairro","id":"Bairro" },
                                    {"name": "NomeFantasia", "id": "NomeFantasia"},
                                    {"name": "Nº Equip.", "id": "Nº Equip."},
                                    {"name": "Nº.Equip.Instalar", "id": "Nº.Equip.Instalar"},
                                    {"name": "Texto Breve Material", "id": "Texto Breve Material"}
                                    ],

                                    data=df_deleted.to_dict("records"),
                                    filter_action="native",
                                    sort_action="native",
                                    sort_mode="multi",
                                    row_deletable=True,
                                    editable=True,
                                    filter_options={
                                        "placeholder_text": " ",
                                        "case": "insensitive",
                                    },
                                    **table_styles,
                                    derived_virtual_data=[],
                                    # export_format="xlsx",
                                    # export_headers="display",
                                ),
                            ],
                        ),
                        html.Div(
                            id="tabs-caminhoes-div",
                            style={"display": "none"},
                            children=[
                                html.Div(
                                    [
                                        dbc.Button(
                                            "Exportar", id="Imprimir", n_clicks=0
                                        ),
                                        dcc.Tabs(
                                            id="tabs-example-graph",
                                            value="tab-C1",
                                            children=[
                                                dcc.Tab(
                                                    label="Caminhão 1",
                                                    value="tab-C1",
                                                    children=[
                                                        html.Div(
                                                            className="tab-caminhao",
                                                            children=[
                                                                html.Hr(),
                                                                html.Div(
                                                                    className="tab-caminhao-div-container",
                                                                    children=[
                                                                        html.Div(
                                                                            className="tab-caminhao-div-buttons",
                                                                            children=[
                                                                                dbc.Button(
                                                                                    "Selecionar tudo",
                                                                                    id="C1-select-all",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                                dbc.Button(
                                                                                    "Mover para Caminhão 2",
                                                                                    id="C1-move-C2",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                                dbc.Button(
                                                                                    "Mover para Caminhão 3",
                                                                                    id="C1-move-C3",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                            ],
                                                                        ),
                                                                        html.Div(
                                                                            className="tab-caminhao-div-quantidade",
                                                                            id="tab-caminhao-quantidade-c1",
                                                                        ),
                                                                    ],
                                                                ),
                                                                html.Hr(),
                                                                dash_table.DataTable(
                                                                    id="tabela-solicitacoes-c-1",
                                                                    columns=[
                                                                    {"name": "Solicitação", "id": "Solicitacao"},
                                                                    {"name": "Nível", "id": "nivAprAtual", "presentation": "markdown"},
                                                                    {"name": "DescTipoSolic", "id": "DescTipoSolic"},
                                                                    {"name": "Depós.", "id": "Depós."},
                                                                    {"name": "Endereco", "id": "Endereco"},
                                                                    {"name": "Cidade", "id": "Cidade" },
                                                                    {"name": "Bairro","id":"Bairro" },             
                                                                    {"name": "NomeFantasia", "id": "NomeFantasia"},
                                                                    {"name": "Nº Equip.", "id": "Nº Equip."},
                                                                    {"name": "Nº.Equip.Instalar", "id": "Nº.Equip.Instalar"},
                                                                    {"name": "Texto Breve Material", "id": "Texto Breve Material"}
                                                                    ],
                                                                    data=df_solicitacoes_C1.to_dict(
                                                                        "records"
                                                                    ),
                                                                    filter_action="native",
                                                                    sort_action="native",
                                                                    sort_mode="multi",
                                                                    row_deletable=True,
                                                                    editable=True,
                                                                    row_selectable="multi",
                                                                    selected_rows=[],
                                                                    filter_options={
                                                                        "placeholder_text": " ",
                                                                        "case": "insensitive",
                                                                    },
                                                                    **table_styles,
                                                                    derived_virtual_data=[],
                                                                ),
                                                            ],
                                                        )
                                                    ],
                                                ),
                                                dcc.Tab(
                                                    label="Caminhão 2",
                                                    value="tab-C2",
                                                    children=[
                                                        html.Div(
                                                            className="tab-caminhao",
                                                            children=[
                                                                html.Hr(),
                                                                html.Div(
                                                                    className="tab-caminhao-div-container",
                                                                    children=[
                                                                        html.Div(
                                                                            className="tab-caminhao-div-buttons",
                                                                            children=[
                                                                                dbc.Button(
                                                                                    "Selecionar tudo",
                                                                                    id="C2-select-all",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                                dbc.Button(
                                                                                    "Mover para Caminhão 1",
                                                                                    id="C2-move-C1",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                                dbc.Button(
                                                                                    "Mover para Caminhão 3",
                                                                                    id="C2-move-C3",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                            ],
                                                                        ),
                                                                        html.Div(
                                                                            className="tab-caminhao-div-quantidade",
                                                                            id="tab-caminhao-quantidade-c2",
                                                                        ),
                                                                    ],
                                                                ),
                                                                html.Hr(),
                                                                dash_table.DataTable(
                                                                    id="tabela-solicitacoes-c-2",
                                                                    columns=[
                                                                    {"name": "Solicitação", "id": "Solicitacao"},
                                                                    {"name": "Nível", "id": "nivAprAtual", "presentation": "markdown"},
                                                                    {"name": "DescTipoSolic", "id": "DescTipoSolic"},
                                                                    {"name": "Depós.", "id": "Depós."},
                                                                    {"name": "Cidade", "id": "Cidade"},
                                                                    {"name": "Bairro", "id": "Bairro"},
                                                                    {"name": "Endereco", "id": "Endereco"},
                                                                    {"name": "NomeFantasia", "id": "NomeFantasia"},
                                                                    {"name": "Nº Equip.", "id": "Nº Equip."},
                                                                    {"name": "Nº.Equip.Instalar", "id": "Nº.Equip.Instalar"},
                                                                    {"name": "Texto Breve Material", "id": "Texto Breve Material"},
                                                                 ], 
                                                                    data=df_solicitacoes_C2.to_dict(
                                                                        "records"
                                                                    ),
                                                                    filter_action="native",
                                                                    sort_action="native",
                                                                    sort_mode="multi",
                                                                    row_deletable=True,
                                                                    editable=True,
                                                                    row_selectable="multi",
                                                                    selected_rows=[],
                                                                    filter_options={
                                                                        "placeholder_text": " ",
                                                                        "case": "insensitive",
                                                                    },
                                                                    **table_styles,
                                                                    derived_virtual_data=[],
                                                                ),
                                                            ],
                                                        )
                                                    ],
                                                ),
                                                dcc.Tab(
                                                    label="Caminhão 3",
                                                    value="tab-C3",
                                                    children=[
                                                        html.Div(
                                                            className="tab-caminhao",
                                                            children=[
                                                                html.Hr(),
                                                                html.Div(
                                                                    className="tab-caminhao-div-container",
                                                                    children=[
                                                                        html.Div(
                                                                            className="tab-caminhao-div-buttons",
                                                                            children=[
                                                                                dbc.Button(
                                                                                    "Selecionar tudo",
                                                                                    id="C3-select-all",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                                dbc.Button(
                                                                                    "Mover para Caminhão 1",
                                                                                    id="C3-move-C1",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                                dbc.Button(
                                                                                    "Mover para Caminhão 2",
                                                                                    id="C3-move-C2",
                                                                                    className="buttons",
                                                                                    n_clicks=0,
                                                                                ),
                                                                            ],
                                                                        ),
                                                                        html.Div(
                                                                            className="tab-caminhao-div-quantidade",
                                                                            id="tab-caminhao-quantidade-c3",
                                                                        ),
                                                                    ],
                                                                ),
                                                                html.Hr(),
                                                                dash_table.DataTable(
                                                                    id="tabela-solicitacoes-c-3",
                                                                    columns=[
                                                                    {"name": "Solicitação", "id": "Solicitacao"},
                                                                    {"name": "Nível", "id": "nivAprAtual", "presentation": "markdown"},
                                                                    {"name": "DescTipoSolic", "id": "DescTipoSolic"},
                                                                    {"name": "Depós.", "id": "Depós."},
                                                                    {"name": "Endereco", "id": "Endereco"},
                                                                    {"name": "Cidade", "id": "Cidade"},
                                                                    {"name": "Bairro", "id": "Bairro"},
                                                                    {"name": "NomeFantasia", "id": "NomeFantasia"},
                                                                    {"name": "Nº Equip.", "id": "Nº Equip."},
                                                                    {"name": "Nº.Equip.Instalar", "id": "Nº.Equip.Instalar"},
                                                                    {"name": "Texto Breve Material", "id": "Texto Breve Material"},
                                                                    ],
                                                                    data=df_solicitacoes_C3.to_dict(
                                                                        "records"
                                                                    ),
                                                                    filter_action="native",
                                                                    sort_action="native",
                                                                    sort_mode="multi",
                                                                    row_deletable=True,
                                                                    editable=True,
                                                                    row_selectable="multi",
                                                                    selected_rows=[],
                                                                    filter_options={
                                                                        "placeholder_text": " ",
                                                                        "case": "insensitive",
                                                                    },
                                                                    **table_styles,
                                                                    derived_virtual_data=[],
                                                                ),
                                                            ],
                                                        )
                                                    ],
                                                ),
                                            ],
                                        ),
                                        html.Div(id="tabs-content-example-graph"),
                                    ]
                                )
                            ],
                        ),
                    ],
                ),
            ],
        ),
        dash.page_container,
    ]
)


# Filtra solicitacoes por dia da semana
@callback(
    Output("tabela-solicitacoes", "filter_query"),
    Output("instalacao-div", "children"),
    Output("recolhimento-div", "children"),
    Output("collapse-content", "children"),
    Input("filtro-rota-dia", "value"),
)
def filtrar_por_dia(dia):
    # print(dia)
    query = "{Rota} eq "
    add_query = " or {Rota} eq "
    return_query = ""
    if dia:
        if len(dia) > 0:
            qtd_install = len(df_install_2[df_install_2["ROTA"].isin(dia)])
            qtd_recolhimento = len(
                df_recolhimento_1[df_recolhimento_1["ROTA"].isin(dia)]
            )
            qtd_install2 = len(df_install_2[df_install_2["ROTA"].isin(dia)])
            qtd_recolhimento2 = len(
                df_recolhimento_2[df_recolhimento_2["ROTA"].isin(dia)]
            )
            solicitacoes_install = df_install_2[df_install_2["ROTA"].isin(dia)][
                "Nº Solicitação"
            ].to_list()
            solicitacoes_recolhimento = df_recolhimento_1[
                df_recolhimento_1["ROTA"].isin(dia)
            ]["Nº Solicitação"].to_list()
            solicitacoes = solicitacoes_install + solicitacoes_recolhimento
            print(solicitacoes)
            return_divs = []
            for solicitacao in solicitacoes:
                return_divs.append(
                    dcc.Markdown(
                        "##### {}".format(solicitacao),
                        style={"margin-top": "10px"},
                    )
                )
            # print(df_install_2[df_install_2["ROTA"].isin(dia)])
            # print(df_recolhimento_1[df_recolhimento_1["ROTA"].isin(dia)])
            if len(dia) == 1:
                return (
                    query + dia[0],
                    dcc.Markdown("##### Instalação Nivel 2 = {}".format(qtd_install)),
                    dcc.Markdown(
                        "##### Recolhimento Nivel 1 = {}".format(qtd_recolhimento)
                    ),
                    return_divs,
                )
            else:
                return_query = query + dia[0]
                for i in range(len(dia) - 1):
                    return_query += add_query + dia[i + 1]
                # print(return_query)
                return (
                    return_query,
                    dcc.Markdown("##### Instalação Nivel 2 = {}".format(qtd_install)),
                    dcc.Markdown(
                        "##### Recolhimento Nivel 1 = {}".format(qtd_recolhimento)
                    ),
                    return_divs,
                )
        else:
            return (
                "",
                dcc.Markdown("##### Instalação Nivel 2 = {}".format(qtd_install)),
                dcc.Markdown(
                    "##### Recolhimento Nivel 1 = {}".format(qtd_recolhimento)
                ),
                "",
            )
    return (
        "",
        dcc.Markdown("##### Instalação Nivel 2 = {}".format(0)),
        dcc.Markdown("##### Recolhimento Nivel 1 = {}".format(0)),
        "",
    )


# Atualiza solicitacoes / Roda script do SAP
@callback(
    Output("tabela-solicitacoes", "data", allow_duplicate=True),
    Input("atualiza-solicitacoes", "n_clicks"),
    prevent_initial_call=True,
)
def atualiza_solicitacoes(n_clicks):
    global deleted_rows, df_deleted, df_solicitacoes, df_install_2, df_recolhimento_1
    deleted_rows = []
    df_deleted = pd.DataFrame()
    if n_clicks > 0:
        updateSolicitacoes()
        df_solicitacoes, df_install_2, df_recolhimento_1,df_recolhimento_2,df_recolhimento_3= getData(
            r"C:\Users\rudso\Downloads\ROTA_DASH (1)\ListarSolic20251.XLSX"
        )
        return df_solicitacoes.to_dict("records")
    return df_solicitacoes.to_dict("records")


@callback(
    Output("tab-solicitacoes-div", "style"),
    Output("tabs-caminhoes-div", "style"),
    Output("separa-rota", "children"),
    Input("separa-rota", "n_clicks"),
)
def separar_rota(n_clicks):
    global show_table
    if n_clicks > 0:
        show_table = False if show_table else True
        if show_table:
            return (
                {"display": "block"},
                {"display": "none"},
                "Caminhões",
            )
        else:
            return (
                {"display": "none"},
                {"display": "block"},
                "Solicitações",
            )
    return (
        {"display": "block"},
        {"display": "none"},
        "Caminhões",
    )


# Atualiza quantidade de solicitacoes filtradas e tabela de solicitacoes deletadas
@callback(
    Output("quantidade-filtrada", "children"),
    Output("tabela-solicitacoes-deletadas", "data"),
    Output("tabela-solicitacoes", "data", allow_duplicate=True),
    Input("tabela-solicitacoes", "filter_query"),
    Input("tabela-solicitacoes", "data_previous"),
    State("tabela-solicitacoes", "derived_virtual_data"),
    State("tabela-solicitacoes", "data"),
    Input("tabela-solicitacoes-deletadas", "data_previous"),
    State("tabela-solicitacoes-deletadas", "data"),
    prevent_initial_call=True,
)
def update_rows_value(
    value,
    data_previous,
    derived_virtual_data,
    data,
    data_deleted_previous,
    data_deleted,
):
    global deleted_rows, df_deleted
    triggered_id = ctx.triggered_id
    df_deleted = pd.DataFrame.from_dict(deleted_rows)
    data_copy = data

    if value == "":
        qtd = len(data)
    else:
        qtd = len(derived_virtual_data)

    if triggered_id == "tabela-solicitacoes":
        if data_previous is not None:
            deleted_indices = set([row["Solicitacao"] for row in data_previous]) - set(
                [row["Solicitacao"] for row in data]
            )
            for row in derived_virtual_data:
                if row["Solicitacao"] in deleted_indices:
                    deleted_rows.append(row)
                    qtd -= 1
    elif triggered_id == "tabela-solicitacoes-deletadas":
        if data_deleted_previous is not None:
            deleted_indices = set(
                [row["Solicitacao"] for row in data_deleted_previous]
            ) - set([row["Solicitacao"] for row in data_deleted])
            for row in data_deleted_previous:
                if row["Solicitacao"] in deleted_indices:
                    deleted_rows.remove(row)
                    qtd += 1
                    data_copy.append(row)

    df_deleted = pd.DataFrame.from_dict(deleted_rows)
    return (
        dcc.Markdown("#### Quantidade = {}".format(qtd)),
        df_deleted.to_dict("records"),
        pd.DataFrame.from_dict(data_copy).to_dict("records"),
    )


@callback(
    Output("tabela-solicitacoes-c-1", "data", allow_duplicate=True),
    Output("tabela-solicitacoes-c-2", "data"),
    Output("tabela-solicitacoes-c-3", "data"),
    Output("tabela-solicitacoes-c-1", "selected_rows"),
    Output("tabela-solicitacoes-c-2", "selected_rows"),
    Output("tabela-solicitacoes-c-3", "selected_rows"),
    Output("tab-caminhao-quantidade-c1", "children", allow_duplicate=True),
    Output("tab-caminhao-quantidade-c2", "children", allow_duplicate=True),
    Output("tab-caminhao-quantidade-c3", "children", allow_duplicate=True),
    Output("tabela-solicitacoes-c-1", "filter_query"),
    Output("tabela-solicitacoes-c-2", "filter_query"),
    Output("tabela-solicitacoes-c-3", "filter_query"),
    Input("C1-move-C2", "n_clicks"),
    Input("C1-move-C3", "n_clicks"),
    Input("C2-move-C1", "n_clicks"),
    Input("C2-move-C3", "n_clicks"),
    Input("C3-move-C1", "n_clicks"),
    Input("C3-move-C2", "n_clicks"),
    State("tabela-solicitacoes-c-1", "data"),
    State("tabela-solicitacoes-c-2", "data"),
    State("tabela-solicitacoes-c-3", "data"),
    State("tabela-solicitacoes-c-1", "selected_rows"),
    State("tabela-solicitacoes-c-2", "selected_rows"),
    State("tabela-solicitacoes-c-3", "selected_rows"),
    prevent_initial_call=True,
)
def mover_solicitacoes_entre_caminhoes(
    n_clicks_1,
    n_clicks_2,
    n_clicks_3,
    n_clicks_4,
    n_clicks_5,
    n_clicks_6,
    df_C1,
    df_C2,
    df_C3,
    selected_rows_C1,
    selected_rows_C2,
    selected_rows_C3,
):
    button_clicked = ctx.triggered_id
    df_C1_new = df_C1
    df_C2_new = df_C2
    df_C3_new = df_C3

    if button_clicked == "C1-move-C2":
        rows_to_move = [df_C1_new[i] for i in selected_rows_C1]
        df_C1_new = [
            df_C1_new[i] for i in range(len(df_C1_new)) if i not in selected_rows_C1
        ]
        for row in rows_to_move:
            df_C2_new.append(row)

    elif button_clicked == "C1-move-C3":
        rows_to_move = [df_C1_new[i] for i in selected_rows_C1]
        df_C1_new = [
            df_C1_new[i] for i in range(len(df_C1_new)) if i not in selected_rows_C1
        ]
        for row in rows_to_move:
            df_C3_new.append(row)
    elif button_clicked == "C2-move-C1":
        rows_to_move = [df_C2_new[i] for i in selected_rows_C2]
        df_C2_new = [
            df_C2_new[i] for i in range(len(df_C2_new)) if i not in selected_rows_C2
        ]
        for row in rows_to_move:
            df_C1_new.append(row)
    elif button_clicked == "C2-move-C3":
        rows_to_move = [df_C2_new[i] for i in selected_rows_C2]
        df_C2_new = [
            df_C2_new[i] for i in range(len(df_C2_new)) if i not in selected_rows_C2
        ]
        for row in rows_to_move:
            df_C3_new.append(row)
    elif button_clicked == "C3-move-C1":
        rows_to_move = [df_C3_new[i] for i in selected_rows_C3]
        df_C3_new = [
            df_C3_new[i] for i in range(len(df_C3_new)) if i not in selected_rows_C3
        ]
        for row in rows_to_move:
            df_C1_new.append(row)
    elif button_clicked == "C3-move-C2":
        rows_to_move = [df_C3_new[i] for i in selected_rows_C3]
        df_C3_new = [
            df_C3_new[i] for i in range(len(df_C3_new)) if i not in selected_rows_C3
        ]
        for row in rows_to_move:
            df_C2_new.append(row)

    qtd_c1 = len(df_C1_new)
    qtd_c2 = len(df_C2_new)
    qtd_c3 = len(df_C3_new)

    return (
        pd.DataFrame.from_dict(df_C1_new).to_dict("records"),
        pd.DataFrame.from_dict(df_C2_new).to_dict("records"),
        pd.DataFrame.from_dict(df_C3_new).to_dict("records"),
        [],
        [],
        [],
        dcc.Markdown("### Quantidade = {}".format(qtd_c1)),
        dcc.Markdown("### Quantidade = {}".format(qtd_c2)),
        dcc.Markdown("### Quantidade = {}".format(qtd_c3)),
        "",
        "",
        "",
    )


@callback(
    Output("download-dataframe-xlsx", "data"),
    Input("Imprimir", "n_clicks"),
    State("tabela-solicitacoes-c-1", "data"),
    State("tabela-solicitacoes-c-2", "data"),
    State("tabela-solicitacoes-c-3", "data"),
    prevent_initial_call=True,
)
def exportar_para_excel(n_clicks, df_C1, df_C2, df_C3):
    if n_clicks > 0:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pd.DataFrame(df_C1).to_excel(writer, sheet_name="Caminhão 1", index=False)
            pd.DataFrame(df_C2).to_excel(writer, sheet_name="Caminhão 2", index=False)
            pd.DataFrame(df_C3).to_excel(writer, sheet_name="Caminhão 3", index=False)
        
        output.seek(0)
        return dcc.send_bytes(output.read(), "rotas_caminhoes.xlsx")
    
def update_tab(n_clicks, df_C1, df_C2, df_C3):
    if n_clicks > 0:
        writer = pd.ExcelWriter("rota.xlsx", engine="xlsxwriter")
        df_C1 = pd.DataFrame.from_dict(df_C1)
        df_C2 = pd.DataFrame.from_dict(df_C2)
        df_C3 = pd.DataFrame.from_dict(df_C3)
        df_C1.to_excel(writer, sheet_name="Caminhão 1")
        df_C2.to_excel(writer, sheet_name="Caminhão 2")
        df_C3.to_excel(writer, sheet_name="Caminhão 3")
        writer.save()
        return dcc.send_file("rota.xlsx")


@callback(
    Output("tabela-solicitacoes-c-1", "data", allow_duplicate=True),
    Output("tabela-solicitacoes-c-2", "data", allow_duplicate=True),
    Output("tabela-solicitacoes-c-3", "data", allow_duplicate=True),
    Output("tabela-solicitacoes", "data", allow_duplicate=True),
    Output("tab-caminhao-quantidade-c1", "children"),
    Output("tab-caminhao-quantidade-c2", "children"),
    Output("tab-caminhao-quantidade-c3", "children"),
    Input("move-C1", "n_clicks"),
    Input("move-C2", "n_clicks"),
    Input("move-C3", "n_clicks"),
    State("tabela-solicitacoes", "derived_virtual_selected_rows"),  # Linhas selecionadas
    State("tabela-solicitacoes", "derived_virtual_data"),  # Dados filtrados
    State("tabela-solicitacoes-c-1", "data"),
    State("tabela-solicitacoes-c-2", "data"),
    State("tabela-solicitacoes-c-3", "data"),
    State("tabela-solicitacoes", "data"),
    prevent_initial_call=True,
)
def mover_solicitacoes_para_caminhao(
    n_clicks_1, n_clicks_2, n_clicks_3, 
    selected_rows, derived_virtual_data, 
    df_C1, df_C2, df_C3, data
):
    button_clicked = ctx.triggered_id
    
    # Dados atuais das tabelas
    df_main = pd.DataFrame(data)
    df_C1_current = pd.DataFrame(df_C1)
    df_C2_current = pd.DataFrame(df_C2)
    df_C3_current = pd.DataFrame(df_C3)
    
    if selected_rows and button_clicked:
        # Extrai as linhas selecionadas (índices virtuais)
        selected_indices = selected_rows
        rows_to_move = [derived_virtual_data[i] for i in selected_indices]
        
        # Remove as linhas da tabela principal
        solicitacoes_to_remove = {row["Solicitacao"] for row in rows_to_move}
        df_main_updated = df_main[~df_main["Solicitacao"].isin(solicitacoes_to_remove)]
        
        # Adiciona às tabelas dos caminhões
        if button_clicked == "move-C1":
            df_C1_updated = pd.concat([df_C1_current, pd.DataFrame(rows_to_move)])
            df_C2_updated = df_C2_current
            df_C3_updated = df_C3_current
        elif button_clicked == "move-C2":
            df_C2_updated = pd.concat([df_C2_current, pd.DataFrame(rows_to_move)])
            df_C1_updated = df_C1_current
            df_C3_updated = df_C3_current
        elif button_clicked == "move-C3":
            df_C3_updated = pd.concat([df_C3_current, pd.DataFrame(rows_to_move)])
            df_C1_updated = df_C1_current
            df_C2_updated = df_C2_current
        
        # Atualiza quantidades
        qtd_c1 = len(df_C1_updated)
        qtd_c2 = len(df_C2_updated)
        qtd_c3 = len(df_C3_updated)
        
        return (
            df_C1_updated.to_dict("records"),
            df_C2_updated.to_dict("records"),
            df_C3_updated.to_dict("records"),
            df_main_updated.to_dict("records"),
            dcc.Markdown(f"### Quantidade = {qtd_c1}"),
            dcc.Markdown(f"### Quantidade = {qtd_c2}"),
            dcc.Markdown(f"### Quantidade = {qtd_c3}"),
        )
    
    # Retorna os dados originais se não houver seleção
    return (
        df_C1_current.to_dict("records"),
        df_C2_current.to_dict("records"),
        df_C3_current.to_dict("records"),
        df_main.to_dict("records"),
        dcc.Markdown(f"### Quantidade = {len(df_C1_current)}"),
        dcc.Markdown(f"### Quantidade = {len(df_C2_current)}"),
        dcc.Markdown(f"### Quantidade = {len(df_C3_current)}"),
    )

@callback(
    Output("tabela-solicitacoes-c-1", "selected_rows", allow_duplicate=True),
    Output("tabela-solicitacoes-c-2", "selected_rows", allow_duplicate=True),
    Output("tabela-solicitacoes-c-3", "selected_rows", allow_duplicate=True),
    Input("C1-select-all", "n_clicks"),
    Input("C2-select-all", "n_clicks"),
    Input("C3-select-all", "n_clicks"),
    State("tabela-solicitacoes-c-1", "derived_virtual_data"),
    State("tabela-solicitacoes-c-2", "derived_virtual_data"),
    State("tabela-solicitacoes-c-3", "derived_virtual_data"),
    State("tabela-solicitacoes-c-1", "data"),
    State("tabela-solicitacoes-c-2", "data"),
    State("tabela-solicitacoes-c-3", "data"),
    State("tabela-solicitacoes-c-1", "selected_rows"),
    State("tabela-solicitacoes-c-2", "selected_rows"),
    State("tabela-solicitacoes-c-3", "selected_rows"),
    prevent_initial_call=True,
)
def select_all(
    n_clicks_1,
    n_clicks_2,
    n_clicks_3,
    derived_virtual_data_C1,
    derived_virtual_data_C2,
    derived_virtual_data_C3,
    df_C1,
    df_C2,
    df_C3,
    selected_rows_C1,
    selected_rows_C2,
    selected_rows_C3,
):
    button_clicked = ctx.triggered_id
    if button_clicked == "C1-select-all":
        if selected_rows_C1 == []:
            indexs = [
                i
                for i in range(len(df_C1))
                if df_C1[i]["Solicitacao"]
                in [row["Solicitacao"] for row in derived_virtual_data_C1]
            ]
            return indexs, [], []
        else:
            return [], [], []
    elif button_clicked == "C2-select-all":
        if selected_rows_C2 == []:
            indexs = [
                i
                for i in range(len(df_C2))
                if df_C2[i]["Solicitacao"]
                in [row["Solicitacao"] for row in derived_virtual_data_C2]
            ]
            return [], indexs, []
        else:
            return [], [], []
    elif button_clicked == "C3-select-all":
        if selected_rows_C3 == []:
            indexs = [
                i
                for i in range(len(df_C3))
                if df_C3[i]["Solicitacao"]
                in [row["Solicitacao"] for row in derived_virtual_data_C3]
            ]
            return [], [], indexs
        else:
            return [], [], []

    return [], [], []


@callback(
    Output("tabela-solicitacoes", "data"),
    Input("tabela-solicitacoes", "selected_rows"),
)
def refresh_page(selected_rows):
    print("Reload")
    df_solicitacoes, df_install_2, df_recolhimento_1,df_recolhimento_2,df_recolhimento_3 = getData(r"C:\Users\rudso\Downloads\ROTA_DASH (1)\ListarSolic20251.XLSX")
    print(len(df_solicitacoes))
    return df_solicitacoes.to_dict("records")


@callback(
    Output("collapse", "is_open"),
    Input("collapse-button", "n_clicks"),
    State("collapse", "is_open"),
)
def toggle_collapse(n, is_open):
    if n:
        return not is_open
    return is_open


@callback(
    Output("tabela-solicitacoes", "data", allow_duplicate=True),
    Output("tabela-solicitacoes-c-1", "data", allow_duplicate=True),
    Output("tabela-solicitacoes-c-2", "data", allow_duplicate=True),
    Output("tabela-solicitacoes-c-3", "data", allow_duplicate=True),
    Input("tabela-solicitacoes", "data_timestamp"),
    Input("tabela-solicitacoes-c-1", "data_timestamp"),
    State("tabela-solicitacoes", "data"),
    State("tabela-solicitacoes-c-1", "data"),
    State("tabela-solicitacoes-c-2", "data"),
    State("tabela-solicitacoes-c-3", "data"),
    prevent_initial_call=True,
)
def formatar_niveis(ts1, ts2, data_main, data_c1, data_c2, data_c3):
    def add_badge(row):
        nivel = row["nivAprAtual"]
        cor = "#e74c3c" if nivel == 1 else "#f1c40f" if nivel == 2 else "#2ecc71"
        row["nivAprAtual"] = f'<span style="background-color: {cor}; padding: 4px; border-radius: 4px; color: white">N{nivel}</span>'
        return row

    data_main_formatted = [add_badge(row) for row in data_main]
    data_c1_formatted = [add_badge(row) for row in data_c1]
    data_c2_formatted = [add_badge(row) for row in data_c2]
    data_c3_formatted = [add_badge(row) for row in data_c3]

    return data_main_formatted, data_c1_formatted, data_c2_formatted, data_c3_formatted

if __name__ == "__main__":
    try:
        app.run(debug=False)
    except Exception:
        input("Press Enter to continue...")
