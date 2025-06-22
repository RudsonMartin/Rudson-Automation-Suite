import pandas as pd
import joblib
import os
import re
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import Pipeline
from sklearn.metrics import classification_report, confusion_matrix
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.utils import resample

# ====================
# CONFIGURAÇÕES
# ====================
ARQUIVO_ENTRADA = "ocorrencias_industriais.csv"
ARQUIVO_ROTULADO = "ocorrencias_rotuladas.csv"
ARQUIVO_MODELO = "modelo_manutencao.joblib"
ARQUIVO_SAIDA = "ocorrencias_classificadas.csv"
SEED = 42

# ====================
# PRÉ-PROCESSAMENTO
# ====================
def preprocessar_texto(texto):
    """Limpa e normaliza textos para melhorar a extração de features"""
    if not isinstance(texto, str):
        return ""
    
    # Conversão para minúsculas
    texto = texto.lower()
    # Remoção de caracteres especiais e números
    texto = re.sub(r'[^\w\s]', ' ', texto)
    texto = re.sub(r'\d+', ' ', texto)
    # Redução de espaços múltiplos
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# ====================
# BALANCEAMENTO DE DADOS
# ====================
def balancear_dados(df, coluna_alvo):
    """Balanceia classes minoritárias através de oversampling"""
    classes = df[coluna_alvo].unique()
    max_size = max(df[coluna_alvo].value_counts())
    
    dfs_balanceados = []
    for classe in classes:
        df_classe = df[df[coluna_alvo] == classe]
        if len(df_classe) < max_size:
            df_classe = resample(df_classe,
                                 replace=True,
                                 n_samples=max_size,
                                 random_state=SEED)
        dfs_balanceados.append(df_classe)
    
    return pd.concat(dfs_balanceados)

# ====================
# GERAR RECOMENDAÇÕES
# ====================
def gerar_recomendacao(classe):
    """Gera recomendações técnicas baseadas na classe predita"""
    respostas = {
        "falha_mecanica": "🔧 [Prioridade Alta] Verificar sistemas mecânicos: lubrificação, acoplamentos e rolamentos. Realizar análise vibracional.",
        "falha_eletrica": "⚡ [Prioridade Alta] Inspecionar cabos, conexões e alimentação elétrica. Verificar isolamento térmico e carga dos circuitos.",
        "vazamento": "💧 [Prioridade Urgente] Verificar juntas, válvulas e estanqueidade dos dutos. Monitorar pontos críticos com termografia.",
        "sobreaquecimento": "🌡️ [Prioridade Urgente] Avaliar sistemas de refrigeração e ventilação forçada. Verificar limpeza de trocadores de calor.",
        "vibracao_excessiva": "🔎 [Prioridade Média] Checar balanceamento e alinhamento de rotores. Inspecionar fundações e fixações mecânicas.",
        "pre_alarme": "⚠️ [Preventivo] Iniciar verificação preventiva antes da parada. Agendar manutenção programada.",
        "normal": "✅ [Operacional] Nenhuma ação corretiva necessária no momento. Continuar monitoramento rotineiro.",
        "outros": "📋 [Diagnóstico] Revisão técnica necessária. Coletar mais dados para análise detalhada."
    }
    return respostas.get(classe, "📋 [Diagnóstico] Revisão técnica necessária. Coletar mais dados para análise detalhada.")

# ====================
# VALIDAÇÃO INICIAL
# ====================
print("🔍 Validando arquivos e dados...")
for arquivo in [ARQUIVO_ENTRADA, ARQUIVO_ROTULADO]:
    if not os.path.exists(arquivo):
        print(f"❌ Arquivo não encontrado: {arquivo}")
        exit()

# Carregar dados rotulados
try:
    train_df = pd.read_csv(ARQUIVO_ROTULADO)
    required_columns = ['descricao', 'classe']
    if not all(col in train_df.columns for col in required_columns):
        print(f"❌ Arquivo de treinamento precisa das colunas: {required_columns}")
        exit()
except Exception as e:
    print(f"❌ Erro ao ler arquivo de treinamento: {str(e)}")
    exit()

# ====================
# PREPARAÇÃO DOS DADOS
# ====================
print("\n🧹 Pré-processando dados...")
train_df['descricao'] = train_df['descricao'].apply(preprocessar_texto)

# Balanceamento de classes
print("⚖️ Balanceando classes...")
train_df = balancear_dados(train_df, 'classe')

X = train_df["descricao"]
y = train_df["classe"]

# ====================
# TREINAMENTO DO MODELO
# ====================
print("\n🎯 Treinando modelo preditivo...")
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=SEED
)

modelo = Pipeline([
    ("tfidf", TfidfVectorizer(
        max_features=10000,
        ngram_range=(1, 3),
    ),
    ("clf", MultinomialNB(alpha=0.01))
])

modelo.fit(X_train, y_train)

# ====================
# AVALIAÇÃO DO MODELO
# ====================
print("\n📊 Avaliando performance do modelo...")
y_pred = modelo.predict(X_test)

print("\n📝 Relatório de classificação:")
print(classification_report(y_test, y_pred, digits=4))

# Matriz de confusão
plt.figure(figsize=(10, 8))
cm = confusion_matrix(y_test, y_pred)
sns.heatmap(cm, 
            annot=True, 
            fmt='d', 
            cmap='Blues',
            xticklabels=modelo.classes_,
            yticklabels=modelo.classes_)
plt.title("Matriz de Confusão - Classificação de Falhas")
plt.xlabel('Predito')
plt.ylabel('Real')
plt.savefig("matriz_confusao_manutencao.png", bbox_inches='tight', dpi=300)
print("\n📈 Matriz de confusão salva em alta resolução")

# ====================
# SALVAR MODELO
# ====================
joblib.dump(modelo, ARQUIVO_MODELO)
print(f"\n✅ Modelo salvo como '{ARQUIVO_MODELO}'")

# ====================
# CLASSIFICAÇÃO FINAL
# ====================
print("\n🔎 Classificando novas ocorrências...")
try:
    input_df = pd.read_csv(ARQUIVO_ENTRADA)
    input_df['descricao'] = input_df['descricao'].apply(preprocessar_texto)
    input_df["classe_predita"] = modelo.predict(input_df["descricao"])
    
    # Adicionar probabilidades das predições
    probabilidades = modelo.predict_proba(input_df["descricao"])
    for i, classe in enumerate(modelo.classes_):
        input_df[f'prob_{classe}'] = probabilidades[:, i]
    
    # Gerar recomendações
    input_df["recomendacao"] = input_df["classe_predita"].apply(gerar_recomendacao)
    
    # Ordenar por criticidade
    ordem_criticidade = {
        'vazamento': 1,
        'sobreaquecimento': 2,
        'falha_eletrica': 3,
        'falha_mecanica': 4,
        'vibracao_excessiva': 5,
        'pre_alarme': 6,
        'normal': 7,
        'outros': 8
    }
    input_df['prioridade'] = input_df['classe_predita'].map(ordem_criticidade)
    input_df = input_df.sort_values('prioridade')
    
    input_df.to_csv(ARQUIVO_SAIDA, index=False)
    print(f"📄 Dados classificados salvos em '{ARQUIVO_SAIDA}'")
    print("\n✅ Processo concluído com sucesso!")

except Exception as e:
    print(f"❌ Erro durante classificação: {str(e)}")