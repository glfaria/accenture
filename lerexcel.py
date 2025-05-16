import pandas as pd

def ler_dados_excel(caminho_ficheiro: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(caminho_ficheiro, engine="openpyxl")
        print(f"✅ Excel lido com sucesso: {caminho_ficheiro}")
        print(f"🔢 Número de linhas lidas: {len(df)}")
        return df
    except Exception as e:
        print(f"❌ Erro ao ler o Excel: {e}")
        return pd.DataFrame()
