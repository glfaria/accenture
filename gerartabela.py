import pandas as pd
from typing import List, Dict

def gerar_perfil(row: pd.Series) -> str:
    """Gera um texto de perfil com base em áreas de atuação e função."""
    industry = row.get('Industry Networks', 'diversas indústrias')
    function = row.get('Function Networks', 'funções variadas')
    technology = row.get('Technology Networks', 'tecnologias diversas')
    job_family = row.get('Job Family', 'diferentes áreas funcionais')

    return (
        f"Profissional com experiência em {industry}, "
        f"com atuação em {function} e uso de {technology}. "
        f"Integrado na área de {job_family}."
    )

def gerar_skills(rows: List[pd.Series]) -> str:
    """Gera uma lista formatada de competências com proficiência."""
    seen = set()
    skills = []
    for row in rows:
        skill = str(row.get('Skill', '')).strip()
        proficiency = str(row.get('Skill Proficiency', '')).strip()
        if skill and skill not in seen:
            skills.append(f"- {skill} ({proficiency})")
            seen.add(skill)
    return "\n".join(skills)

def gerar_resumo_trabalhadores(df: pd.DataFrame) -> pd.DataFrame:
    """
    Processa os dados do Excel e gera uma tabela resumo por trabalhador.
    
    Retorna:
        pd.DataFrame com colunas:
            - Worker Name
            - Job Title / Role
            - Profile
            - Education
            - Relevant Skills & Qualifications
            - Relevant Experience
    """
    resumo = []

    for nome, grupo in df.groupby("Worker Name"):
        row_exemplo = grupo.iloc[0]
        job_title = row_exemplo.get("Job Title", "Função não especificada")
        profile = gerar_perfil(row_exemplo)
        skills = gerar_skills(grupo.to_dict("records"))

        resumo.append({
            "First Name Last Name": nome,
            "Job Title / Role": job_title,
            "Profile": profile,
            "Education": "Informação disponível sob pedido.",
            "Relevant Skills & Qualifications": skills,
            "Relevant Experience": "Informação disponível sob pedido."
        })

    return pd.DataFrame(resumo)

def ver_perfil_trabalhador(tabela_resumo: pd.DataFrame, nome: str) -> None:
    """
    Exibe o perfil completo de um trabalhador com base no nome.
    
    Args:
        tabela_resumo (pd.DataFrame): DataFrame com os perfis dos trabalhadores.
        nome (str): Nome do trabalhador (deve corresponder exatamente ao da coluna 'First Name Last Name').
    """
    trabalhador = tabela_resumo[tabela_resumo["First Name Last Name"] == nome]

    if trabalhador.empty:
        print(f"❌ Nenhum trabalhador encontrado com o nome: {nome}")
        return

    linha = trabalhador.iloc[0]
    print("📄 Perfil do Trabalhador")
    print("-" * 40)
    print(f"👤 Nome: {linha['First Name Last Name']}")
    print(f"💼 Cargo: {linha['Job Title / Role']}")
    print(f"\n🧾 Perfil:\n{linha['Profile']}")
    print(f"\n📚 Educação:\n{linha['Education']}")
    print(f"\n🛠️ Competências:\n{linha['Relevant Skills & Qualifications']}")
    print(f"\n🏆 Experiência:\n{linha['Relevant Experience']}")
    print("-" * 40)


# Exemplo de uso
if __name__ == "__main__":
    from lerexcel import ler_dados_excel  # Certifica-te que tens esse módulo
    df = ler_dados_excel("Skills.xlsx")
    tabela_resumo = gerar_resumo_trabalhadores(df)
    print(tabela_resumo)  # Mostra a tabela resumo gerada
    nome_para_ver = "Employee B"  # substitui por um nome existente
    ver_perfil_trabalhador(tabela_resumo, nome_para_ver)
