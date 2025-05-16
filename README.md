# Accenture Employee Presentation Generator

Este projeto gera apresentações individuais em PowerPoint (formato 16:9) para colaboradores, com base num ficheiro Excel com dados sobre o seu perfil, experiência, competências e formação académica.

---

## 📁 Estrutura do Projeto


```text
.
├── CV.pptx                       # Exemplo de apresentação gerada
├── Employee_Presentations/       # Apresentações individuais geradas automaticamente (.pptx)
├── Skills.xlsx                   # Ficheiro Excel com os dados dos colaboradores
├── src/                          # Código-fonte principal
│   ├── lerexcel.py               # Lê e trata os dados do Excel
│   ├── gerartabela.py            # Resume e organiza os dados por colaborador
│   └── gerar_ppt.py              # Gera os slides individuais
├── requirements.txt              # Dependências Python
├── README.md                     # Este ficheiro
└── venv/                         # Ambiente virtual (opcional)