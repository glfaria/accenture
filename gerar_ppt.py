import pandas as pd
from pptx import Presentation
from pptx.util import Inches

# Lê os dados do Excel
df = pd.read_excel("Skills.xlsx")

# Cria a apresentação PowerPoint
ppt = Presentation()

# Slide de título
slide_titulo = ppt.slides.add_slide(ppt.slide_layouts[0])
slide_titulo.shapes.title.text = "Lista de Funcionários"
slide_titulo.placeholders[1].text = "Gerado automaticamente com Python"

# Um slide por funcionário
for _, row in df.iterrows():
    slide = ppt.slides.add_slide(ppt.slide_layouts[1])
    slide.shapes.title.text = row['Nome']
    corpo = slide.placeholders[1]
    corpo.text = (
        f"Cargo: {row['Cargo']}\n"
        f"Departamento: {row['Departamento']}\n"
        f"Email: {row['Email']}"
    )

# Guarda o ficheiro
ppt.save("apresentacao_funcionarios.pptx")
