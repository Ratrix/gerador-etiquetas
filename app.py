import io
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import streamlit as st

st.set_page_config(
    page_title="Gerador de Etiquetas - Langner Assados", layout="centered"
)

st.title("🏷️ Gerador de Etiquetas")
st.write("Preencha os dados abaixo para gerar as etiquetas no Word.")

with st.form("form_etiquetas"):
  produto = st.text_input("Nome do Produto", value="Feijão tropeiro")
  data_fab = st.text_input("Data de Fabricação", value="01/11/2026")
  obs = st.text_input("Observação / Validade", value="Consumo Imediato!")
  empresa = st.text_input("Nome da Empresa", value="Langner Assados")
  peso = st.text_input("Peso / Tamanho", value="250G")
  qtd_total = st.number_input(
      "Quantidade total de etiquetas",
      min_value=1,
      max_value=200,
      value=12,
      step=1,
  )

  submitted = st.form_submit_button("Gerar Arquivo Word")

if submitted:
  doc = Document()

  # Configurar margens da folha (0,5 polegada)
  for section in doc.sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

  total_linhas = (qtd_total + 1) // 2
  table = doc.add_table(rows=total_linhas, cols=2)
  table.style = "Table Grid"  # Adiciona a grade/bordas visíveis
  table.alignment = WD_TABLE_ALIGNMENT.CENTER

  count = 0
  for row in table.rows:
    for cell in row.cells:
      if count < qtd_total:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        texto = f"{produto}\nfáb.{data_fab}\n{obs}\n{empresa}\n{peso}"
        run = p.add_run(texto)
        run.font.name = "Californian FB"
        run.font.size = Pt(16)
        count += 1
      else:
        cell.text = ""

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)

  st.success(f"✅ {qtd_total} etiquetas geradas com bordas!")
  st.download_button(
      label="📄 Baixar Documento Word (.docx)",
      data=buffer,
      file_name=f"Etiquetas_{produto.replace(' ', '_')}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
