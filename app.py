import io
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import streamlit as st

st.set_page_config(
    page_title="Gerador de Etiquetas - Langner Assados", layout="centered"
)

st.title("🏷️ Gerador de Etiquetas")
st.write(
    "Preencha os dados abaixo, escolha o peso e ajuste a escala da logo."
)

with st.form("form_etiquetas"):
  produto = st.text_input("Nome do Produto", value="Feijão tropeiro")
  data_fab = st.text_input("Data de Fabricação", value="17/10/2026")
  obs = st.text_input("Observação / Validade", value="Consumo Imediato!")
  empresa = st.text_input("Nome da Empresa", value="Langner Assados")

  # Opções de peso pré-definidas de 100g até 1000g + opção manual
  opcoes_peso = [
      "100G",
      "150G",
      "200G",
      "250G",
      "300G",
      "350G",
      "400G",
      "450G",
      "500G",
      "600G",
      "700G",
      "800G",
      "900G",
      "1000G (1KG)",
      "Outro (digitar)",
  ]

  peso_selecionado = st.selectbox(
      "Peso / Tamanho", opcoes_peso, index=3
  )  # Padrão selecionado em 250G

  # Caso precise de um peso fora do padrão
  if peso_selecionado == "Outro (digitar)":
    peso = st.text_input("Digite o peso personalizado", value="250G")
  else:
    peso = peso_selecionado

  # Upload da Logo
  logo_file = st.file_uploader(
      "Envie a imagem da Logo (Opcional - PNG ou JPG)",
      type=["png", "jpg", "jpeg"],
  )

  # Controle de Escala da Logo
  tamanho_logo = st.slider(
      "Tamanho / Escala da Logo (em polegadas)",
      min_value=0.3,
      max_value=1.5,
      value=0.60,
      step=0.05,
      help=(
          "Diminua se o texto ficar apertado ou aumente se a imagem estiver"
          " pequena."
      ),
  )

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

  # Configurar margens da página (0,5 polegada)
  for section in doc.sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

  total_linhas = (qtd_total + 1) // 2
  main_table = doc.add_table(rows=total_linhas, cols=2)
  main_table.style = "Table Grid"
  main_table.alignment = WD_TABLE_ALIGNMENT.CENTER

  count = 0
  for row in main_table.rows:
    for cell in row.cells:
      if count < qtd_total:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        if logo_file is not None:
          sub_table = cell.add_table(rows=1, cols=2)
          sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER

          # Célula da Esquerda: Logo com tamanho dinâmico
          cell_logo = sub_table.rows[0].cells[0]
          cell_logo.width = Inches(tamanho_logo + 0.15)
          p_logo = cell_logo.paragraphs[0]
          p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
          logo_bytes = io.BytesIO(logo_file.getvalue())
          p_logo.add_run().add_picture(logo_bytes, width=Inches(tamanho_logo))

          # Célula da Direita: Texto
          cell_text = sub_table.rows[0].cells[1]
          cell_text.width = Inches(3.2 - (tamanho_logo + 0.15))
          p_text = cell_text.paragraphs[0]
          p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
          texto = f"{produto}\nfáb.{data_fab}\n{obs}\n{empresa}\n{peso}"
          run = p_text.add_run(texto)
          run.font.name = "Californian FB"
          run.font.size = Pt(13)
        else:
          # Sem logo: Texto centralizado na célula inteira
          p = cell.paragraphs[0]
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
          texto = f"{produto}\nfáb.{data_fab}\n{obs}\n{empresa}\n{peso}"
          run = p.add_run(texto)
          run.font.name = "Californian FB"
          run.font.size = Pt(15)

        count += 1
      else:
        cell.text = ""

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)

  st.success(f"✅ {qtd_total} etiquetas geradas com sucesso!")
  st.download_button(
      label="📄 Baixar Documento Word (.docx)",
      data=buffer,
      file_name=f"Etiquetas_{produto.replace(' ', '_')}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
