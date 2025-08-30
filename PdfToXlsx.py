import streamlit as st
import pandas as pd
import tabula
import io

st.set_page_config(page_title="PDF → Excel Extractor", page_icon="📄", layout="centered")

st.title("📄 Extraction de tableaux PDF → Excel")
st.write("Téléchargez un PDF, choisissez les pages et récupérez un Excel avec les tableaux.")

uploaded_pdf = st.file_uploader("📤 Téléchargez votre PDF", type=["pdf"])
pages = st.text_input("📑 Entrez les pages (ex : `1`, `1,3,5` ou `2-5`)", value="1")

if uploaded_pdf is not None:
    if st.button("🚀 Extraire les tableaux"):
        try:
            with open("temp.pdf", "wb") as f:
                f.write(uploaded_pdf.read())

            # Extraction avec Tabula
            st.info("⏳ Extraction des tableaux...")
            tables = tabula.read_pdf("temp.pdf", pages=pages, multiple_tables=True)

            if not tables or len(tables) == 0:
                st.warning("⚠️ Aucun tableau trouvé sur ces pages.")
            else:
                st.success(f"✅ {len(tables)} tableau(x) trouvé(s).")

                # Écriture Excel dans un buffer mémoire
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for i, df in enumerate(tables):
                        df.to_excel(writer, sheet_name=f"Tableau_{i+1}", index=False)

                output.seek(0)
                st.download_button(
                    label="📥 Télécharger l'Excel",
                    data=output,
                    file_name="tableaux_extraits.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Erreur : {str(e)}")
else:
    st.info("⬆️ Téléchargez un PDF pour commencer.")
