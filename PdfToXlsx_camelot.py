import streamlit as st
import pandas as pd
import camelot
import io

st.set_page_config(page_title="PDF → Excel Extractor", page_icon="📄", layout="centered")

# --- TITRE DE L'APPLICATION ---
st.title("📄 Extraction de tableaux PDF → Excel")
st.write("Téléchargez un fichier PDF, indiquez les pages, et obtenez un fichier Excel contenant les tableaux.")

# --- UPLOAD DU PDF ---
uploaded_pdf = st.file_uploader("📤 Téléchargez votre fichier PDF", type=["pdf"])

# --- SAISIE DES PAGES ---
pages = st.text_input(
    "📑 Entrez les numéros de pages (exemple : `1`, `1,2,3` ou `5-10`)",
    value="1"
)

# --- BOUTON POUR LANCER L'EXTRACTION ---
if uploaded_pdf is not None:
    if st.button("🚀 Extraire les tableaux"):
        try:
            # Lecture du PDF depuis le buffer mémoire
            with open("temp.pdf", "wb") as f:
                f.write(uploaded_pdf.read())

            # Extraction des tableaux avec Camelot
            st.info("⏳ Extraction des tableaux en cours...")
            tables = camelot.read_pdf("temp.pdf", pages=pages, flavor="stream")

            if len(tables) == 0:
                st.warning("⚠️ Aucun tableau trouvé dans ces pages.")
            else:
                st.success(f"✅ {len(tables)} tableau(x) trouvé(s).")

                # Création d'un buffer mémoire pour l'Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for i, table in enumerate(tables):
                        df = table.df
                        df.to_excel(writer, sheet_name=f"Tableau_{i+1}", index=False)

                # Préparer le fichier pour téléchargement
                output.seek(0)
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=output,
                    file_name="tableaux_extraits.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Une erreur est survenue : {str(e)}")
else:
    st.info("⬆️ Veuillez téléverser un fichier PDF pour commencer.")
