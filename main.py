import streamlit as st
from docx import Document
from docx.shared import Pt
from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer
import datetime
import os
from pdf2docx import parse
import docx2pdf

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
    
        .stApp {
            background: linear-gradient(to top, #0e4166 40%, #000000 90%); #for background;
            font-family: 'Roboto', sans-serif;
        }

        [data-testid=stSidebar] {
            background: linear-gradient(to top, #0e4166 60%, #000000 90%); # for sidebar;
            margin-top:60px;
        }

        [data-testid=stHeader] {
            background: linear-gradient( #0e4166 20%, #000000 100%); #for header;
        }
        [data-testid=stWidgetLabel] {
            color:#f4a303; # for small headings;
        }

        .stHeading h1{ 
            color:#f4a303; # for Heading "select Languages";
        }

        [data-testid=stAppViewBlockContainer],[data-testid=stVerticalBlock]{
            margin-top:30px; # for centering the content and sidebar;
        }

        .title-container {
            background: rgb(0,27,44);
            backdrop-filter: blur(10px);
            border-radius: 30px;
            text-align:center;
            padding:5px;
            margin-bottom:10px;
            box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1);
        }
   
    <style>""", unsafe_allow_html=True)
    # st.title("Welcome To Novintix Language Translator")
header_html = f"""
    <div class="title-container">
        <span style="font-size: 30px; font-weight: bold;color:#023d59;">Welcome to</span>
        <img src="https://novintix.com/wp-content/uploads/2023/12/logov2.png" style="height: 45px;padding:5px;  margin-bottom: 10px;">
        <span style="font-size: 30px; font-weight: bold; color:#f4a303;">Multilingual Language Translation Tool</span>
    """

st.markdown(header_html, unsafe_allow_html=True)
# Mapping of full language names to language codes
language_mapping = {
    "Bulgarian": "bg",
    "Chinese": "zh",
    "Croatian": "hr",
    "Czech": "cs",
    "Danish": "da",
    "Dutch": "nl",
    "Estonian": "et",
    "English": "en",
    "Finnish": "fi",
    "French": "fr",
    "German": "de",
    "Greek": "el",
    "Hungarian": "hu",
    "Icelandic": "is",
    "Indonesian": "id",
    "Italian": "it",
    "Kazakh": "kk",
    "Korean": "ko",
    "Latvian": "lv",
    "Lithuanian": "lt",
    "Macedonian": "mk",
    "Norwegian": "no",
    "Polish": "pl",
    "Portuguese": "pt",
    "Romanian": "ro",
    "Russian": "ru",
    "Serbian": "sr",
    "Slovak": "sk",
    "Slovenian": "sl",
    "Spanish": "es",
    "Swedish": "sv",
    "Turkish": "tr",
    "Vietnamese": "vi"
}

# Function to load the translation model and tokenizer
def load_translation_model():
    model_name = 'facebook/m2m100_418M'
    tokenizer = M2M100Tokenizer.from_pretrained(model_name)
    model = M2M100ForConditionalGeneration.from_pretrained(model_name)
    return tokenizer, model

# Function to translate text
def translate_text(text: str, src_lang: str, tgt_lang: str, tokenizer, model):
    tokenizer.src_lang = src_lang
    encoded = tokenizer(text, return_tensors="pt",
                        padding=True, truncation=True)
    generated_tokens = model.generate(
        **encoded, forced_bos_token_id=tokenizer.get_lang_id(tgt_lang))
    translated_text = tokenizer.batch_decode(
        generated_tokens, skip_special_tokens=True)[0]
    return translated_text

# Function to copy formatting from one run to another
def copy_run_format(source_run, target_run):
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline
    target_run.font.name = source_run.font.name
    target_run.font.size = source_run.font.size
    target_run.font.color.rgb = source_run.font.color.rgb
    text = source_run.text
    if text.isupper():
        target_run.text = text.upper()
    elif text.islower():
        target_run.text = text.lower()

# Function to translate text while preserving the format and font size
def translate_text_with_format(paragraph, src_lang, tgt_lang, tokenizer, model):
    new_runs = []
    for run in paragraph.runs:
        # Translate text
        translated_text = translate_text(
            run.text, src_lang, tgt_lang, tokenizer, model)

        # Check and preserve the trademark symbol ("™")
        if "™" in run.text:
            translated_text = translated_text.replace("TM", "™")

        # Add a new run with the translated text
        new_run = paragraph.add_run(translated_text)
        copy_run_format(run, new_run)

        # Clear the original run's text to avoid duplication
        run.clear()

# Function to convert PDF to DOCX using pdf2docx
def convert_pdf_to_docx(pdf_path, docx_path):
    parse(pdf_path, docx_path, start=0, end=None)
    print("PDF to DOCX Done...")

# Function to convert DOCX to PDF using docx2pdf
def convert_docx_to_pdf(docx_path, pdf_path):
    docx2pdf.convert(docx_path, pdf_path)
    print("DOCX to PDF Done...")
    # to delete the converted docx
    os.remove('temp.docx')

# Function to translate a DOCX file while maintaining formatting
def translate_docx(doc_path, src_lang, tgt_langs, download_location, input_file_type):
    start_time = datetime.datetime.now()

    # Load the translation model and tokenizer
    tokenizer, model = load_translation_model()

    translated_files = {}

    for tgt_lang in tgt_langs:
        # Reload the DOCX file for each target language
        doc = Document(doc_path)

        # Translate paragraphs
        for para in doc.paragraphs:
            if para.text.strip():  # Check if the paragraph is not empty
                translate_text_with_format(
                    para, language_mapping[src_lang], language_mapping[tgt_lang], tokenizer, model)

        # Translate tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():  # Check if the cell is not empty
                        cell.text = translate_text(
                            cell.text, language_mapping[src_lang], language_mapping[tgt_lang], tokenizer, model)

        # Ensure output directory exists
        output_dir = os.path.join(download_location, "output")
        os.makedirs(output_dir, exist_ok=True)

        # Save the translated DOCX file
        output_docx_path = os.path.join(
            output_dir, f"translated_{src_lang}_to_{tgt_lang}.docx")

        # output_docx_path = os.path.join(output_dir, f"test_output_{tgt_lang}.docx")
        doc.save(output_docx_path)

        # Convert DOCX back to PDF if the original file was PDF
        if input_file_type == 'pdf':
            pdf_output_path = output_docx_path.replace(".docx", ".pdf")
            convert_docx_to_pdf(output_docx_path, pdf_output_path)
            os.remove(output_docx_path)  # Remove the DOCX file after conversion
            translated_files[tgt_lang] = pdf_output_path
        else:
            translated_files[tgt_lang] = output_docx_path

    end_time = datetime.datetime.now()
    time_diff = end_time - start_time
    hours, remainder = divmod(time_diff.seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    print(f"Translation completed in {hours} hours, {minutes} minutes, and {seconds} seconds.")

    return translated_files


def main():

    # Language selection
    st.sidebar.title("Select Languages")
    src_lang = st.sidebar.selectbox(
        "Select source language:", list(language_mapping.keys()))
    tgt_langs = st.sidebar.multiselect(
        "Select target languages:", list(language_mapping.keys()))

    # File uploader for Word or PDF file
    input_file = st.file_uploader(
        "Upload Word or PDF file for translation:", type=["docx", "pdf"])

    if input_file:
        input_file_type = input_file.name.split('.')[-1].lower()
        input_file_path = os.path.join(os.getcwd(), input_file.name)
        with open(input_file_path, 'wb') as f:
            f.write(input_file.read())

    # Download location selection
    download_location = st.text_input(
        "Enter download location:")

    # Translate button
    if st.button("Translate"):
        if not input_file:
            st.warning("Please upload a Word or PDF file for translation.")
        elif not src_lang:
            st.warning("Please select a source language.")
        elif not tgt_langs:
            st.warning("Please select at least one target language.")
        elif not download_location:
            st.warning("Please enter the download location.")

        else:
            with st.spinner("Translating..."):
                temp_docx_path = os.path.join(download_location, "temp.docx")
                if input_file_type == 'pdf':
                    # Convert PDF to DOCX
                    convert_pdf_to_docx(input_file_path, temp_docx_path)
                else:
                    temp_docx_path = input_file_path

                translated_files = translate_docx(
                    temp_docx_path, src_lang, tgt_langs, download_location, input_file_type)
                st.success(
                    "Translation completed. Output files saved successfully.")
                st.info(
                    f"Translated files saved in {os.path.join(download_location, 'output')}")


if __name__ == "__main__":
    main()
