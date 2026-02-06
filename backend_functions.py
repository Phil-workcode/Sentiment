def extract_words(input_file: str, output_folder: str, progress = None) -> str:

    def report(message: str):
        if progress is not None:
            progress(message)

    # Set up the environment.
    import pathlib
    input_path = pathlib.Path(input_file)
    if not input_path.exists():
        return "❌ Input file does not exist."
    output_path = pathlib.Path(output_folder)
    if not output_path.exists():
        output_path.mkdir()


    import openpyxl
    import pandas as pd

    report("\n🔄 Identifying the column.")
    df_headers = pd.read_excel(input_file, nrows = 1)
    improvement_column = next((column for column in df_headers.columns if 'improve' in str(column).lower()), None)
    strength_column = next((column for column in df_headers.columns if any(keyword in str(column).lower() for keyword in ['strong', 'strength'])), None)
    if improvement_column is None:
        return f"❌ Missing improvement keyword, searched all of the following columns: {list(df_headers.columns)}"
    elif strength_column is None:
        return f"❌ Missing strength keyword, searched all of the following columns: {list(df_headers.columns)}"

    report("\n🔄 Loading the neural network.")
    try:
        import os, sys, spacy
        from pathlib import Path
        def _model_path() -> Path:
            if getattr(sys, 'frozen', False):
                base = Path(sys._MEIPASS)
            else:
                base = Path(__file__).parent

            root = base / 'en_core_web_sm'
            for sub in root.iterdir():
                if sub.is_dir() and (sub / 'config.cfg').exists():
                    return sub
            if (root / 'config.cfg').exists():
                return root
            return root

        model_dir = _model_path()
        nlp = spacy.load(model_dir)
    except Exception as e:
        return f"❌ Failed to load the neural network due to: {str(e)}"



    report("\n🔄 Extracting column contents.")
    improvement_paragraphs = pd.read_excel(input_file, usecols=[improvement_column])
    strength_paragraphs = pd.read_excel(input_file, usecols=[strength_column])

    improvement_words = {"Adjectives": [], "Nouns": []}
    strength_words = {"Adjectives": [], "Nouns": []}

    for index, row in improvement_paragraphs.iterrows():
        for content in row:
            if type(content) == str:
                sentences = nlp(content)
                adjectives = [token.text for token in sentences if token.pos_ == "ADJ"]
                nouns = [token.text for token in sentences if token.pos_ == "NOUN"]

                improvement_words["Adjectives"].extend(adjectives)
                improvement_words["Nouns"].extend(nouns)
            else:
                pass

    for index, row in strength_paragraphs.iterrows():
        for content in row:
            if type(content) == str:
                sentences = nlp(content)
                adjectives = [token.text for token in sentences if token.pos_ == "ADJ"]
                nouns = [token.text for token in sentences if token.pos_ == "NOUN"]

                strength_words["Adjectives"].extend(adjectives)
                strength_words["Nouns"].extend(nouns)
            else:
                pass


    # Create three dataframes, each dedicated to a word type, per dictionary.
    improvement_adj_df = pd.DataFrame({"Adjectives": improvement_words["Adjectives"]})
    strength_adj_df = pd.DataFrame({"Adjectives": strength_words["Adjectives"]})

    improvement_noun_df = pd.DataFrame({"Nouns": improvement_words["Nouns"]})
    strength_noun_df = pd.DataFrame({"Nouns": strength_words["Nouns"]})

    try:
        import xlsxwriter
        report("\n🔄 Converting to Excel.")

        improvement_adj_df.to_excel(output_path / "Improvement adjectives.xlsx", engine = 'xlsxwriter', index = False)
        strength_adj_df.to_excel(output_path / "Strength adjectives.xlsx", engine = 'xlsxwriter', index = False)

        improvement_noun_df.to_excel(output_path / "Improvement nouns.xlsx", engine = 'xlsxwriter', index = False)
        strength_noun_df.to_excel(output_path / "Strength nouns.xlsx", engine = 'xlsxwriter', index = False)
    except Exception as e:
        return f"❌ Could not save spreadsheet(s) due to: {str(e)}"


    return f"✅ Extracted {len(improvement_paragraphs)} improvement rows and {len(strength_paragraphs)} strength rows."
