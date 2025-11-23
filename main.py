import os
import re
import sys
import pandas as pd
import matplotlib.pyplot as plt
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from preprocessing import read_forbidden_words, preprocess_data
from nlp_analysis import create_topics
from visualization import visualize_and_save_results, save_to_excel, generate_wordcloud, generate_html_report

# **Kolommen die altijd in de HTML moeten staan (maar niet in NLP-analyse)**
EXTRA_COLUMNS = [
    "nummer", "opened_at", "service_offering", "assignment_group",
    "category", "reassignment_count", "reopen_count"
]

def read_excel(file_path):
    """Lees een Excel-bestand en laad het in een pandas DataFrame."""
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        sys.exit(1)

def perform_sentiment_analysis(data, output_folder):
    """Voert sentimentanalyse uit en slaat de distributie als een afbeelding op."""
    print("\nPerforming sentiment analysis...")
    sid = SentimentIntensityAnalyzer()
    data['sentiment'] = data['lemmatized'].apply(lambda x: sid.polarity_scores(x)['compound'])

    plt.figure(figsize=(10, 5))
    plt.hist(data['sentiment'], bins=20, color='blue', edgecolor='black')
    plt.title('Sentiment Distribution of Service Desk Questions')
    plt.xlabel('Sentiment Score')
    plt.ylabel('Frequency')
    plt.savefig(os.path.join(output_folder, 'sentiment_distribution.png'))
    plt.close()
    print("Sentiment distribution saved as 'sentiment_distribution.png'.")

def main():
    """Hoofdfunctie die NLP-analyse uitvoert, inclusief preprocessing, topic modeling en visualisatie."""
    if len(sys.argv) != 3:
        print("Usage: python main.py <path_to_excel_file> <path_to_forbidden_words_file>")
        sys.exit(1)

    excel_file_path = sys.argv[1]
    forbidden_words_file_path = sys.argv[2]
    
    # **Stap 1: Data inlezen**
    df = read_excel(excel_file_path)
    forbidden_words = read_forbidden_words(forbidden_words_file_path)
    
    print("\nBeschikbare kolommen in het dataset:")
    for idx, column in enumerate(df.columns):
        print(f"{idx}: {column}")
    
    # **Stap 2: Kolommen selecteren voor NLP-analyse**
    selected_columns_indices = input("Enter the column numbers to include in the NLP analysis (comma-separated): ")
    selected_columns_indices = [int(idx) for idx in selected_columns_indices.split(',')]

    # **Gebruiker kiest NLP-kolommen, HTML-kolommen blijven apart**
    nlp_columns = [df.columns[idx] for idx in selected_columns_indices]
    html_columns = EXTRA_COLUMNS  # Extra kolommen voor HTML

    # **Verwijder HTML-kolommen uit de NLP-analyse**
    for col in html_columns:
        if col in nlp_columns:
            nlp_columns.remove(col)

    print(f"\nKolommen gebruikt voor NLP-analyse: {nlp_columns}")
    print(f"Kolommen bewaard voor HTML-output: {html_columns}")
    
    # **Stap 3: Aantal te analyseren rijen bepalen**
    num_entries = input("Enter the number of entries to analyze (or 'all' to analyze the entire dataset): ")
    if num_entries.lower() == 'all':
        num_entries = len(df)
    else:
        num_entries = int(num_entries)
    
    # **Maak aparte subsets voor NLP en HTML**
    nlp_data_subset = df.iloc[:num_entries][nlp_columns + ["nummer"]]
    html_data_subset = df.iloc[:num_entries][html_columns]
    
    # **Stap 4: NLP Preprocessing**
    preprocessed_beschrijving, nummer = preprocess_data(nlp_data_subset, forbidden_words)
    nlp_data_subset['lemmatized'] = preprocessed_beschrijving
    
    print(f"\nAantal verwerkte teksten: {len(preprocessed_beschrijving)}")
    print('Training the topic model, this could take a while...')
    
    # **Stap 5: Output-folder aanmaken**
    output_folder = "nlp_analysis_results"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # **Stap 6: Topic Modeling**
    topic_model, topic_to_inc = create_topics(preprocessed_beschrijving, nummer)
    
    # **Stap 7: Visualisaties genereren**
    visualize_and_save_results(topic_model, topic_to_inc, output_folder)
    generate_wordcloud(nlp_data_subset[nlp_columns[0]], forbidden_words, output_folder)
    perform_sentiment_analysis(nlp_data_subset, output_folder)
    save_to_excel(topic_model, nlp_data_subset, output_folder)  

    # **Stap 8: HTML-rapport genereren**
    topic_info = topic_model.get_topic_info()

    def clean_topic_name(name):
        """Verbeter de leesbaarheid van topic-namen door structuur aan te brengen."""
        
        # 1️⃣ Verwijder getallen aan het begin (zoals "6_printer_printen_papercut")
        name = re.sub(r'^\d+\s*', '', name)

        # 2️⃣ Onderstreping vervangen door spaties en woorden splitsen
        words = name.replace("_", " ").split()

        # 3️⃣ Verwijder dubbele woorden
        unique_words = list(dict.fromkeys(words))

        # 4️⃣ Stopwoorden verwijderen (maar geen essentiële woorden)
        stopwoorden = {"de", "het", "een", "voor", "van", "met", "op", "in", "bij", "te", "om", "door"}
        filtered_words = [word for word in unique_words if word.lower() not in stopwoorden]

        # 6️5️⃣ Eerste letter van de zin een hoofdletter geven
        formatted_topic = " ".join(filtered_words).capitalize()

        return formatted_topic

    # **Pas deze functie toe op alle topic-namen in `main.py`**
    topic_names = {topic: clean_topic_name(name) for topic, name in zip(topic_info['Topic'], topic_info['Name'])}


    generate_html_report(topic_to_inc, topic_names, html_data_subset, output_folder)  

    print("\n✅ Analyse voltooid! Check de gegenereerde bestanden in de 'nlp_analysis_results' map.")

if __name__ == "__main__":
    main()
