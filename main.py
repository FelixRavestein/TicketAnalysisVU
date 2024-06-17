import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from preprocessing import read_forbidden_words, preprocess_texts
from nlp_analysis import perform_nlp_analysis
from visualization import visualize_and_save_results, save_to_excel, generate_wordcloud

def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        sys.exit(1)

def perform_sentiment_analysis(data, output_folder):
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
    if len(sys.argv) != 3:
        print("Usage: python main.py <path_to_excel_file> <path_to_forbidden_words_file>")
        sys.exit(1)

    excel_file_path = sys.argv[1]
    forbidden_words_file_path = sys.argv[2]
    
    df = read_excel(excel_file_path)
    forbidden_words = read_forbidden_words(forbidden_words_file_path)
    
    print("Available columns:")
    for idx, column in enumerate(df.columns):
        print(f"{idx}: {column}")
    
    selected_column_index = int(input("Enter the column number to perform the analysis: "))
    selected_column = df.columns[selected_column_index]
    print(f"Selected column: {selected_column}")
    
    num_entries = input("Enter the number of entries to analyze (or 'all' to analyze the entire dataset): ")
    if num_entries.lower() == 'all':
        num_entries = len(df)
    else:
        num_entries = int(num_entries)
    
    data_subset = df.iloc[:num_entries]
    
    preprocessed_texts = preprocess_texts(data_subset[selected_column].astype(str), forbidden_words)
    data_subset['lemmatized'] = preprocessed_texts
    
    print(f"Number of preprocessed texts: {len(preprocessed_texts)}")
    print('Training the topic model, this could take a while...')
    
    output_folder = f"nlp_analysis_results_{selected_column}"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Set a default number of topics to 50 and adjust min_topic_size as needed
    n_topics = 50
    min_topic_size = 20  # Adjust this parameter as needed
    topic_model = perform_nlp_analysis(preprocessed_texts, n_topics, min_topic_size)
    
    visualize_and_save_results(topic_model, output_folder)
    generate_wordcloud(data_subset[selected_column], forbidden_words, output_folder)
    perform_sentiment_analysis(data_subset, output_folder)
    save_to_excel(topic_model, data_subset, output_folder)

if __name__ == "__main__":
    main()
