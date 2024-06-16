import os
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
import nltk
from nltk.corpus import stopwords
import plotly.io as pio
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import pandas as pd

nltk.download('stopwords')

def visualize_and_save_results(topic_model, output_folder):
    print("\nGenerating intertopic distance map...")
    fig_topics = topic_model.visualize_topics(width=1200, height=800)
    pio.write_html(fig_topics, file=os.path.join(output_folder, 'intertopic_distance_map.html'), auto_open=False)
    print("Intertopic distance map saved as 'intertopic_distance_map.html'.")

    print("\nGenerating bar chart of top words per topic...")
    fig_barchart = topic_model.visualize_barchart(top_n_topics=10, n_words=10, width=1200, height=800)
    pio.write_html(fig_barchart, file=os.path.join(output_folder, 'top_words_per_topic.html'), auto_open=False)
    print("Bar chart of top words per topic saved as 'top_words_per_topic.html'.")

def preprocess_for_wordcloud(texts, forbidden_words):
    stopwords_english = set(STOPWORDS)
    stopwords_dutch = set(stopwords.words("dutch"))
    stopwords_combined = stopwords_english.union(stopwords_dutch).union(set(forbidden_words))
    
    # Custom function to clean and preprocess the text
    def clean_text(text):
        tokens = text.split()
        tokens = [token.lower() for token in tokens if token.isalpha() and token.lower() not in stopwords_combined]
        return ' '.join(tokens)
    
    preprocessed_texts = [clean_text(text) for text in texts]
    return ' '.join(preprocessed_texts)

def generate_wordcloud(texts, forbidden_words, output_folder):
    print("\nGenerating word cloud...")
    
    # Preprocess texts for the word cloud
    data = preprocess_for_wordcloud(texts, forbidden_words)

    wordcloud = WordCloud(
        background_color='white',
        stopwords=STOPWORDS,
        max_words=200,
        max_font_size=40,
        scale=3,
        random_state=1
    ).generate(data)

    plt.figure(figsize=(12, 12))
    plt.axis('off')
    plt.imshow(wordcloud, interpolation='bilinear')
    wordcloud_path = os.path.join(output_folder, 'wordcloud.png')
    plt.savefig(wordcloud_path)
    plt.close()
    print(f"Word cloud saved as '{wordcloud_path}'.")

def save_to_excel(topic_model, data, output_folder):
    topic_info = topic_model.get_topic_info()
    top_topics = topic_model.get_topic_freq()

    def clean_topic_names(topic_model, top_topics):
        names = []
        for topic in top_topics['Topic']:
            if topic == -1:
                names.append("Outliers")
            else:
                words = [word for word, _ in topic_model.get_topic(topic)]
                name = ", ".join(words[:5])
                names.append(name)
        return names

    top_topics['Name'] = clean_topic_names(topic_model, top_topics)

    output_file = os.path.join(output_folder, 'topic_modeling_report.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        topic_info.to_excel(writer, sheet_name='Topic Information', index=False)
        top_topics.to_excel(writer, sheet_name='Top Topics', index=False)
        
        worksheet = writer.book.create_sheet('Sentiment Distribution')
        img = Image(os.path.join(output_folder, 'sentiment_distribution.png'))
        worksheet.add_image(img, 'A1')

        worksheet = writer.book.create_sheet('Visualizations')
        worksheet['A1'] = 'Intertopic Distance Map'
        worksheet['A2'] = 'Top Words Per Topic'
        worksheet['A3'] = 'Word Cloud'
        worksheet['B1'] = 'Link'
        worksheet['B2'] = 'Link'
        worksheet['B3'] = 'Image'
        worksheet['B1'].hyperlink = 'intertopic_distance_map.html'
        worksheet['B2'].hyperlink = 'top_words_per_topic.html'
        worksheet['B1'].style = 'Hyperlink'
        worksheet['B2'].style = 'Hyperlink'

        img = Image(os.path.join(output_folder, 'wordcloud.png'))
        worksheet.add_image(img, 'A3')

    print(f"Report generated and saved as '{output_file}'.")
