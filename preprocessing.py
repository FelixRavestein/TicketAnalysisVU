import string
import os
import sys
import re
import nltk
from nltk.corpus import stopwords, wordnet as wn
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer

nltk_data_path = os.path.join(os.path.expanduser('~'), 'nltk_data')
if not os.path.exists(nltk_data_path):
    os.makedirs(nltk_data_path)
    nltk.download('stopwords', download_dir=nltk_data_path)
    nltk.download('punkt', download_dir=nltk_data_path)
    nltk.download('wordnet', download_dir=nltk_data_path)
else:
    nltk.data.path.append(nltk_data_path)

def read_forbidden_words(file_path):
    try:
        with open(file_path, 'r', encoding='utf8') as file:
            forbidden_words = file.read().splitlines()
            forbidden_words = [word.strip().lower().translate(str.maketrans('', '', string.punctuation)) for word in forbidden_words]
        return forbidden_words
    except Exception as e:
        print(f"Error reading forbidden words file: {e}")
        sys.exit(1)

def remove_punctuation(text):
    return "".join([i for i in text if i not in string.punctuation])

def tokenization(text):
    return word_tokenize(text)

def remove_stopwords(text, stopwords):
    return [i for i in text if i not in stopwords]

def get_lemma(word):
    lemma = wn.morphy(word)
    return word if lemma is None else lemma

def lemmatizer(text):
    lemmatizer = WordNetLemmatizer()
    return [lemmatizer.lemmatize(get_lemma(word)) for word in text]

def is_date(string):
    date_pattern = re.compile(r'\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b')
    return bool(date_pattern.match(string))

def is_numeric(string):
    return string.isdigit()

def preprocess_texts(texts, forbidden_words):
    stops = stopwords.words("dutch") + stopwords.words("english") + forbidden_words
    print("Stopwords including provided forbidden words successfully loaded")
    
    clean_texts = []
    for text in texts:
        if not isinstance(text, str):
            text = str(text)
        tokens = tokenization(text)
        tokens = [token for token in tokens if not is_date(token) and not is_numeric(token)]
        
        text = " ".join(tokens)
        text = remove_punctuation(text)
        text = text.lower()
        tokens = tokenization(text)
        tokens = remove_stopwords(tokens, stops)
        tokens = lemmatizer(tokens)
        clean_texts.append(" ".join(tokens))  # Join tokens back into a single string
    
    return clean_texts
