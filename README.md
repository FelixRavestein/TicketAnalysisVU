# NLP Executor for ServiceDesk Analysis

This project is designed to analyze ServiceDesk data using Natural Language Processing (NLP) techniques. It uses BERTopic to extract topics from the data and provides visualizations such as intertopic distance maps and word clouds. The results are saved in an Excel file and additional visualizations are generated.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Output](#output)
- [Customization](#customization)
- [Dependencies](#dependencies)
- [Troubleshooting](#troubleshooting)
- [Project Structure](#project-structure)
- [Preprocessing Steps](#preprocessing-steps)
- [Analysis Steps](#analysis-steps)
- [Notes](#notes)
- [License](#license)
- [Contact](#contact)

## Prerequisites

Ensure you have Python 3.9.18 installed on your machine.

## Requirements

- Python 3.9.18
- Dependencies listed in `requirements.txt`

## Installation

1. **Clone the repository** or download the `main.py`, `preprocessing.py`, `nlp_analysis.py`, and `visualization.py` scripts to your local machine.

2. **Install the required Python libraries**:

   pip install -r requirements.txt

3. **Ensure NLTK stopwords and VADER lexicon are downloaded**:

   python -m nltk.downloader stopwords
   python -m nltk.downloader punkt
   python -m nltk.downloader wordnet
   python -m nltk.downloader vader_lexicon

## Usage

1. **Prepare Your Data**:
    - Ensure your data is in an Excel file (.xlsx).
    - Create a text file containing forbidden words, one word per line.

2. **Run the Script**:
    - Execute the `main.py` script with the paths to your Excel file and forbidden words file as arguments:

      python main.py <path_to_your_excel_file.xlsx> <path_to_your_forbidden_words.txt>

3. **Select Column and Entries**:
    - The script will prompt you to select the column to perform the analysis on.
    - You can specify the number of entries to analyze or choose 'all' to analyze the entire dataset.

4. **Follow the prompts**:
    - The script will display the available columns in the Excel file and prompt you to enter the column number to perform the analysis on.
    - It will then ask you to specify the number of entries to analyze. You can enter a number or type 'all' to analyze the entire dataset.
    - The script will inform you about its progress and save the results in a folder named `nlp_analysis_results` along with visualizations and an Excel report.

## Example

1. **Run the script**:

   python main.py data.xlsx forbidden_words.txt

2. **Script prompts**:
    - The script lists available columns:
      ```
      Available columns:
      0: Question
      1: Answer
      2: Date
      Enter the column number to perform the analysis: 0
      Selected column: Question
      ```

    - The script asks for the number of entries to analyze:
      ```
      You can specify the number of entries to analyze.
      Enter a number, or type 'all' to analyze the entire dataset.
      How many entries would you like to analyze? 1000
      Analyzing the first 1000 entries.
      ```

3. **Output**:
    - The script will preprocess the texts, perform NLP analysis, and save the results.
    - Results are saved in a folder named `nlp_analysis_results_<column_name>`.
    - Visualizations (HTML and PNG files) and an Excel report are saved in the same directory.

## Customization

- **Number of Topics**: The script is set to limit the number of topics to 50 by default. You can adjust this in the `perform_nlp_analysis` function in `nlp_analysis.py`.
- **Minimum Topic Size**: You can adjust the minimum topic size parameter to reduce the number of outliers.

## Dependencies

The project requires the following Python packages:

- bertopic==0.15.0
- gensim==4.3.2
- hdbscan==0.8.33
- matplotlib==3.7.1
- nltk==3.8.1
- numpy==1.24.3
- pandas==1.3.4
- pandas_stubs==1.2.0.35
- scikit_learn==1.2.2
- sentence_transformers==2.2.2
- spacy==3.7.1
- umap-learn==0.5.3
- wordcloud==1.9.2
- keybert==0.4.0
- openpyxl==3.0.10
- plotly==5.9.0

Install these dependencies using:

   pip install -r requirements.txt

## Troubleshooting

- **Data Quality**: Ensure your data is clean and well-preprocessed.
- **Adjust Parameters**: Fine-tune the BERTopic model parameters if the number of outliers is too high or if the topics are not meaningful.

## Project Structure

The project consists of the following main files:

1. **main.py**: The main script that runs the analysis.
2. **preprocessing.py**: Contains functions for preprocessing the text data.
3. **nlp_analysis.py**: Contains functions for performing the NLP analysis using BERTopic.
4. **visualization.py**: Contains functions for generating and saving visualizations.

## Preprocessing Steps

1. **Remove Punctuation**: Strips punctuation from the text.
2. **Lowercase Conversion**: Converts text to lowercase.
3. **Tokenization**: Splits text into individual words (tokens).
4. **Remove Stopwords**: Removes common English and Dutch stopwords, as well as any additional forbidden words.
5. **Lemmatization**: Reduces words to their base or root form.
6. **Remove Dates and Numeric Strings**: Ensures that dates and purely numeric strings are not included in the analysis.

## Analysis Steps

1. **Topic Modeling**: Uses BERTopic to identify and extract topics from the text.
2. **Sentiment Analysis**: Uses VADER sentiment analysis to determine the sentiment of each document.
3. **Visualization**: Generates visualizations such as intertopic distance maps, bar charts of top words per topic, and word clouds.

## Notes

- Ensure that the column names in your Excel file are correctly named and match the column number you provide.
- The forbidden words file should contain one word per line; the code will convert all words  to lowercase.
- If you encounter any issues or have suggestions, feel free to open an issue or contribute to the project.

## License

This project is licensed under the MIT License.

## Contact

For any issues or questions, please contact [Felix Ravestein] at [f.ravestein@vu.nl].
