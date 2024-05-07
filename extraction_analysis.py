import pandas as pd
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
from textblob import TextBlob
from textstat import textstat

# Load nltk resources
nltk.download('punkt')

# Read input data from Excel
input_data = pd.read_excel('Input.xlsx')

# Function to perform text analysis and compute variables
def perform_text_analysis(article_text):
    # Tokenize text into sentences and words
    sentences = sent_tokenize(article_text)
    words = word_tokenize(article_text)
    
    # Compute variables
    word_count = len(words)
    sentence_count = len(sentences)
    avg_sentence_length = word_count / sentence_count if sentence_count > 0 else 0
    avg_word_length = sum(len(word) for word in words) / word_count if word_count > 0 else 0
    
    # Compute POSITIVE SCORE, NEGATIVE SCORE, POLARITY SCORE, and SUBJECTIVITY SCORE using TextBlob
    text_blob = TextBlob(article_text)
    polarity_score = text_blob.sentiment.polarity
    subjectivity_score = text_blob.sentiment.subjectivity
    
    # Compute PERCENTAGE OF COMPLEX WORDS
    complex_words = [word for word in words if textstat.syllable_count(word) > 2]  # Consider words with more than 2 syllables as complex
    percentage_complex_words = (len(complex_words) / word_count) * 100 if word_count > 0 else 0
    
    # Compute FOG INDEX
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    # Compute AVG NUMBER OF WORDS PER SENTENCE
    avg_words_per_sentence = word_count / sentence_count if sentence_count > 0 else 0
    
    # Count PERSONAL PRONOUNS
    personal_pronouns = ['I', 'me', 'my', 'mine', 'myself', 'we', 'us', 'our', 'ours', 'ourselves', 
                         'you', 'your', 'yours', 'yourself', 'yourselves', 'he', 'him', 'his', 'himself', 
                         'she', 'her', 'hers', 'herself', 'it', 'its', 'itself', 'they', 'them', 'their', 
                         'theirs', 'themselves']
    personal_pronoun_count = sum(article_text.lower().count(pronoun) for pronoun in personal_pronouns)
    
    # Compute COMPLEX WORD COUNT
    complex_word_count = len(complex_words)
    
    # Compute SYLLABLE PER WORD
    syllable_per_word = sum(textstat.syllable_count(word) for word in words) / word_count if word_count > 0 else 0
    
    # Return computed variables
    return {
        'WORD COUNT': word_count,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'AVG WORD LENGTH': avg_word_length,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'PERSONAL PRONOUNS': personal_pronoun_count,
        'COMPLEX WORD COUNT': complex_word_count,
        'SYLLABLE PER WORD': syllable_per_word
    }

# Initialize a list to store computed variables for each article
computed_variables_list = []

# Iterate over each row in input_data and perform text analysis
for index, row in input_data.iterrows():
    url_id = row['URL_ID']
    article_title = row['ARTICLE_TITLE']
    article_text = row['ARTICLE_TEXT']
    
    if article_text:
        # Perform text analysis and compute variables
        computed_variables = perform_text_analysis(article_text)
        computed_variables['URL_ID'] = url_id  # Add URL_ID to the computed variables
        
        # Append computed variables to the list
        computed_variables_list.append(computed_variables)

# Convert the list of dictionaries to a DataFrame
output_data = pd.DataFrame(computed_variables_list)

# Reorder columns to match the order in the output structure file
output_data = output_data[['URL_ID', 'WORD COUNT', 'AVG SENTENCE LENGTH', 'AVG WORD LENGTH',
                           'POLARITY SCORE', 'SUBJECTIVITY SCORE', 'PERCENTAGE OF COMPLEX WORDS',
                           'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'PERSONAL PRONOUNS',
                           'COMPLEX WORD COUNT', 'SYLLABLE PER WORD']]

# Write the computed variables to an Excel file
output_data.to_excel('Output.xlsx', index=False)

print("Text analysis complete. Output saved to Output.xlsx")
