# -*- coding: utf-8 -*-
"""
Created on Thu Jun  8 10:49:41 2023

@author: Martin56
"""
#%%
import time
import pandas as pd
import xlwings as xw
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from xlwings.constants import Direction

# Define a function to get the full range of cells in a column
def get_full_range(start_range):
    end_range = start_range.end(Direction.xlDown)
    full_range = start_range.get_address() + ':' + end_range.get_address()
    return full_range

# Define the tokenization and lemmatization function for the text
def tokenize_and_lemmatize(text):
    tokens = nltk.word_tokenize(text)
    tokens = [t.lower() for t in tokens if t.isalnum()]
    tokens = [t for t in tokens if t not in stop_words]
    lemmas = [lemmatizer.lemmatize(t) for t in tokens]
    return lemmas

if __name__ == "__main__":
    start_time = time.time()

    # Get the active workbook and sheet
    wb = xw.books.active
    ws = wb.sheets['Comparables Analysis']
    ws_screening = wb.sheets['Screening']
    company_name = ws.range("C4").value
    
    print(f"Running Similarity Comps for {company_name} \nPlease stand by...")
    
    #Company Description
    reference_company_description = ws.range('cDescription').value

    # Create a DataFrame from the active sheet's data
    full_df = ws_screening.range('A1').expand().options(pd.DataFrame, index=False, header=True).value
    full_df = full_df[(full_df['Long Business Description']) != 0]


    # Check which industry or sector switch is set
    industry_sector_switch = ws.range('IS_Switch').value

    # Define the criteria based on the switch value
    if industry_sector_switch == 2:
        industry_sector_criteria = 'Industry'
        criteria_value = ws.range('C9').value
    elif industry_sector_switch == 1:
        industry_sector_criteria = 'Sector'
        criteria_value = ws.range('C10').value
    elif industry_sector_switch == 3:
        industry_sector_criteria = None
        criteria_value = None

    # Filter the DataFrame based on the selected criteria
    if industry_sector_switch==3:
        full_df_filtered = full_df
        new_df = full_df_filtered.drop(columns=full_df_filtered.columns[[0,1,2,3]])
        description_df = new_df.rename(columns={new_df.columns[0]:'description'})

    else:
        full_df_filtered = full_df[(full_df[industry_sector_criteria]==criteria_value)]
        new_df = full_df_filtered.drop(columns=full_df_filtered.columns[[0,1,2,3]])
        description_df = new_df.rename(columns={new_df.columns[0]:'description'})

    # Add the reference description to the DataFrame
    ref_desc_dict = {'description': reference_company_description}
    ref_desc_df = pd.DataFrame([ref_desc_dict])
    description_df = pd.concat([description_df, ref_desc_df], ignore_index=True)

    # Instantiate the lemmatizer and define the stop words
    lemmatizer = WordNetLemmatizer()
    stop_words = set(stopwords.words('english'))

    # Vectorize the descriptions using TfidfVectorizer
    vectorizer = TfidfVectorizer(tokenizer=tokenize_and_lemmatize, token_pattern=None)
    X = vectorizer.fit_transform(description_df['description'].tolist())

    # Compute the cosine similarity matrix
    similarity_matrix = cosine_similarity(X)

    # Create a DataFrame containing the similarity scores
    similarity_scores = {'Description': description_df.iloc[:-1]['description'].tolist(), 'Similarity': similarity_matrix[:-1, -1]}
    df = pd.DataFrame(similarity_scores)

    # Sort the DataFrame based on similarity scores
    df_sorted = df.sort_values(by='Similarity', ascending=False).reset_index(drop=True)

    # Find the index of the most similar description
    most_similar_index = -1
    max_similarity = 0.0
    for i in range(similarity_matrix.shape[0] - 1): 
        if similarity_matrix[i, -1] > max_similarity:
            max_similarity = similarity_matrix[i, -1]
            most_similar_index = i

 
    # Combine the original filtered DataFrame with the similarity scores
    df2 = df
    full_df_filtered = full_df_filtered.reset_index(drop=True)
    final_df = pd.concat([full_df_filtered, df2], axis=1)
    final_df_filtered = final_df[(final_df['Similarity'] != 0) & (final_df['Similarity'] != 1)]

    # Remove unnecessary columns and sort the final DataFrame by similarity
    final_df_filtered = final_df_filtered.drop(columns=final_df_filtered.columns[[3, 4]])
    final_df_sorted = final_df_filtered.sort_values(by='Similarity', ascending=False)
    
    # Get the top 10 most similar companies
    top_10 = final_df_sorted.head(15)
    top_10 = top_10[['Similarity','AI-Template']]
    ws.range('B37').options(index=False, header=False).value = top_10

    # Print the elapsed time information
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Elapsed time: {elapsed_time:.2f} seconds for a {len(df)} Dataframe")