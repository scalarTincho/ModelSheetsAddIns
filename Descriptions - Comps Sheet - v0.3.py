#%%
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import xlwings as xw
from xlwings.constants import Direction
import collections
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from collections import OrderedDict
import pandas as pd

...

#Funtions    
#%%

# Get range from a start cell
def get_full_range(start_range):
    end_range = start_range.end(Direction.xlDown)

    # Create the full range using the start and end cell addresses
    full_range = start_range.get_address() + ':' + end_range.get_address()
    return full_range   


#%%
def get_unique_values_from_range(range_sheet_name,range_address):
    # Connect to the active workbook

    # Connect to the active sheet
    sheet = wb.sheets[range_sheet_name]
    
    #Get Industry Clasification Range

    # Read the specified range from the active sheet
    data_range = sheet.range(range_address).value

    # Get unique values using OrderedDict
    unique_values = list(OrderedDict.fromkeys(data_range))
    return unique_values


'''

PROGRAM STARTS HERE

'''
if __name__ == '__main__':
    
    #Get the active sheet & Workbook
    wb = xw.books.active
    ws = wb.sheets.active

    company_name = ws.range("C4").value    
    print(f"Running  Letter Soup for {company_name} \nPlease stand by...")

    desc_name = 'Comparables Analysis'
    main_sheet = 'Comparables Analysis'
    
    #Create a function to get the de unique industries
    #start_row = 15
    #end_row =  ws.range(f"C{start_row}").end(Direction.xlDown).row
    range_ind = get_full_range(wb.sheets[main_sheet].range('E37'))

    
    #define de range from company descriptions
    range_string = get_full_range(wb.sheets[desc_name].range('G37'))
    
    #load the range to a varaible
    description_range = wb.sheets[desc_name].range(range_string)
    
    # Define the parts you want to discard
    discard_tags = ['JJ', 'JJR', 'JJS', 'RB', 'RBR', 'RBS', 'IN', ',', 'Inc.', 'CC', ';']
    
    # Define the stop words you want to discard
    stop_words = set(stopwords.words('english')) | set(['et', 'al'])
    
    # Define the excluded words list
    excluded_words = ['company', ';', 'inc.', 'founded', 'incorporated', 'llc', '!',
                       'inc', ':', '(', ')','.', 'provides','offer','operates','headquartered','including']
    

    # Convert the range to a list of descriptions
    # descriptions = [cell.value for cell in description_range if cell.value is not None]
    descriptions = [cell.value for cell in description_range if cell.value is not None and cell.value != '(Invalid Identifier)']
    
    # Split the descriptions into individual words and tag them with their parts of speech
    lemmatizer = WordNetLemmatizer()
    tagged_descriptions = []
    
    for description in descriptions:
        words = nltk.word_tokenize(description.lower())
        tagged_words = nltk.pos_tag(words)
        

        filtered_words = [lemmatizer.lemmatize(word) for word, tag in tagged_words 
                          if tag not in discard_tags and word not in stop_words  
                          and not word.isdigit() and word not in excluded_words]
        
        tagged_descriptions.append(filtered_words)
    
    # Count the frequency of each word in the descriptions
    word_counts = collections.Counter()
    
    for description_words in tagged_descriptions:
        for word in description_words:
            word_counts[word] += 1
  
    #print(tagged_descriptions)
    
    #%%
    # Assuming word_counts is a Counter object
    # Get the top 10 most common words
    top_10_words = word_counts.most_common(10)
    
    # Create a DataFrame
    top_10_words_df = pd.DataFrame(top_10_words, columns=['Word', 'Frequency'])
        
    #%%
    #Get Unique Industries
    industries = get_unique_values_from_range(main_sheet,range_ind)
    industries = [industry for industry in industries if industry != '(Invalid Identifier)']

# Assuming industries is a list
    industries_df = pd.DataFrame(industries, columns=['Industry'])
    industries_df = industries_df[industries_df['Industry'] != '(Invalid Identifier)']
    # print(industries)
    
    wsResults = wb.sheets['Comparables Analysis']
        
    # Create a wordcloud of the most common words
    wordcloud = WordCloud(background_color='white',width=1200, height=600).generate_from_frequencies(word_counts)
    
    # Plot the wordcloud
    plt.figure(figsize=(10, 6.5))
    plt.imshow(wordcloud,interpolation='bilinear')
    plt.axis('off')
    
    plt.figtext(0.50, 0.02,"Industries: " + str(industries),horizontalalignment='left', wrap=True)
    
    plt.tight_layout(pad=0)
    # Save the plot as an image
    # plt.savefig('wordcloud_plot.png', dpi=300)
    
    '''
    Save imane in documents
    '''
    import os
    from os.path import expanduser
    
    downloads_folder = os.path.join(expanduser("~"), "Downloads")
    plt.savefig(os.path.join(downloads_folder, 'wordcloud_plot.png'), dpi=300)
 


            
    
