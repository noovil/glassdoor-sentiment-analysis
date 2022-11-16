#Data processing packages
import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)
from sklearn.metrics import mean_squared_error
from math import nan, isnan
import random
#pd.set_option('display.max_colwidth', 300)

#NLP packages
from textblob import TextBlob

#helper functions
def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

def containKeyword(text, keyword_list):
    for keyword in keyword_list:
        if isfloat(keyword) and isnan(keyword):
            continue
        index = text.find(keyword)
        
        if index > -1:
            if text.find("the composition of a normal") > -1:
                print("extreme cons keyword ", keyword)
            return True
    return False

def category(sentiment):
    category_predicted = []
    for i in sentiment:
        if i >= 5:
            category_predicted.append('strongly positive')
        elif 0 < i < 5:
            category_predicted.append('mildly positive')
        elif i == 0:
            category_predicted.append('neutral')
        elif -5 < i < 0:
            category_predicted.append('mildly negative')
        else:
            category_predicted.append('strongly negative')
    return category_predicted
                

def analyse(text, polarity, subjectivity, col=None):
    try:
        analysis =TextBlob(text).lower().replace('\n', ' ').replace('\r', '')
        model_sentiment = analysis.sentiment.polarity
        subjectivity.append(analysis.sentiment.subjectivity)
        word_count = len(text.split())

        if col == "Cons":

            found_extreme_bad_keyword = containKeyword(analysis, extreme_bad_keywords)
            if found_extreme_bad_keyword:
                if abs(model_sentiment) > 0.5:
                    predicted_sentiment = -10
                else:
                    predicted_sentiment = -9
                polarity.append(predicted_sentiment)
                return

            found_strong_bad_keyword = containKeyword(analysis, strong_bad_keywords)
            if found_strong_bad_keyword:
                if abs(model_sentiment) > 0.7:
                    predicted_sentiment = -8
                elif abs(model_sentiment) > 0.4:
                    predicted_sentiment = -7
                else:
                    predicted_sentiment = -6
                polarity.append(predicted_sentiment)
                return
            found_mild_bad_keyword = containKeyword(analysis, mild_bad_keywords)
            if found_mild_bad_keyword:
                if abs(model_sentiment) > 0.7:
                    predicted_sentiment = -5
                elif abs(model_sentiment) > 0.4:
                    predicted_sentiment = -4
                else:
                    predicted_sentiment = -3
                if word_count >= 100:
                    predicted_sentiment -= 2
                polarity.append(predicted_sentiment)
                return
            
            if abs(model_sentiment) >= 0.7:
                    predicted_sentiment = -2
            elif abs(model_sentiment) >= 0.3:
                predicted_sentiment = -1
            else:
                predicted_sentiment = 0
            polarity.append(predicted_sentiment)
            return

        if col == "Pros":

            found_extreme_good_keyword = containKeyword(analysis, extreme_good_keywords)
            if found_extreme_good_keyword:
                if abs(model_sentiment) > 0.5:
                    predicted_sentiment = 10
                else:
                    predicted_sentiment = 9
                polarity.append(predicted_sentiment)
                return
            found_strong_good_keyword = containKeyword(analysis, strong_good_keywords)
            if found_strong_good_keyword:
                if abs(model_sentiment) > 0.7:
                    predicted_sentiment = 8
                elif abs(model_sentiment) > 0.4:
                    predicted_sentiment = 7
                else:
                    predicted_sentiment = 6
                polarity.append(predicted_sentiment)
                return
        
            found_mild_good_keyword = containKeyword(analysis, mild_good_keywords)
            if found_mild_good_keyword:
                if abs(model_sentiment) >= 0.7:
                    predicted_sentiment = 5
                elif abs(model_sentiment) > 0.4:
                    predicted_sentiment = 4
                else:
                    predicted_sentiment = 3
                polarity.append(predicted_sentiment)
                return

            if abs(model_sentiment) > 0.7:
                    predicted_sentiment = 2
            elif abs(model_sentiment) >= 0.3:
                predicted_sentiment = 1
            else:
                predicted_sentiment = 0
            polarity.append(predicted_sentiment)
            return
            
    except Exception as e:
        print(e)
        polarity.append(0)
        subjectivity.append(0)

#data cleaning
data = pd.read_excel('C:/Users/Li Voon/OneDrive/sentiment_analysis/files/ASX100 Glassdoor_Reporting Department - 300 comments.xlsx')
cols = [0,1,5,6,7,8,9]
data = data[data.columns[cols]]
col_names = list(data.columns)
col_names[4] = "Pros rating"
col_names[5] = "Cons rating"
data.columns = col_names
data.fillna("None", inplace=True)
#print(df.head())

keywords = pd.read_excel('C:/Users/Li Voon/OneDrive/sentiment_analysis/files/keywords_updated.xlsx')
extreme_good_keywords = keywords['extreme_pros']
strong_good_keywords = keywords['strong_pros']
mild_good_keywords = keywords['mild_pros']
extreme_bad_keywords = keywords['extreme_cons']
strong_bad_keywords = keywords['strong_cons']
mild_bad_keywords = keywords['mild_cons']
#print(strong_good_keywords)
#print(mild_bad_keywords)

#Calculating the Sentiment Polarity
polarity_pros=[]
subjectivity_pros=[]
sentiment_pros=[] 
polarity_cons=[] 
subjectivity_cons=[] 
sentiment_cons=[] 
for i in data['Pros'].values:
    analyse(i,polarity_pros, subjectivity_pros,'Pros')
for i in data['Cons'].values:
    analyse(i,polarity_cons, subjectivity_cons, 'Cons')

#Adding the Sentiment Polarity column to the data
#data['polarity_pros']=polarity_pros
#data['subjectivity_pros']=subjectivity_pros
data['predicted_pros'] = polarity_pros
#data['polarity_cons']=polarity_cons
#data['subjectivity_cons']=subjectivity_cons
data['predicted_cons'] = polarity_cons

combined_sentiment = []
for i in range(len(polarity_cons)):
    combined_sentiment.append(polarity_cons[i] + polarity_pros[i])
    
    
data['predicted_combined_rating'] = combined_sentiment
data['difference'] = combined_sentiment - data['Overall']
#data['MSE'] = np.sqrt(mean_squared_error(data['combined_sentiment'], data['Overall']))
#data['category'] = category(data['Overall'])
#data['predicted_category'] = category(combined_sentiment)
'''
count = 0
for i in range(len(data['predicted_category'])):
    if data['predicted_category'][i] == data['category'][i]:
        count += 1
print(count)
'''

print('MSE: ', np.sqrt(mean_squared_error(combined_sentiment, data['Overall'])))
#print(f'categorical accuracy: {count}/{len(combined_sentiment)}',  )
data.to_excel('C:/Users/Li Voon/OneDrive/sentiment_analysis/files/sentiment_analysis_31.xlsx')



#best so far: 30

