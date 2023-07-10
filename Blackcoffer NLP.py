from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
import numpy as np
import traceback
from urllib.request import urlopen
import requests

links=pd.read_excel("C:/Input.xlsx")
links=links.set_index("URL_ID")

title={}
article={}
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0'}
for i in range(len(links.index)):
    page=requests.get(links.loc[links.index[i]][0],headers=headers)
    soup=BeautifulSoup(page.content,"html.parser")
    title[links.index[i]]=soup.find("span", {"class": "td-bred-no-url-last"}).text.strip().replace("\n"," ")
    article[links.index[i]]=soup.find("div", {"class": "td-post-content"}).text.strip().replace("\n"," ")

#DATA PROCESSING
from nltk.corpus import stopwords
from nltk.tokenize import RegexpTokenizer

import nltk
nltk.download('punkt')

# 1. STOP WORDS LOGIC
i = 0
stop_words = list(map(str.lower,set(stopwords.words('english'))))
article_token={}
tokenizer=RegexpTokenizer(r'\w+')
for i in article.keys():
    article_token[i] = list(map(str.lower,tokenizer.tokenize(article[i])))

x=[]
i = 0
article_filtered={}
for i in article_token.keys():
    for w in article_token[i]:
        if w not in stop_words:
            x.append(w)
    article_filtered[i]=x
    x=[]

#POSITIVE WORDS/NEGATIVE WORDS
Sent=pd.read_excel("C:/LoughranMcDonald_MasterDictionary_2020.xlsx")
Positive=list(Sent[Sent["Positive"]>1]["Word"])
Negative=list(Sent[Sent["Negative"]>1]["Word"])
i = 0
for i in Negative:
    if type(i)==bool:
        c=1
        Negative.remove(i)
if c==1:
    Negative.append("FALSE")

c=0
i = 0
Positive_count={}
for i in article_filtered:
    for j in list(map(str.lower,Positive)):
        c=c+article_filtered[i].count(j)
    Positive_count[i]=c
    c=0

c=0
i = 0
Negative_count={}
for i in article_filtered:
    for j in list(map(str.lower,Negative)):
        c=c+article_filtered[i].count(j)
    Negative_count[i]=c
    c=0

#POLARITY SCORE
#Polarity Score = (Positive Score – Negative Score)/ ((Positive Score + Negative Score) + 0.000001)
Polarity_Score={}
i = 0
for i in article_filtered:
    Polarity_Score[i] = (Positive_count[i]-Negative_count[i])/ ((Positive_count[i] + Negative_count[i]) + 0.000001)

#Subjectivity Score
#Subjectivity Score = (Positive Score + Negative Score)/ ((Total Words after cleaning) + 0.000001)
Subjectivity_Score={}
i = 0
for i in article_filtered:
    Subjectivity_Score[i] = (Positive_count[i]+Negative_count[i])/ ((len(article_filtered[i])) + 0.000001)


##Analysis of Readability

#Average Sentence Length = the number of words / the number of sentences
Average_Sentence_len={}
i = 0
for i in article_filtered:
    Average_Sentence_len[i]=len(article_filtered[i])/len(article[i].split("."))
#Percentage of Complex words = the number of complex words / the number of words
Complex={}
c=0
i = 0
j = 0
count=0
for i in article_filtered:
    for j in article_filtered[i]:
        for vowel in "aeiou":
            c=c+j.count(vowel)
        if c>2 and j[-2:]!="es" and j[-2:]!="ed":
            count=count+1
        c=0
    Complex[i]=count
    count=0
Complex_words_percent={}
for i in article_filtered:
    Complex_words_percent[i]=(Complex[i])/len(article_token[i])
#Fog Index = 0.4 * (Average Sentence Length + Percentage of Complex words)
fog_index = {}
for i in article_filtered:
    fog_index[i] = (Average_Sentence_len[i] + (Complex_words_percent[i]*100))*0.4

#AVG Number of Words per Sentence
#Average Number of Words Per Sentence = the total number of words / the total number of sentences
avg_number_of_words_per_sentence = Average_Sentence_len

#Complex Word Count
#Complex words are words in the text that contain more than two syllables.
complex_word_count = Complex

#Word Count
#We count the total cleaned words present in the text by
# 1.removing the stop words (using stopwords class of nltk package).
# 2.removing any punctuations like ? ! , . from the word before counting.
word_count = {}
count=0
for i in article_filtered:
    for j in article_filtered[i]:
            count = count+1
    word_count[i] = count
    count = 0

#SYLLABLE PER WORD
syllable_per_word ={}
total_syllables = {}
c=0
count = 0
for i in article_filtered:
    for j in article_filtered[i]:
        for vowel in "aeiou":
            c=c+j.count(vowel)
        if c>0 and j[-2:]!="es" and j[-2:]!="ed":
            count = count+c
        else:
            count = count-1
        c=0
    total_syllables[i]=count
    count=0
for i in article_filtered:
    syllable_per_word[i] = total_syllables[i]/word_count[i]

#PERSONAL PRONOUNS
#Find the counts of the words - “I,” “we,” “my,” “ours,” and “us”.
PP = {}
check=["I","WE","ME","OURS","us","i","We","we","Me","me","Ours","ours","Us"]
y=0
for i in article_filtered.keys():
    for x in check:
        y=y+article_filtered[i].count(x)
    PP[i]=y
    y=0

#AVG WORD LENGTH
#AVG WORD LENGTH = Sum of the total number of characters in each word/Total number of words
avg_word_len={}
number_char = {}
k = 0
for i in article_filtered:
    for j in article_filtered[i]:
        k = k+len(j)
    number_char[i] = k
    k=0
for i in article_filtered:
    avg_word_len[i] = number_char[i]/word_count[i]

#Combining the data into a DATAFRAME
data = {'URL ID' : links.index,
        'URL' : links["URL"].values,
       'POSITIVE NUMBER' : Positive_count.values(),
       'NEGATIVE NUMBER' : Negative_count.values(),
       'POLARITY SCORE' : Polarity_Score.values(),
       'SUBJECTIVITY SCORE' : Subjectivity_Score.values(),
       'AVG SENTENCE LENGTH' : Average_Sentence_len.values(),
       'PERCENTAGE OF COMPLEX WORDS' : Complex_words_percent.values(),
       'FOG INDEX' : fog_index.values(),
       'AVG NUMBER OF WORDS PER SENTENCE' : avg_number_of_words_per_sentence.values(),
       'COMPLEX WORD COUNT' : complex_word_count.values(),
       'WORD COUNT' : word_count.values(),
       'SYLLABLE PER WORD' : syllable_per_word.values(),
       'PERSONAL PRONOUNS' : PP.values(),
       'AVG WORD LENGTH' : avg_word_len.values()
       }
df = pd.DataFrame(data)
df.to_excel("C:/Blackcoffer NLP.xlsx", index = False)