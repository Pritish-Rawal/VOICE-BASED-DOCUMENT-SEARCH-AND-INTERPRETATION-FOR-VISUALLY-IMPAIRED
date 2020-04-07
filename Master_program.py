# -*- coding: utf-8 -*-
"""
Created on Sat Apr 28 21:19:28 2018

@author: pritishrawal
"""

# -*- coding: utf-8 -*-
# import the necessary packages
# Importing Gensimno

import gensim
from gensim import corpora 
import speech_recognition as sr
import webbrowser
import requests

#import os
from nltk.tokenize import sent_tokenize,word_tokenize
from nltk.corpus import stopwords
from collections import defaultdict
from string import punctuation
from heapq import nlargest
from PIL import Image
from googlesearch import search 
import pytesseract
import cv2
import os
import numpy as np
from matplotlib import pyplot as plt 
import win32com.client as wincl
from autocorrect import spell
import sqlite3
from nltk.stem.wordnet import WordNetLemmatizer
import string

image_name=""
r = sr.Recognizer()
speak = wincl.Dispatch("SAPI.SpVoice")
print("What do you want to scan?")
speak.Speak("What do you want to scan?")

with sr.Microphone() as source:
    speak.Speak("Listening.....Say something!")
    print("Listening.....Say something!\n\n")
    r.energy_threshold += 280
    audio = r.listen(source)

# Speech recognition using Google Speech Recognition
try:
    # for testing purposes, we're just using the default API key
    # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
    # instead of `r.recognize_google(audio)`
    print("recognizing........\n")
    image_name=r.recognize_google(audio)
except sr.UnknownValueError:
    print("Google Speech Recognition could not understand audio")
except sr.RequestError as e:
    print("Could not request results from Google Speech Recognition service; {0}".format(e))


print("you said:",image_name)
speak.Speak("you said ")
speak.Speak(image_name)
image_name=image_name+".png"


pytesseract.pytesseract.tesseract_cmd = r'G:\Study\SEM 8\Tesseract-OCR\tesseract.exe'
#pytesseract.pytesseract.tesseract_cmd = r'C:\Users\320005936\AppData\Local\Continuum\anaconda3\Lib\site-packages\pytesseract'
 
# load the example image and convert it to grayscale
#change name of images
image = cv2.imread(image_name)
#image = cv2.imread(r"G:\Study\SEM 8\trial\os1.png")
#image = cv2.imread(r"G:\Study\SEM 8\trial\text4.png")
#image = cv2.imread(r"G:\Study\SEM 8\trial\prisha.png")
#image = cv2.imread(r"G:\Study\SEM 8\trial\os.png")


gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
 
# check to see if we should apply thresholding to preprocess the
# image
ret3,gray = cv2.threshold(gray,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)
#ret,thresh = cv2.threshold(gray,100,255,0)
# make a check to see if median blurring should be done to remove
# noise
#gray = cv2.medianBlur(gray, 3)
#kernel = np.ones((1,1),np.uint8)
#gray = cv2.dilate(gray,kernel,iterations = 1)
kernel = np.ones((2,2),np.uint8)
gray = cv2.erode(gray,kernel,iterations = 1)

fig=plt.figure()
fig.set_size_inches(9,9)
plt.imshow(gray,cmap="gray")
plt.show()


filename = "read_me.png"
cv2.imwrite(filename, gray)


# load the image as a PIL/Pillow image, apply OCR, and then delete
# the temporary file
text_read=""
text = pytesseract.image_to_string(cv2.imread("read_me.png"))
os.remove(filename)

#performing spell check of passage word by word
sents = sent_tokenize(text)
word_sent = [word_tokenize(s.lower()) for s in sents]
for num,a in enumerate(word_sent):
    for b in range(len(a)):
        if(len(word_sent)>=3):
            word_sent[num][b]=spell(word_sent[num][b])
        else:
            word_sent[num][b]=word_sent[num][b]
        #text_read=text_read+" "+word_sent[num][b]
#text=spell(text)

#printing the passage after performing spell check      
for sentence in sents:
    sentence=sentence.replace("\n"," ")
    for word in sentence.split(" "):
        if(len(word)>=3):
            print(spell(word),end=" ")
        else:
            print(word,end=" ")
    print(" ")



# show the output images
#cv2.imshow("Image", image)
#cv2.imshow("Output", gray)
#cv2.waitKey(0)



class FrequencySummarizer:
    def __init__(self, min_cut=0.1, max_cut=0.9):
        self._min_cut = min_cut
        self._max_cut = max_cut 
        self._stopwords = set(stopwords.words('english') + list(punctuation))

    def _compute_frequencies(self, word_sent):
        freq = defaultdict(int)
        for s in word_sent:
          for word in s:
            if word not in self._stopwords:
                freq[word] += 1
        m = float(max(freq.values()))
        if freq[word]==m:
            print(word)
        for w in list(freq.keys()):
            freq[w] = freq[w]/m
            if freq[w] >= self._max_cut or freq[w] <= self._min_cut:
                del freq[w]
        return freq    # return the frequency list

    def summarize(self, text, n,sents):
        
        # split the text into sentences
        assert n <= len(sents)
        word_sent = text
        self._freq = self._compute_frequencies(word_sent)
        ranking = defaultdict(int)
        for i,sent in enumerate(word_sent):
            for w in sent:
                # for each word in this sentence
                if w in self._freq:
                    # if this is not a stopword (common word), add the frequency of that word 
                    # to the weightage assigned to that sentence 
                    ranking[i] += self._freq[w]
        # OK - we are outside the for loop and now have rankings for all the sentences
        sents_idx = nlargest(n, ranking, key=ranking.get)
        # we want to return the first n sentences with highest ranking, use the nlargest function to do so
        # this function needs to know how to get the list of values to rank, so give it a function - simply the 
        # get method of the dictionary
        return [sents[j] for j in sents_idx]

#textOfUrl = text
fs = FrequencySummarizer()
#summary = fs.summarize(textOfUrl, 5,sents)
summary = fs.summarize(word_sent, 5,sents)

print("\n\n*************************** SUMMARY **************************")
count=1
for a in summary:
    print(count,". ",end=" ")
    count+=1
    a=a.replace("\n"," ")
    for b in a.split(" "):
        if(len(b)>=3):
        #print(b)
            print(spell(b),end=' ')
            
        else:
            print(b,end=" ")
    print("")



stop = set(stopwords.words('english'))
exclude = set(string.punctuation) 
lemma = WordNetLemmatizer()

#removing stop words, punctuations 
def clean(doc):
    stop_free = " ".join([i for i in doc.lower().split() if i not in stop])
    punc_free = ''.join(ch for ch in stop_free if ch not in exclude)
    normalized = " ".join(lemma.lemmatize(word) for word in punc_free.split())
    return normalized

final=[]

for a in word_sent:
    temp=""
    for b in a:
        temp=temp+" "+b
    final.append(temp[:-2])
    
doc_clean = [clean(doc).split() for doc in final]  

# Creating the term dictionary of our courpus, where every unique term is assigned an index. 
dictionary = corpora.Dictionary(doc_clean)

# Converting list of documents (corpus) into Document Term Matrix using dictionary prepared above.
doc_term_matrix = [dictionary.doc2bow(doc) for doc in doc_clean]

# Creating the object for LDA model using gensim library
Lda = gensim.models.ldamodel.LdaModel

# Running and Training LDA model on the document term matrix.
ldamodel = Lda(doc_term_matrix, num_topics=3, id2word = dictionary, passes=100)

lda=ldamodel.print_topics(num_topics=3, num_words=3)


#tts = gTTS(text=summary[0], lang='en')
#tts.save("good.mp3")
#os.system("mpg321 good.mp3")



speak.Speak("do you want to hear summary of the passage?  Yes  or  No   ")
speak.Speak("Enter your choice by typing or speak when asked")


choice=input("do you want to hear summary of the passage?\n1. Yes\n2. No\n\n")

with sr.Microphone() as source:
    speak.Speak("Listening.....Say something!")
    print("Listening.....Say something!\n")
    r.energy_threshold += 280
    audio = r.listen(source)

# Speech recognition using Google Speech Recognition
try:
    # for testing purposes, we're just using the default API key
    # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
    # instead of `r.recognize_google(audio)`
    print("recognizing........\n")
    sum_text=r.recognize_google(audio)
except sr.UnknownValueError:
    print("Google Speech Recognition could not understand audio\n")
except sr.RequestError as e:
    print("Could not request results from Google Speech Recognition service; {0}".format(e))

print("you said:",sum_text)
speak.Speak("you said ")
speak.Speak(sum_text)
if(choice=="1" or choice=="Yes" or choice=="yes" or sum_text=="yes" or sum_text=="Yes"):
    for a in summary:  
 
        speak.Speak(a)

print("\n\n*************************** TOPICS **************************")
lda_d=[]
print("the topics of passage are:\n\n")
speak.Speak("\n\nthe topics of passage are:")
for a in lda:  
    
    temp=a[1]
    te=temp.split("+")
    for a in te:
        tem=(a.split("*")[1])
        tem=tem.split("\"")[1]
        if(tem not in lda_d and len(tem)>1):
            lda_d.append(tem)
count=1
for a in lda_d:
    print(count,". ",a,end=" ")    
    count=count+1
    print("")
    speak.Speak(a)
    

speak.Speak("would you like more information on these topics?  Yes or  No")
speak.Speak("Enter your choice by typing or speak when asked")

choice=input("would you like more information on these topics? \n1. Yes \n2. No\n\n\n")

with sr.Microphone() as source:
    speak.Speak("Listening.....Say something!")
    print("Listening.....Say something!\n\n")
    r.energy_threshold += 280
    audio = r.listen(source)

# Speech recognition using Google Speech Recognition
try:
    # for testing purposes, we're just using the default API key
    # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
    # instead of `r.recognize_google(audio)`
    print("recognizing........\n")
    topic_choice_text=r.recognize_google(audio)
except sr.UnknownValueError:
    print("Google Speech Recognition could not understand audio")
except sr.RequestError as e:
    print("Could not request results from Google Speech Recognition service; {0}".format(e))

print("you said:",topic_choice_text)
speak.Speak("you said ")
speak.Speak(topic_choice_text)


if(choice=="1" or choice=="Yes" or choice=="yes" or topic_choice_text=="yes" or topic_choice_text=="Yes"):

    speak.Speak("which all topics would you like to search?(enter comma separated)")
    ch_par=input("which all topics would you like to search?(enter comma separated):\n")
    
    #voice input for searching
    speak.Speak("which all topics would you like to search?")
    print("which all topics would you like to search?")
    with sr.Microphone() as source:
        speak.Speak("Listening.....Say something!")
        print("Listening.....Say something!\n\n")
        r.energy_threshold += 280
        audio = r.listen(source)
        
    # Speech recognition using Google Speech Recognition
    try:
    # for testing purposes, we're just using the default API key
    # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
    # instead of `r.recognize_google(audio)`
        print("recognizing........\n")
        topics_text=r.recognize_google(audio)
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
    
    print("you said:",topics_text)
    speak.Speak("you said ")
    speak.Speak(topics_text)
    
    
    print()
    print("searching....\n\n")
    new_search=""
    for a in ch_par.split(","):
        new_search=new_search+" "+lda_d[int(a)-1]
    result=[]
    #add try catch block here 
    for j in search(new_search, tld="co.in", num=5, stop=1, pause=2):
        result.append(j)
    
    response = requests.get(result[0])
    #if(response.status_code==400):
    url = result[0] 
    # Open URL in a new tab, if a browser window is already open.
    #webbrowser.open_new_tab(url)
    
    # Open URL in new window, raising the window if possible.
    webbrowser.open_new(url)
    
    
    #searching by voice
    new_search_speech=topics_text
    

    result=[]
    #add try catch block here 
    for j in search(new_search_speech, tld="co.in", num=5, stop=1, pause=2):
        result.append(j)
    
    response = requests.get(result[0])
    #if(response.status_code==400):
    url = result[0] 
    # Open URL in a new tab, if a browser window is already open.
    #webbrowser.open_new_tab(url)
    
    # Open URL in new window, raising the window if possible.
    webbrowser.open_new(url)

#opening firefox database
data_path = r"C:\Users\pritishrawal\AppData\Roaming\Mozilla\Firefox\Profiles\kdo4gmo4.default"
files = os.listdir(data_path)
history_db = os.path.join(data_path, 'places.sqlite')

c = sqlite3.connect(history_db)
cursor = c.cursor()

select_statement = "select moz_places.url, moz_places.visit_count from moz_places;"
cursor.execute(select_statement)
#fetching database
results = cursor.fetchall()

#sorting the tuple according to hits
def last(n):
    return n[-1]  

def sort(tuples):
    return sorted(tuples, key=last)

new_res = sort(results)
new_res.reverse()

#display most relevant browsing history sorted in decreasing order of hits 
speak.Speak("Following are the most relevant sites you visited based on the topics you selected")
print("Following are the most relevant sites you visited based on the topics you selected")
for sites in new_res:
    for topic in lda_d:
        if(topic in sites[0].lower()):
            print(topic, "is in site ",sites)
        
