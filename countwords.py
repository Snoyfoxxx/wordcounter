import nltk
import string
from nltk.corpus import stopwords # Import the stop word list
import docx2txt
import os
from docx import Document
#empty list to hold the path to all your docx files
document_list = []

#erase the contents of extracted.txt to have a blank slate
print 'erasing extracted.txt'
open('extracted.txt', 'w').close()

# make a list of all the paths to the docxfiles in a folder with subfolders

for path, subdirs, files in os.walk(r"C:\Users\helder.opdebeeck\Desktop\docx"):
    for name in files:
        if os.path.splitext(os.path.join(path, name))[1] == ".docx":
            document_list.append(os.path.join(path, name))
                                 
#extract the text from all the docx files (defined by the document_path) ,
            #code them to utf-8 and past all that text in a one giant txt file



for document_path in document_list:
    print 'extracting text from ' + document_path
    text = docx2txt.process(document_path)
    text = text.encode('utf-8')
    a = open('extracted.txt', 'a')
    a.write(text)
    a.close()


# Open the input file and read()
s = open('extracted.txt','r').read()
# returns a translation table that maps each character in the intabstring
#into the character at the same position in the outtab string.
table = string.maketrans("","")
 
# returns a copy of the string in which all characters have been translated
#using table and a list of characters to be removed from the source string.
s = s.translate(table, string.punctuation)
#break the text strings into tokens of words with NLTK 
words = nltk.word_tokenize(s)
# Remove single-character tokens (mostly punctuation)
words = [word for word in words if len(word) > 1]



# Remove numbers
#words = [word for word in words if not word.isnumeric()]
# Lowercase all words (default_stopwords are lowercase too)
words = [word.lower() for word in words]
# creating a set of stopwords to make searching faster
stop_words_dutch = set(nltk.corpus.stopwords.words('dutch'))
stop_words_eng = set(nltk.corpus.stopwords.words('english'))
#create a set of words from a file with a new irrelevant word on each line"
extraWords = open('extrawords_dutch_english_a.txt', 'r')
extra_words = set(extraWords)


 
# Remove stopwords
words = [word for word in words if word not in stop_words_dutch and word not in stop_words_eng and word not in extra_words]

# list removed stopwords

#exwords = [word for word in words if word in stop_words_dutch_english]


# Calculate frequency distribution
print 'calculating frequency'
fdist = nltk.FreqDist(words)


 
#Assign the number of words you want to see
n = raw_input("How many common words do you wish to see?")
 
# Output top n words and format to utf8
for word, frequency in fdist.most_common(int(n)):
    print(u'{} - {}'.encode('utf-8').format(word, frequency))

 



