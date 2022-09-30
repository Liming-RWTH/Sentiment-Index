import sys, os
from win32com.client import Dispatch # for reading RTF format file
import re # for search the data in the news content
import pysentiment as ps # for sentiment analysis

root = "C:/Users/Admin/News/Newsdata/Altanta2011/"
abs_address_list = []
Harvard_Scores_List = [] # Harvard_Scores: Use Harvard IV-4 sentiment dictionaries
Loughran_Scores_List = [] # Loughran and McDonald Financial Sentiment Dictionaries
News_Data_List =[] # news report data
News_Title_List =[] 
News_Content_List =[]

for path, subdirs, files in os.walk(root):
    for name in files:
        abs_address_list.append(os.path.join(path+'/',name))
        News_Title_List.append(name)

for x in range(len(abs_address_list)):
    word = Dispatch('Word.Application') 
    word.Visible = 0 
    word.DisplayAlerts = 0
    news = abs_address_list[x]
    path = news.replace("/", "\\")
    doc = word.Documents.Open(FileName=path, Encoding='gbk')
    news_content = '\n'.join([para.Range.Text for para in doc.paragraphs])
    doc.Close()
    word.Quit()
    
    pattern = '(Load-Date:)(.*\n)'
    match = re.search(pattern, news_content) 
    news_data = "NULL"
    if match:
        news_data = match.group(2)
    News_Data_List.append(news_data)

    hiv4 = ps.HIV4()
    tokens = hiv4.tokenize(news_content)  
    hiv4_score = hiv4.get_score(tokens)
    Harvard_Scores_List.append(hiv4_score)

    lm = ps.LM()
    tokens2 = lm.tokenize(news_content)
    lm_score = lm.get_score(tokens2)
    Loughran_Scores_List.append(lm_score)
