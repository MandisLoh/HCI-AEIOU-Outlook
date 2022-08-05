from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer 
from sklearn.decomposition import PCA
import pickle
import sys
import pandas as pd
import joblib
# kekw

print('helo')
print(str(sys.argv[1]))

model = joblib.load('pipeline.pkl')
text = str(sys.argv[1])
prediction = model.predict([text])
prob = model.predict_proba([text])
print(prediction, prob)

if (prediction==1):
    prediction = 'meeting'
else:
    prediction = 'normal'
print(prediction)
sys.stdout.flush()