from sklearn.feature_extraction.text import CountVectorizer
import pickle
import sys

model = pickle.load(open('model.pkl', 'rb'))
count_vect = CountVectorizer()
final_features = sys.argv[1]
prediction = model.predict(count_vect.fit_transform([final_features]))

if (prediction==1):
    prediction = 1
else:
    prediction = 0
sys.stdout.flush()