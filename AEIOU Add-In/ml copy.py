from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer 
from sklearn.decomposition import PCA
import pickle
import sys
import pandas as pd
# kekw

print('helo')
print(str(sys.argv[1]))

model = pickle.load(open('model new.pkl', 'rb'))
count_vect = TfidfVectorizer(ngram_range=(1,2),
                     token_pattern=u'(?ui)\\b\\w*[a-z]+\\w*[a-z]+\\w*\\b')
final_features = str(sys.argv[1])
final_features = [final_features]
count_vect.fit(final_features)
count_vect_dtm = count_vect.transform(final_features)
count_vect_dense = pd.DataFrame(count_vect_dtm.toarray(), columns = count_vect.get_feature_names())
pca = PCA(n_components = 1) # put a reasonable number from graph (70%-90%)
count_vect_pca = pca.fit_transform(count_vect_dense)
prediction = model.predict(count_vect_pca)
print(count_vect_pca,
count_vect_dense,
count_vect_dtm,
)
#prediction = model.predict(count_vect.fit_transform(final_features))

if (prediction==2):
    prediction = 'meeting'
else:
    prediction = 'normal'
print(prediction)
sys.stdout.flush()