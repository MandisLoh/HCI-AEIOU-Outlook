import sys
import joblib
import os
# kekw


def readinput():
    print('helo')
    print(str(sys.argv[1]))

    text = str(sys.argv[1])
    return text

def main():
    text = readinput()
    model = joblib.load('pipeline.pkl')
    prediction = model.predict([text])
    prob = model.predict_proba([text])
    print(prediction, prob)

    if (prediction==1):
        prediction = 'meeting'
    else:
        prediction = 'normal'
    print(prediction)
    sys.stdout.flush()
    return prediction

if __name__ == '__main__':
    main()