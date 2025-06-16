# modules/ml_model.py

import joblib
import pandas as pd

model = joblib.load("ticket_classifier.joblib")

def predict_ticket_priority(subject, body, type_, queue, priority, language):
    text = f"{subject} {body} {type_} {queue} {priority} {language}"
    df = pd.DataFrame({"text": [text]})
    prediction = model.predict(df["text"])
    return prediction[0]
