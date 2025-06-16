import pandas as pd
import joblib
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.svm import LinearSVC
from sklearn.pipeline import Pipeline
from sklearn.metrics import classification_report

# 1. Charger le fichier CSV
df = pd.read_csv(r"C:\Users\mohammed\Desktop\2\ai-assistant-sdm\dataset-tickets-multi-lang-4-20k.csv", encoding="utf-8")


# 2. Nettoyage des colonnes texte
df.fillna('', inplace=True)
df["text"] = df["subject"] + " " + df["body"] + " " + df["type"] + " " + df["queue"] + " " + df["priority"] + " " + df["language"]

# Ajout des tags s'ils existent
for i in range(1, 9):
    col = f"tag_{i}"
    if col in df.columns:
        df["text"] += " " + df[col]

# 3. Préparation des labels
df['priority'] = df['priority'].map({
    "high": "P3",
    "medium": "P2",
    "low": "P1"
})


X = df["text"]
y = df["priority"]

# 4. Découpage entraînement/test
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 5. Pipeline TF-IDF + SVM
model = Pipeline([
    ("tfidf", TfidfVectorizer()),
    ("clf", LinearSVC())
])

# 6. Entraînement
model.fit(X_train, y_train)

# 7. Évaluation
y_pred = model.predict(X_test)
print(classification_report(y_test, y_pred))

# 8. Sauvegarde du modèle
joblib.dump(model, "ticket_classifier.joblib")
print("✅ Modèle entraîné et sauvegardé sous 'ticket_classifier.joblib'")
