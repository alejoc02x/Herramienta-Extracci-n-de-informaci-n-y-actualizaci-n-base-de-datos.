import pandas as pd
import re
import nltk
import joblib
import matplotlib.pyplot as plt
import seaborn as sns
from nltk.corpus import stopwords
from nltk.stem import SnowballStemmer
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.svm import SVC
from sklearn.neural_network import MLPClassifier
from xgboost import XGBClassifier
from sklearn.metrics import classification_report, confusion_matrix

nltk.download('stopwords')

# --- 1. Cargar y limpiar datos ---
file_path = "/content/ALERTAS DISPOSITIVOS MÉDICOS 2024.xlsx"
xls = pd.ExcelFile(file_path)
df = xls.parse('ALERTAS 2024', skiprows=3)

df.columns = [
    'Mes', 'Fecha Emision', 'Codigo Fuente', 'Fuente', 'Tipo Alerta',
    'Dispositivo/Equipo', 'Tipo Dispositivo', 'Registro INVIMA', 'Imagen',
    'Descripcion Alerta', 'Responsable Verificacion', 'Medio Socializacion',
    'Aplicabilidad', 'Soporte'
]

df = df[['Dispositivo/Equipo', 'Tipo Dispositivo']].dropna()
df['Tipo Dispositivo'] = df['Tipo Dispositivo'].astype(str).str.strip().str.lower().str.replace('ó', 'o')
df = df[df['Tipo Dispositivo'] != 'tipo dispositivo']  # Eliminar fila de título accidental

# --- 2. Preprocesamiento de texto ---
spanish_stopwords = set(stopwords.words('spanish'))
stemmer = SnowballStemmer('spanish')

def limpiar_texto(texto):
    texto = re.sub(r'\W', ' ', texto)
    texto = re.sub(r'\d', '', texto)
    texto = texto.lower()
    palabras = texto.split()
    palabras = [stemmer.stem(p) for p in palabras if p not in spanish_stopwords]
    return ' '.join(palabras)

df['Dispositivo/Equipo'] = df['Dispositivo/Equipo'].astype(str).apply(limpiar_texto)

# --- 3. Dividir los datos ---
X = df['Dispositivo/Equipo']
y = df['Tipo Dispositivo']

from sklearn.preprocessing import LabelEncoder

# Codificar etiquetas
label_encoder = LabelEncoder()
y = label_encoder.fit_transform(df['Tipo Dispositivo'])


X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.35, random_state=42, stratify=y)

# --- 4. Vectorización ---
vectorizer = TfidfVectorizer(ngram_range=(1,2), max_features=10000, min_df=2, max_df=0.9)
X_train_vect = vectorizer.fit_transform(X_train)
X_test_vect = vectorizer.transform(X_test)

# --- 5. Definición de modelos y grids ---
modelos = {
    'RandomForest': (RandomForestClassifier(random_state=42), {
        'n_estimators': [100, 300],
        'max_depth': [None, 20],
        'min_samples_split': [2, 5]
    }),
    'XGBoost': (XGBClassifier(use_label_encoder=False, eval_metric='mlogloss', random_state=42), {
        'n_estimators': [100, 300],
        'max_depth': [3, 6]
    }),
    'KNN': (KNeighborsClassifier(), {
        'n_neighbors': [3, 5, 7],
        'weights': ['uniform', 'distance']
    }),
    'SVM': (SVC(probability=True), {
        'C': [0.1, 1, 10],
        'kernel': ['linear', 'rbf']
    }),
    'MLP': (MLPClassifier(max_iter=500, random_state=42), {
        'hidden_layer_sizes': [(100,), (50,50)],
        'activation': ['relu', 'tanh']
    }),
}

# --- 6. Entrenamiento y evaluación ---
for nombre, (modelo, grid) in modelos.items():
    print(f"\n📌 Entrenando modelo: {nombre}")
    grid_search = GridSearchCV(modelo, grid, cv=3, scoring='f1_weighted', n_jobs=-1)
    grid_search.fit(X_train_vect, y_train)
    best_model = grid_search.best_estimator_

    y_pred = best_model.predict(X_test_vect)
    print(f"🔎 Mejor configuración para {nombre}: {grid_search.best_params_}")
    from sklearn.metrics import ConfusionMatrixDisplay

# Decodificar predicciones y reales
    y_pred_labels = label_encoder.inverse_transform(y_pred)
    y_test_labels = label_encoder.inverse_transform(y_test)
    print(classification_report(y_test_labels, y_pred_labels))


    # Matriz de confusión
    cm = confusion_matrix(y_test, y_pred, labels=best_model.classes_)
    plt.figure(figsize=(6, 5))
    sns.heatmap(cm, annot=True, fmt="d", cmap="Blues", xticklabels=best_model.classes_, yticklabels=best_model.classes_)
    plt.title(f"Matriz de Confusión - {nombre}")
    plt.xlabel("Predicción")
    plt.ylabel("Real")
    plt.tight_layout()
    plt.show()

    # Guardar modelo y vectorizador
    joblib.dump(best_model, f'modelo_{nombre}.pkl')

# Guardar vectorizador
joblib.dump(vectorizer, 'vectorizador_dispositivos.pkl')

