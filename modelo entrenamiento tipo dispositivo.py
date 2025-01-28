import pandas as pd
import re
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, confusion_matrix
import matplotlib.pyplot as plt
import seaborn as sns
import joblib

# 1. Cargar el archivo Excel
file_path = "D:/alejo/MINERIADATOS/ALERTAS DISPOSITIVOS MÉDICOS 2024.xlsx"  # Cambia esta ruta al archivo en tu entorno
data = pd.ExcelFile(file_path)

# Leer la hoja relevante y limpiar las filas irrelevantes
df = data.parse('ALERTAS 2024', skiprows=3)

# Renombrar las columnas para facilitar el manejo
df.columns = [
    'Mes', 'Fecha Emision', 'Codigo Fuente', 'Fuente', 'Tipo Alerta',
    'Dispositivo/Equipo', 'Tipo Dispositivo', 'Registro INVIMA', 'Imagen',
    'Descripcion Alerta', 'Responsable Verificacion', 'Medio Socializacion',
    'Aplicabilidad', 'Soporte'
]

# Seleccionar solo las columnas necesarias y eliminar filas vacías
df = df[['Dispositivo/Equipo', 'Tipo Dispositivo']].dropna()

# 2. Limpieza del texto
from nltk.corpus import stopwords
from nltk.stem import SnowballStemmer
import nltk
nltk.download('stopwords')

spanish_stopwords = set(stopwords.words('spanish'))
stemmer = SnowballStemmer('spanish')

def limpiar_texto(texto):
    texto = re.sub(r'\W', ' ', texto)  # Elimina caracteres especiales
    texto = re.sub(r'\d', '', texto)  # Elimina dígitos
    texto = texto.lower()             # Convierte a minúsculas
    palabras = texto.split()
    palabras = [stemmer.stem(palabra) for palabra in palabras if palabra not in spanish_stopwords]
    return ' '.join(palabras)

df['Dispositivo/Equipo'] = df['Dispositivo/Equipo'].apply(limpiar_texto)

# Dividir los datos en entrenamiento y prueba
X = df['Dispositivo/Equipo']  # Características
y = df['Tipo Dispositivo']    # Etiquetas

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.35, random_state=42)

# 3. Vectorización del texto
vectorizer = TfidfVectorizer(
    ngram_range=(1, 2),
    max_features=10000,
    min_df=2,
    max_df=0.9
)
X_train_vect = vectorizer.fit_transform(X_train)
X_test_vect = vectorizer.transform(X_test)

# 4. Búsqueda de hiperparámetros para Random Forest
param_grid = {
    'n_estimators': [100, 300, 500],
    'max_depth': [None, 10, 20, 30],
    'min_samples_split': [2, 5, 10],
    'min_samples_leaf': [1, 2, 4]
}

grid_search = GridSearchCV(
    RandomForestClassifier(random_state=42),
    param_grid,
    cv=3,
    scoring='f1_weighted',
    n_jobs=-1
)

# Entrenar el modelo con la búsqueda de hiperparámetros
grid_search.fit(X_train_vect, y_train)
best_model = grid_search.best_estimator_

# 5. Evaluar el modelo
y_pred = best_model.predict(X_test_vect)
print("Reporte de clasificación:\n", classification_report(y_test, y_pred))

# Mostrar la matriz de confusión
cm = confusion_matrix(y_test, y_pred)
plt.figure(figsize=(8, 6))
sns.heatmap(cm, annot=True, fmt="d", cmap="Blues", xticklabels=best_model.classes_, yticklabels=best_model.classes_)
plt.xlabel('Predicción')
plt.ylabel('Real')
plt.title('Matriz de Confusión')
plt.show()

# 6. Guardar el modelo y el vectorizador
joblib.dump(best_model, 'modelo_dispositivos.pkl')          # Guarda el modelo entrenado
joblib.dump(vectorizer, 'vectorizador_dispositivos.pkl')  # Guarda el vectorizador

# 7. Función para predecir nuevos tipos de dispositivos
def predecir_tipo_dispositivo(texto):
    texto_limpio = limpiar_texto(texto)
    texto_vect = vectorizer.transform([texto_limpio])
    return best_model.predict(texto_vect)[0]

# Ejemplo de predicción
nuevo_dispositivo = "Monitor cardíaco de última generación"
prediccion = predecir_tipo_dispositivo(nuevo_dispositivo)
print(f"Predicción para '{nuevo_dispositivo}': {prediccion}")
