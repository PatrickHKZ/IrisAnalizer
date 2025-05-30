import pandas as pd
import nltk
import ssl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import SVC
from sklearn.pipeline import make_pipeline
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report
from nltk.corpus import stopwords
from nltk.stem import SnowballStemmer
import re
import json
from joblib import dump, load
import os

# Configuración inicial para evitar problemas SSL
try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context

# Descargar stopwords si no están disponibles
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

# Preprocesamiento de texto
stopwords_es = set(stopwords.words('spanish'))
stemmer = SnowballStemmer('spanish')

def preprocesar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.lower()
    texto = re.sub(r'[^\w\s]', '', texto)
    texto = " ".join([stemmer.stem(palabra) for palabra in texto.split() 
                     if palabra not in stopwords_es and len(palabra) > 2])
    return texto

# Cargar datos de entrenamiento
def cargar_datos_entrenamiento():
    with open('data/json/entrenamientoia.json', 'r', encoding='utf-8') as cat:
        datos = json.load(cat)
    return pd.DataFrame(datos)

# Entrenar y guardar el modelo
def entrenar_modelo():
    df = cargar_datos_entrenamiento()
    
    # Filtrar categorías con muy pocos ejemplos
    min_samples = 2
    value_counts = df['categoria'].value_counts()
    to_remove = value_counts[value_counts < min_samples].index
    df = df[~df['categoria'].isin(to_remove)]
    
    # Preprocesar
    df['descripcion'] = df['descripcion'].apply(preprocesar_texto)
    
    # Dividir datos
    X_train, X_test, y_train, y_test = train_test_split(
        df['descripcion'], 
        df['categoria'],
        test_size=0.2, 
        random_state=42
    )
    
    # Crear y entrenar modelo
    modelo = make_pipeline(
        TfidfVectorizer(ngram_range=(1, 2), max_features=5000),
        SVC(kernel='linear', class_weight='balanced', probability=True)
    )
    
    modelo.fit(X_train, y_train)
    
    # Evaluación
    y_pred = modelo.predict(X_test)
    print(classification_report(y_test, y_pred, zero_division=0))
    
    # Guardar modelo
    os.makedirs('data/modelos', exist_ok=True)
    dump(modelo, 'data/modelos/modelo_reclamos.joblib')
    return modelo

# Función para predecir categoría
def predecir_categoria(descripcion, modelo):
    descripcion_procesada = preprocesar_texto(descripcion)
    return modelo.predict([descripcion_procesada])[0]

# Función principal para categorizar reclamos
def categorizar_reclamos():
    # Cargar o entrenar modelo
    modelo_path = 'data/modelos/modelo_reclamos.joblib'
    if os.path.exists(modelo_path):
        modelo = load(modelo_path)
    else:
        modelo = entrenar_modelo()
    
    # Cargar reclamos a categorizar
    with open('data/json/inforeclamos.json', 'r', encoding='utf-8') as f:
        inforeclamos = json.load(f)
    
    # Procesar cada reclamo
    for reclamo in inforeclamos:
        if 'Comentario' in reclamo:
            comentario = reclamo['Comentario']
            categoria_predicha = predecir_categoria(comentario, modelo)
            reclamo['CategorIA'] = categoria_predicha
    
    # Guardar resultados
    with open('data/json/inforeclamos.json', 'w', encoding='utf-8') as f:
        json.dump(inforeclamos, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    categorizar_reclamos()