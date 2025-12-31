import pandas as pd
import joblib
from catboost import CatBoostClassifier
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.base import BaseEstimator, ClassifierMixin
# --- NUEVAS LIBRERÍAS PARA MÉTRICAS ---
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score

# --- CONFIGURACIÓN ---
ARCHIVO_DATASET = "dataset_masivo.csv"
ARCHIVO_MODELO = "cerebro_priorizacion.joblib" 

# --- WRAPPER PARA CORREGIR ERROR DE SKLEARN 1.6 ---
class CatBoostWrapper(BaseEstimator, ClassifierMixin):
    def __init__(self, **kwargs):
        self.model = CatBoostClassifier(**kwargs)
    
    def fit(self, X, y):
        self.model.fit(X, y)
        self.classes_ = self.model.classes_
        return self
    
    def predict(self, X):
        return self.model.predict(X)
    
    def predict_proba(self, X):
        return self.model.predict_proba(X)
    
    def __sklearn_tags__(self):
        from sklearn.utils._tags import _safe_tags
        return _safe_tags(BaseEstimator(), key=None)

def entrenar_modelo_definitivo():
    print("--- 🐱 Entrenando el CEREBRO FINAL (CatBoost) ---")
    
    # 1. Cargar Datos
    try:
        df = pd.read_csv(ARCHIVO_DATASET, sep='|')
        df['Asunto'] = df['Asunto'].fillna("").astype(str)
        df['Dominio'] = df['Dominio'].fillna("desconocido")
        df = df.fillna(0)
        print(f"✅ Datos cargados: {len(df)} registros.")
    except Exception as e:
        print(f"❌ Error: {e}")
        return

    # 2. Preparar Target
    df['TARGET_BINARIO'] = df['TARGET_IA'].apply(lambda x: 1 if x == 2 else 0)
    
    X = df[['Asunto', 'Dominio', 'Estoy_En_To', 'Estoy_En_CC', 'Total_Destinatarios']]
    y = df['TARGET_BINARIO']

    # 3. Pipeline de Preprocesamiento
    preprocessor = ColumnTransformer(
        transformers=[
            ('txt', TfidfVectorizer(max_features=500, ngram_range=(1, 2)), 'Asunto'),
            ('cat', OneHotEncoder(handle_unknown='ignore'), ['Dominio']),
            ('num', StandardScaler(), ['Total_Destinatarios', 'Estoy_En_To', 'Estoy_En_CC'])
        ]
    )

    # 4. Definición del Modelo
    cat_model = CatBoostWrapper(
        iterations=300,
        depth=6,
        learning_rate=0.1,
        auto_class_weights='Balanced', 
        verbose=0
    )

    clf = Pipeline(steps=[('preprocessor', preprocessor), ('classifier', cat_model)])

    # --- 5. EVALUACIÓN DE RENDIMIENTO (Nuevo Bloque) ---
    print("\n--- 📊 Evaluando Métricas (Validación Cruzada 80/20) ---")
    
    # Separamos solo para ver qué tan bueno es (simulación de la realidad)
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )
    
    clf.fit(X_train, y_train)
    y_pred = clf.predict(X_test)
    
    # Reporte detallado
    print("\nREPORTE DE CLASIFICACIÓN:")
    print(classification_report(y_test, y_pred, target_names=['Normal (0)', 'Urgente (1)']))
    
    # Matriz de Confusión manual para claridad
    cm = confusion_matrix(y_test, y_pred)
    tn, fp, fn, tp = cm.ravel()
    print("MATRIZ DE CONFUSIÓN:")
    print(f"✅ Aciertos Normales: {tn}")
    print(f"🚨 Falsas Alarmas (Ruido): {fp}")
    print(f"⚠️ Urgentes Perdidos (Peligro): {fn}")
    print(f"🏆 Urgentes Detectados: {tp}")
    
    acc = accuracy_score(y_test, y_pred)
    print(f"\nExactitud Global (Accuracy): {acc:.2%}")
    print("-" * 40)

    # --- 6. ENTRENAMIENTO FINAL Y GUARDADO ---
    print("\n🧠 Re-entrenando con el 100% de la historia para producción...")
    # Ahora sí usamos TODO (X, y) para que el archivo guardado sea lo más potente posible
    clf.fit(X, y)

    joblib.dump(clf, ARCHIVO_MODELO)
    print(f"✅ ¡CEREBRO CATBOOST LISTO! Guardado en: {ARCHIVO_MODELO}")
    print("El modelo guardado ha aprendido de todos los datos disponibles.")

if __name__ == "__main__":
    entrenar_modelo_definitivo()