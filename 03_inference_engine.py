import win32com.client
import pandas as pd
import joblib
import re
from sklearn.base import BaseEstimator, ClassifierMixin
from catboost import CatBoostClassifier # Necesario para que reconozca el objeto

# --- ⚙️ CONFIGURACIÓN ---
ARCHIVO_MODELO = "cerebro_priorizacion.joblib"
MI_EMAIL = "wllana@unibanca.pe"
MI_NOMBRE = "Walter Llana"
UMBRAL_ROJO = 0.75
UMBRAL_AMARILLO = 0.60

# --- 🧠 CLASE WRAPPER (CRÍTICO: DEBE ESTAR AQUÍ PARA PODER CARGAR EL MODELO) ---
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
# ---------------------------------------------------------------------------

def inicializar_categorias(outlook_app):
    """Garantiza que existan las etiquetas de color"""
    print("--- 🎨 Verificando Categorías en Outlook ---")
    categories = outlook_app.Session.Categories
    try: categories.Item("IA Urgente")
    except: 
        print("🛠️ Creando categoría 'IA Urgente'...")
        categories.Add("IA Urgente", 1) # 1 = Rojo
    
    try: categories.Item("IA Revisar")
    except: 
        print("🛠️ Creando categoría 'IA Revisar'...")
        categories.Add("IA Revisar", 2) # 2 = Naranja

def limpiar_texto(texto):
    if not texto: return ""
    texto = str(texto)
    texto = re.sub(r'http\S+', 'URL', texto)
    texto = re.sub(r'[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ.,:;?!\s@\-_]', '', texto)
    return re.sub(' +', ' ', re.sub(r'[\n\r\t|]', ' ', texto)).strip()

def obtener_features(item):
    """Extrae toda la data necesaria para la IA"""
    email, dominio = "desconocido", "interno"
    try:
        # Remitente
        if item.SenderEmailAddress and "/o=" in item.SenderEmailAddress.lower():
            try: email = item.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
            except: email = item.SenderName.lower()
        else:
            email = item.SenderEmailAddress.lower() if item.SenderEmailAddress else item.SenderName.lower()
        
        if "@" in email: dominio = email.split("@")[1].strip()
        else: dominio = "unibanca.pe"
    except: pass

    # Audiencia
    en_to, en_cc, total = 0, 0, 0
    try:
        recipients = item.Recipients
        total = recipients.Count
        mi_id = MI_EMAIL.lower()
        mi_name = MI_NOMBRE.lower()
        for i, r in enumerate(recipients):
            if i > 50: break
            try:
                addr = r.Address.lower() if r.Address else ""
                nm = r.Name.lower() if r.Name else ""
                if mi_id in addr or mi_name in nm:
                    if r.Type == 1: en_to = 1
                    elif r.Type == 2: en_cc = 1
            except: pass
    except: pass

    return email, dominio, en_to, en_cc, total

def ejecutar_vigilancia():
    print("--- 👁️ INICIANDO VIGILANCIA IA (CatBoost Engine) ---")
    
    try:
        # Ahora sí funcionará porque Python conoce la clase CatBoostWrapper
        clf = joblib.load(ARCHIVO_MODELO)
        print("✅ Cerebro cargado correctamente.")
    except Exception as e:
        print(f"❌ Error cargando modelo: {e}")
        print("💡 Consejo: Verifica que el archivo .joblib esté en la misma carpeta.")
        return

    outlook_app = win32com.client.Dispatch("Outlook.Application")
    inicializar_categorias(outlook_app)
    
    inbox = outlook_app.GetNamespace("MAPI").GetDefaultFolder(6)
    items = inbox.Items.Restrict("[UnRead] = True")
    items.Sort("[ReceivedTime]", True)
    
    print(f"📩 Analizando {items.Count} correos no leídos...")
    
    count = 0
    for item in items:
        if item.Class != 43: continue
        try:
            email, dom, to, cc, tot = obtener_features(item)
            asunto = limpiar_texto(item.Subject)
            
            # Crear DataFrame con nombres de columnas exactos
            df = pd.DataFrame([{
                'Asunto': asunto, 
                'Dominio': dom,
                'Estoy_En_To': to, 
                'Estoy_En_CC': cc, 
                'Total_Destinatarios': tot
            }])
            
            prob = clf.predict_proba(df)[:, 1][0]
            
            accion = ""
            if prob >= UMBRAL_ROJO:
                item.Categories = "IA Urgente"
                item.Save()
                accion = f"🔴 [URGENTE {prob:.0%}]"
            elif prob >= UMBRAL_AMARILLO:
                item.Categories = "IA Revisar"
                item.Save()
                accion = f"🟡 [REVISAR {prob:.0%}]"
            
            if accion:
                print(f"{accion} {asunto[:40]}...")
            count += 1
                
        except Exception as e: 
            # print(f"Error en un item: {e}") 
            pass

    print(f"✅ Vigilancia terminada. {count} correos escaneados.")

if __name__ == "__main__":
    ejecutar_vigilancia()