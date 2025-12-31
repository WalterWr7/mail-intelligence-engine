import win32com.client
import pandas as pd
import re
import datetime
import sys

# --- ⚙️ CONFIGURACIÓN MASIVA ---
MI_EMAIL_CORPORATIVO = "wllana@unibanca.pe"
MI_NOMBRE_MOSTRAR = "Walter Llana"
DIAS_HISTORIAL = 365  # ¡EXTRAER 1 AÑO COMPLETO!
DIAS_PARA_IGNORADO = 7
SEPARADOR_CSV = "|"

# MAPI Tags
MAPI_LAST_VERB = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

def limpiar_texto(texto):
    """Limpieza profunda: Emojis, URLs y caracteres raros"""
    if not texto: return ""
    texto = str(texto)
    # 1. Reemplazar URLs por token
    texto = re.sub(r'http\S+', 'URL', texto)
    # 2. Dejar solo alfanuméricos y puntuación básica (Adiós emojis)
    texto = re.sub(r'[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ.,:;?!\s@\-_]', '', texto)
    # 3. Limpiar espacios y pipes
    texto = re.sub(r'[\n\r\t|]', ' ', texto)
    return re.sub(' +', ' ', texto).strip()

def obtener_info_remitente(item):
    email_final = "desconocido"
    nombre_final = "desconocido"
    dominio = "interno"
    try:
        nombre_final = item.SenderName
        direccion = item.SenderEmailAddress
        if direccion and "/o=" in direccion.lower():
            try:
                ex_user = item.Sender.GetExchangeUser()
                if ex_user: email_final = ex_user.PrimarySmtpAddress.lower()
                else: email_final = nombre_final.lower()
            except: email_final = nombre_final.lower()
        else:
            email_final = direccion.lower() if direccion else nombre_final.lower()
        
        if "@" in email_final: dominio = email_final.split("@")[1].strip()
        else: dominio = "unibanca.pe"
    except: pass
    return email_final, dominio, nombre_final

def analizar_audiencia(item):
    estoy_en_to = 0
    estoy_en_cc = 0
    total = 0
    try:
        recipients = item.Recipients
        total = recipients.Count
        mi_email = MI_EMAIL_CORPORATIVO.lower()
        mi_nombre = MI_NOMBRE_MOSTRAR.lower()
        
        # Analizamos primeros 50 destinatarios
        for i, r in enumerate(recipients):
            if i > 50: break 
            try:
                addr = r.Address.lower() if r.Address else ""
                name = r.Name.lower() if r.Name else ""
                soy_yo = (mi_email in addr) or (mi_nombre in name)
                
                if soy_yo:
                    if r.Type == 1: estoy_en_to = 1
                    elif r.Type == 2: estoy_en_cc = 1
            except: pass
    except: pass
    return estoy_en_to, estoy_en_cc, total

def verificar_accion_realizada(item):
    try:
        verb = item.PropertyAccessor.GetProperty(MAPI_LAST_VERB)
        if verb in [102, 103]: return 1 
        if verb == 104: return 2
    except: pass
    return 0

def calcular_ground_truth(item, accion_realizada):
    if accion_realizada > 0: return 2
    try:
        if item.UnRead:
            fecha = item.ReceivedTime.replace(tzinfo=None)
            if (datetime.datetime.now() - fecha).days >= DIAS_PARA_IGNORADO:
                return 0
    except: pass
    return 1

def procesar_carpeta_recursiva(carpeta, lista_datos, ruta_actual, fecha_limite):
    nombre_carpeta = carpeta.Name
    ruta_completa = f"{ruta_actual} > {nombre_carpeta}" if ruta_actual else nombre_carpeta
    
    print(f"📂 Escaneando: {ruta_completa} ...")
    
    try:
        items = carpeta.Items
        # Intentar ordenar (con protección)
        try:
            items.Sort("[ReceivedTime]", True) 
        except:
            print("   [WARN] No se pudo ordenar por fecha. Continuando sin orden...")

        # Procesados en esta carpeta
        local_count = 0
        
        for item in items:
            # OPTIMIZACIÓN: No leer todo, solo mails
            if item.Class != 43: continue
            
            try:
                # --- FILTRO DE FECHA (Time Travel) ---
                fecha_item = item.ReceivedTime.replace(tzinfo=None)
                
                # Si el correo es más antiguo que el límite, DEJAMOS DE LEER esta carpeta
                # (Como están ordenados, todos los siguientes serán más viejos)
                if fecha_item < fecha_limite:
                    break 

                email, dominio, nombre = obtener_info_remitente(item)
                en_to, en_cc, total_recip = analizar_audiencia(item)
                accion = verificar_accion_realizada(item)
                target = calcular_ground_truth(item, accion)
                
                lista_datos.append({
                    "Remitente_ID": email,
                    "Dominio": dominio,
                    "Nombre_Mostrar": limpiar_texto(nombre),
                    "Asunto": limpiar_texto(item.Subject),
                    "Cuerpo_Snippet": limpiar_texto(item.Body)[:500],
                    "Estoy_En_To": en_to,
                    "Estoy_En_CC": en_cc,
                    "Total_Destinatarios": total_recip,
                    "Carpeta_Origen": limpiar_texto(nombre_carpeta),
                    "Estado_Lectura": "No Leído" if item.UnRead else "Leído",
                    "Accion_Detectada": "Respondido" if accion==1 else ("Reenviado" if accion==2 else "Ninguna"),
                    "TARGET_IA": target
                })
                local_count += 1
                
                # Feedback visual cada 100 correos para que sepas que sigue vivo
                if local_count % 100 == 0:
                    print(f"   ... extraídos {local_count} correos de {nombre_carpeta}")

            except Exception as e: pass
            
        print(f"   ✅ Terminada carpeta {nombre_carpeta}: {local_count} registros.")

        # Recursividad
        for sub in carpeta.Folders:
            procesar_carpeta_recursiva(sub, lista_datos, ruta_completa, fecha_limite)
            
    except Exception as e:
        print(f"⚠️ Error carpeta {nombre_carpeta}: {e}")

def generar_dataset_masivo(dias=None):
    if dias is None: dias = DIAS_HISTORIAL
    
    print("--- 🚀 DATA MINING MASIVO ---")
    print(f"📅 Fecha límite: {(datetime.datetime.now() - datetime.timedelta(days=dias)).date()}")
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) 
    
    # Calcular fecha de corte
    fecha_limite = datetime.datetime.now() - datetime.timedelta(days=dias)
    
    datos_totales = []
    procesar_carpeta_recursiva(inbox, datos_totales, "", fecha_limite)
    
    df = pd.DataFrame(datos_totales)
    archivo = "dataset_masivo.csv"
    df.to_csv(archivo, index=False, sep=SEPARADOR_CSV, encoding='utf-8-sig')
    
    print(f"\n✅ Dataset generado: {archivo}")
    print(f"📊 Registros totales: {len(df)}")

if __name__ == "__main__":
    generar_dataset_masivo()