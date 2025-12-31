# 📧 Mail Intelligence Engine (V3)

> **Sistema de Priorización Predictiva para Microsoft Outlook con Interfaz Gráfica.**
> *Optimiza tu flujo de trabajo mediante Inteligencia Artificial Local.*

![Python](https://img.shields.io/badge/Python-3.12+-blue?logo=python)
![CatBoost](https://img.shields.io/badge/AI-CatBoost-orange)
![Outlook](https://img.shields.io/badge/Integration-Win32COM-blue)
![GUI](https://img.shields.io/badge/UI-CustomTkinter-green)

## 📖 Descripción

**Mail Intelligence Engine** es una suite de productividad que transforma tu Outlook. Utilizando modelos de Machine Learning (CatBoost + NLP), analiza tus correos históricos para entender qué es importante *para ti*.

El sistema clasifica automáticamente los correos entrantes aplicando **Categorías de Color** en Outlook (🔴 Urgente / 🟡 Revisar), permitiéndote enfocar tu atención donde realmente importa.

### 🚀 Novedades V3
*   **Interfaz Gráfica Unificada:** Todo el poder del sistema en una sola ventana moderna (`app_master.py`).
*   **Modo Ejecutable:** No requiere instalación de Python.
*   **Dashboards:** Visualización de métricas y estadísticas de tu correo.

---

## 📦 Instalación y Uso

Tienes dos formas de usar el sistema:

### Opción A: Ejecutable (Portable)
*Recomendado para usuarios finales.*

1.  Ve a la carpeta `dist\MailIntelligence_Folder`.
2.  Ejecuta `MailIntelligence_Folder.exe`.
3.  ¡Listo! No necesitas instalar nada más.

> **Nota:** Existe una versión de archivo único (`MailIntelligence.exe`), pero la versión en carpeta (`_Folder`) es mucho más rápida al iniciar y evita falsos positivos de antivirus.

### Opción B: Código Fuente (Desarrolladores)

1.  **Clonar repositorio:**
    ```bash
    git clone https://github.com/WalterWr7/mail-intelligence-engine.git
    cd mail-intelligence-engine
    ```

2.  **Instalar dependencias:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Ejecutar:**
    ```bash
    python app_master.py
    ```

---

## 🛠️ Flujo de Trabajo

La aplicación te guía paso a paso:

1.  **Minería de Datos (Data Mining):** 
    *   Extrae tu historial de Outlook (últimos 365 días por defecto).
    *   Genera un dataset local (`dataset_masivo.csv`).

2.  **Entrenamiento (Training):**
    *   Entrena un modelo predictivo personalizado con tus datos.
    *   Genera el "cerebro" (`cerebro_priorizacion.joblib`).

3.  **Vigilancia (Monitoring):**
    *   Activa el agente en tiempo real.
    *   Clasifica correos nuevos según llegan a tu bandeja.

---

## 🏗️ Arquitectura Técnica

El proyecto sigue una arquitectura modular dirigida por la UI:

```text
mail_intelligence/
│
├── 📜 app_master.py           # [MAIN] Interfaz Gráfica (GUI) y Orquestador
│
├── 🧠 Backend (Módulos)
│   ├── 📜 01_data_extractor.py    # ETL: Extracción MAPI y limpieza
│   ├── 📜 02_model_trainer.py     # ML: Entrenamiento CatBoost
│   └── 📜 03_inference_engine.py  # Runtime: Vigilancia en tiempo real
│
├── 📁 dist/                   # Ejecutables generados (Compilados)
│   └── 📁 MailIntelligence_Folder # Versión optimizada (OneDir)
│
└── 📄 requirements.txt        # Dependencias (pandas, catboost, ctk, pywin32)
```

## 🔒 Privacidad y Seguridad

*   **Procesamiento Local:** Ningún correo sale de tu computadora. Todo el análisis ocurre en tu CPU.
*   **No Destructivo:** El sistema **nunca elimina ni mueve** correos. Solo añade etiquetas de color.
*   **Código Abierto:** Puedes auditar cada línea de código.

---

**Desarrollado por Walter Llana**
*v3.0.0 - Edición Enterprise*
