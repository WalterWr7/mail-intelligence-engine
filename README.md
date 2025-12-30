\# 📧 Mail Intelligence Engine



> \*\*Sistema de Priorización Predictiva para Microsoft Outlook basado en Machine Learning (CatBoost + NLP).\*\*



!\[Python](https://img.shields.io/badge/Python-3.12%2B-blue?logo=python)

!\[CatBoost](https://img.shields.io/badge/Model-CatBoost-orange)

!\[Outlook](https://img.shields.io/badge/Integration-Win32COM-blue)



\## 📖 Descripción



\*\*Mail Intelligence Engine\*\* es un asistente virtual diseñado para optimizar el flujo de trabajo en entornos corporativos de alta demanda. A diferencia de las reglas estáticas de Outlook, este sistema utiliza \*\*Aprendizaje Supervisado\*\* para entender tu comportamiento histórico.



El modelo analiza no solo quién envía el correo, sino el contexto semántico del asunto y tu rol en la conversación (To/CC), para predecir la probabilidad de que un correo requiera una acción inmediata.



\### 🚀 Características Principales



\* \*\*Enfoque de Alta Seguridad (High Recall):\*\* El modelo prioriza la sensibilidad (70% Recall) para asegurar que ningún correo crítico sea ignorado.

\* \*\*Aprendizaje Híbrido:\*\* Combina procesamiento de lenguaje natural (TF-IDF en Asuntos) con metadatos estructurados (Dominios, Destinatarios).

\* \*\*Integración No Destructiva:\*\* No mueve correos. Utiliza el sistema de \*\*Categorías de Color\*\* de Outlook (🔴 Urgente / 🟡 Revisar) para una clasificación visual fluida.

\* \*\*Privacidad Total:\*\* Todo el procesamiento ocurre localmente en tu máquina. Ningún dato sale de tu ordenador.



---



\## 🏗️ Arquitectura del Proyecto



El sistema opera en tres fases secuenciales:



1\.  \*\*Minería de Datos (ETL):\*\* Extracción forense del historial de correos (últimos 365 días) vía interfaz MAPI.

2\.  \*\*Entrenamiento (Training):\*\* Generación del modelo predictivo usando \*\*CatBoost\*\* con balanceo de pesos automático.

3\.  \*\*Inferencia (Live):\*\* Un agente "centinela" que monitorea la bandeja de entrada en tiempo real.



```text

mail\_intelligence/

│

├── 📜 01\_data\_extractor.py       # Extrae historial a CSV

├── 📜 02\_model\_trainer.py        # Entrena el modelo y evalúa métricas

├── 📜 03\_inference\_engine.py     # Agente de vigilancia en tiempo real

│

├── 🧠 cerebro\_priorizacion.joblib # Modelo entrenado (Ignorado en git)

├── 📊 dataset\_masivo\_1ano.csv     # Datos históricos (Ignorado en git)

└── 📄 requirements.txt            # Dependencias

