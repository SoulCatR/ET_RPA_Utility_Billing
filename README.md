# ⚡ ET RPA Utility Billing (Procesador de Consumos)

Automatización Robótica de Procesos (RPA) construida con Python y Streamlit para la extracción inteligente de datos desde recibos de servicios públicos (Agua y Energía) en formato PDF y su inyección automática en un reporte maestro de Excel.

Este proyecto elimina la necesidad de digitación manual, reduciendo errores humanos y ahorrando horas de trabajo administrativo.

## 🚀 Características Principales

- **Procesamiento Masivo:** Soporta la carga de múltiples archivos PDF simultáneamente, así como la lectura directa desde archivos comprimidos `.zip`.
- **Extracción Inteligente:** Utiliza algoritmos de búsqueda dinámica en tablas para localizar los códigos de consumo (ej. 840 para Agua, 841 para Energía) sin depender de posiciones fijas.
- **Mapeo Automatizado (Fuzzy Matching):** Analiza las descripciones de los PDFs y las cruza con un diccionario interno para asignar correctamente los importes a la tienda o local correspondiente en el Excel.
- **Interfaz Web:** Interfaz gráfica intuitiva, moderna y accesible desde el navegador web, desarrollada con Streamlit.
- **Procesamiento en Memoria:** Por seguridad y limpieza, el sistema procesa los archivos temporalmente en la memoria RAM y entrega el Excel actualizado sin almacenar datos sensibles en el servidor.

## 🛠️ Tecnologías Utilizadas

- **Python 3.x**
- **Streamlit:** Framework para la creación de la interfaz de usuario web.
- **pdfplumber:** Librería de alta precisión para la extracción de texto y tablas de documentos PDF.
- **openpyxl:** Motor para leer y modificar archivos `.xlsx` manteniendo el formato, fórmulas y celdas combinadas originales.
- **pandas:** Utilizado para la estructuración y visualización del reporte (log) de resultados en pantalla.

## 📂 Estructura del Repositorio

\`\`\`text
ET_RPA_Utility_Billing/
├── app.py # Lógica principal del procesador y UI de Streamlit
├── requirements.txt # Dependencias necesarias para el entorno
├── .gitignore # Archivos y directorios excluidos del control de versiones
└── README.md # Documentación del proyecto
\`\`\`

_Nota: Las carpetas locales con archivos de prueba (`dataset/`, `excel/`) han sido excluidas mediante `.gitignore` para proteger la privacidad de la data y mantener el repositorio ligero._

## 💻 Instalación y Uso Local

Si deseas ejecutar este proyecto en tu propia computadora, sigue estos pasos:

1. **Clona este repositorio:**
   \`\`\`bash
   git clone https://github.com/SoulCatR/ET_RPA_Utility_Billing.git
   cd ET_RPA_Utility_Billing
   \`\`\`

2. **Instala las dependencias:**
   Se recomienda usar un entorno virtual.
   \`\`\`bash
   pip install -r requirements.txt
   \`\`\`

3. **Ejecuta la aplicación web:**
   \`\`\`bash
   streamlit run app.py
   \`\`\`
   _Esto abrirá automáticamente una pestaña en tu navegador web en la dirección `http://localhost:8501`._

## ☁️ Despliegue en la Nube

Este proyecto está preparado para ser desplegado en plataformas PaaS como **Railway**. Simplemente conecta este repositorio de GitHub a tu cuenta de Railway y la plataforma leerá el archivo `requirements.txt` para levantar la aplicación web de forma automática.
