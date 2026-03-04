# 📄 Generador de Documentos

Rellena plantillas **Excel (.xlsx)** y **Word (.docx)** desde una base de datos Excel, con soporte para inserción automática de firmas.

## 🚀 Cómo usar (versión web)

1. Sube tu **Excel de datos** con los registros
2. Sube tus **plantillas** (.xlsx o .docx) con marcadores `{{campo}}`
3. Sube las **imágenes de firma** (opcional, el nombre del archivo debe coincidir con el valor del campo `firma` en el Excel)
4. Pulsa **GENERAR** y descarga el ZIP con todos los documentos

## 📝 Marcadores en plantillas

En tus plantillas usa `{{nombre_columna}}` para insertar valores del Excel.  
Ejemplo: `{{nombre}}`, `{{fecha}}`, `{{empresa}}`

Para insertar una firma: coloca un shape con texto `{{firma}}` en el lugar deseado.

## 🛠 Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Deploy en Streamlit Cloud

1. Sube este repositorio a GitHub
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu repo y selecciona `app.py`
4. ¡Listo!
