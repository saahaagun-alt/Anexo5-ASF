# Anexo 5 ASF - App Web

Aplicación web para generar el Anexo 5 ASF mensual.

## Cómo usar (usuario final)

1. Abre el link en cualquier dispositivo (celular, tablet, PC)
2. Sube la plantilla Anexo 5
3. Sube el auxiliar del mes
4. Sube los CFDIs (opcional)
5. Selecciona el mes
6. Da clic en "Generar Anexo 5"
7. Se descarga el archivo Excel ya lleno

## Probar localmente

```bash
pip install -r requirements.txt
python app.py
```

Abre http://localhost:5000 en el navegador.

## Publicar gratis en Render

1. Sube esta carpeta a un repo de GitHub
2. Ve a https://render.com → New Web Service → Conecta el repo
3. Render detecta `render.yaml` automáticamente
4. Da clic en "Create Web Service"
5. En 2-3 minutos te da una URL pública del tipo `https://anexo5-asf.onrender.com`

La URL se comparte por WhatsApp y cualquiera la usa desde su celular.

## Notas

- Plan free de Render: la app duerme tras 15 min de inactividad y tarda ~30 seg en despertar la primera vez. Después funciona normal.
- Sin límite de uso ni cuenta requerida por parte de los usuarios.
- Los archivos subidos NO se guardan en el servidor — se procesan y se borran.
