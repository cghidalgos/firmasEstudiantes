# Dockerfile para la app Flask de firmas
FROM python:3.13-slim

WORKDIR /app

COPY . /app

# Instalar dependencias
RUN pip install --no-cache-dir flask python-docx pillow

EXPOSE 5006

CMD ["python", "app.py"]
