# Dockerfile para la app Flask de firmas
FROM python:3.13-slim

WORKDIR /app

COPY . /app

# Dependencias del sistema (para convertir .doc -> .docx)
RUN apt-get update \
	&& apt-get install -y --no-install-recommends libreoffice-writer \
	&& rm -rf /var/lib/apt/lists/*

# Instalar dependencias Python
RUN pip install --no-cache-dir flask python-docx pillow

EXPOSE 5006

CMD ["python", "app.py"]
