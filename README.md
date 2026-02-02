# firmasEstudiantes

App en Flask para registrar firmas de estudiantes por curso y generar un acta en Word (DOCX).

## Ejecutar local

1. Instalar dependencias:
	- `pip install flask python-docx pillow`
2. Ejecutar:
	- `python app.py`
3. Abrir:
	- `http://localhost:5006`

## Ejecutar con Docker

- `docker compose up --build`
- Abrir `http://localhost:5006`

## Archivos generados

- `actas/`: actas por curso en formato `F6_Acta_<curso>.docx`
- `firmas_temp/`: imágenes PNG de las firmas
- `cursos.txt`: lista de cursos

## Flujo

1. En "Crear curso" el profesor registra:
	- Nombre del profesor
	- Nombre del curso
	- Sube el acta en `.docx` o `.doc`
2. El acta se guarda en `actas/` con nombre `f6_<nombre_del_curso>.docx` (sanitizado).
3. En "Registro" los estudiantes eligen el curso y firman; la app inserta la firma en ese `.docx`.

## Nota sobre archivos .doc

- Para poder insertar firmas, internamente el acta debe ser `.docx`.
- Si subes `.doc`, la app lo convertirá a `.docx` usando LibreOffice (`soffice`).
- En Docker ya viene incluido; si ejecutas local sin Docker, instala LibreOffice o sube directamente `.docx`.

## Variables de entorno (opcional)

- `SECRET_KEY`: requerido para mensajes flash (si no se define, se usa un valor inseguro solo para desarrollo)
- `FLASK_DEBUG`: `1` o `0`
- `PORT`: puerto (por defecto `5006`)