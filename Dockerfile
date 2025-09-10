# 1. Usar una imagen base oficial de Python
FROM python:3.10-slim

# 2. Establecer el directorio de trabajo dentro del contenedor
WORKDIR /app

# 3. Copiar el archivo de dependencias
COPY requirements.txt .

# 4. Instalar las dependencias
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copiar el resto del código de la aplicación
COPY . .

# 6. Exponer el puerto en el que correrá Gunicorn
EXPOSE 8000

# 7. Definir el comando para iniciar la aplicación
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "app:app"]