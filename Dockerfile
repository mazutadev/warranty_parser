FROM python:3.12-slim

WORKDIR /app

# Копируем только нужные файлы
COPY src/ ./src/
COPY requirements.txt ./

RUN pip install --no-cache-dir -r requirements.txt

# Точка входа
ENTRYPOINT ["python", "src/main.py"] 