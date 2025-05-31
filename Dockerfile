# Use uma imagem Python pública e estável
FROM python:3.11-slim-buster

# Define o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copia o arquivo requirements.txt e instala as dependências
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o código da sua aplicação Flask (app.py)
COPY app.py .

# Copia a pasta 'templates' (que contém o Excel e o HTML)
COPY templates/ templates/

# Informa qual porta a aplicação irá escutar (mais para documentação)
EXPOSE 8000

# Comando para iniciar a aplicação Flask com Gunicorn
# O Render injeta a variável de ambiente $PORT, então a aplicação deve escutar nela.
CMD ["gunicorn", "-b", "0.0.0.0:${PORT}", "app:app"]
