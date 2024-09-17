# Usar uma imagem base oficial do Python (versão 3.9, por exemplo)
FROM python:3.9-slim

# Definir o diretório de trabalho dentro do container
WORKDIR /app

# Copiar o arquivo de requisitos para o diretório de trabalho
COPY requirements.txt ./

# Instalar as dependências listadas no requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o restante do código da aplicação para o container
COPY . .

# Expor a porta que será usada pela aplicação (ajuste a porta conforme necessário)
EXPOSE 5000

# Definir o comando para rodar a aplicação (ajuste conforme a estrutura do seu projeto)
CMD ["python", "Main.py"]
