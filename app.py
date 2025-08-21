      # Ponto de entrada da aplicação



from flask import Flask
from src.routes.upload import upload_bp

# Inicializa a aplicação Flask
app = Flask(__name__)

# Registra o blueprint de upload
app.register_blueprint(upload_bp)

# Rota básica para teste
@app.route("/")
def home():
    return "API está funcionando!"

# Executa a aplicação
if __name__ == "__main__":
    app.run(debug=True)