# Gerador de Alertas de Férias – BWA Global

## 🚀 Deploy no Render

### 1. Suba esta pasta no GitHub
Crie um repositório e envie todos os arquivos:
- app.py
- template_alerta.docx
- requirements.txt
- render.yaml
- README.md

### 2. Configure no Render
1. Acesse render.com → "New → Web Service"
2. Conecte o repositório GitHub
3. O render.yaml já configura tudo automaticamente
4. Clique em "Deploy"

### Execução local
```bash
pip install -r requirements.txt
sudo apt install libreoffice  # Linux
python app.py
```
Acesse: http://localhost:5000
