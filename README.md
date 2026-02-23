<div align="center">
  <img src="logojup.png" alt="Logo" width="150" height="auto">
  <h1>BotEliva - Automação InControl</h1>
  <p>Automatização de extração, relatórios e envio de e-mails para o sistema Intelbras InControl</p>
  
  [![Python](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)
  [![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-green.svg)](https://github.com/TomSchimansky/CustomTkinter)
  [![Selenium](https://img.shields.io/badge/Web-Selenium-orange.svg)](https://www.selenium.dev/)
  [![Pandas](https://img.shields.io/badge/Data-Pandas-red.svg)](https://pandas.pydata.org/)
</div>

<hr>

## 🎯 Finalidade
BotEliva é uma aplicação desktop desenvolvida para automatizar o processo maçante de extração de relatórios de acesso no sistema de catracas e fechaduras **Intelbras InControl**. 

Seu principal objetivo é:
1. Conectar-se diariamente à interface web do InControl.
2. Extrair programaticamente os logs de acessos do dia atual.
3. Processar esses dados localmente, removendo duplicidades e validando as presenças.
4. Gerar estatísticas sobre a ocupação em tempo real (Total de Movimentos, Pessoas na Obra, e por Departamento).
5. Compilar um e-mail HTML bonito e bem desenhado, acompanhado de gráficos e informativos, enviando-o automaticamente para listas de interessados.

## ⚙️ Desenvolvimento
O projeto foi inteiramente construído utilizando **Python** focado em legibilidade e manutenabilidade.
- **Frontend / GUI**: Toda a interface de controle, logs visuais, agendador e gerenciamento de configuração foi desenvolvida com a moderna biblioteca de interface `CustomTkinter`.  
- **Backend Automation**: A manipulação da página web do InControl (autenticação, filtros, botões e download do `.csv`) é feita usando `Selenium WebDriver`. O Chrome corre em background *(headless)* sem interferir no uso normal do computador.
- **Data Engineering**: Os relatórios `.csv` gerados pelo InControl podem ter ruídos, formatos diferentes (UTF-8, Latin-1) e duplicações. O `Pandas` foi extensamente usado para filtrar acessos inválidos, cruzar status e criar analíticos confiáveis.
- **Gráficos e Imagens**: O projeto gera cartões de imagem interativos que podem ser compartilhados via WhatsApp, além de gráficos de picos usando `Matplotlib` e a lib `Pillow (PIL)`.

---

## 🚀 Setup e Uso

### Pré-requisitos
- Ter o **Python 3.x** e o **Google Chrome** instalados.

### 1. Clonando e Instalando Dependências
Abra o seu terminal:

```bash
git clone https://github.com/SEU_USUARIO/boteliva.git
cd boteliva
pip install -r requirements.txt
```

### 2. Configurando o Projeto
Para iniciar pela primeira vez:
Renomeie o arquivo `.env.example` para `.env` e configure seus IPs/senhas:
```dotenv
INCONTROL_URL=https://192.168.0.X:4445/#/home/eventos-usuario
INCONTROL_USER=admin
INCONTROL_PASSWORD=senha
# E para Email:
SMTP_USER=bot@gmail.com
SMTP_PASSWORD=senha_de_app
```
Mas não se preocupe, também é possível fazer tudo pela Interface Gráfica de forma muito mais fácil!

### 3. Rodando o Aplicativo
Você pode rodar diretamente o arquivo da GUI em Python:
```bash
python gui_app.py
```

### 4. Build de Executável (Opcional)
Você também pode gerar um arquivo `.exe` para distribuição no Windows usando o `PyInstaller`. Incluímos um script PowerShell para facilitar ou via comando direto:
```bash
./build.ps1
# Ou via pyinstaller:
pyinstaller BotEliva.spec
```

---

## 🔒 Segurança
Lembre-se de **nunca** committar os arquivos `config.json` ou o `.env` caso esse repositório fique público! Eles já estão sendo ignorados pelo arquivo `.gitignore`. Em e-mails, lembre-se de usar a "App Password" em vez da sua senha direta, ativando na aba de segurança da conta da Google.

<br>
<div align="center">
  <i>Desenvolvido e mantido para automação e melhoria de processos corporativos.</i>
</div>
