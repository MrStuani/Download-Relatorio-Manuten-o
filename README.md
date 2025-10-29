# Download-Relatorio-Manuten-o
Este projeto é uma automação em Python desenvolvida para simplificar o processo de geração, conversão e envio de relatórios de manutenção.
A aplicação acessa uma plataforma web automaticamente via Selenium, realiza login autenticado (com suporte a TOTP/2FA), gera relatórios dentro de um intervalo de datas e converte o arquivo baixado para CSV.
Em seguida, os dados são enviados para uma planilha no Google Sheets através da API do Google.

🧩 Funcionalidades
-Login automatizado com autenticação em duas etapas (TOTP).
-Geração automática de relatórios com intervalo de datas dinâmico.
-Detecção e conversão automática de arquivos .xls, .xlsx e .html para .csv.
-Upload dos dados processados para uma aba específica no Google Sheets.
-Configuração simples via variáveis de ambiente.

⚙️Tecnologias utilizadas
-Python 3
-Selenium WebDriver
-Pandas
-PyOTP
-gspread + Google Sheets API
-webdriver_manager

🧠 Objetivo do projeto

Este projeto foi desenvolvido como uma automação pessoal, com foco em otimizar tarefas repetitivas, demonstrar boas práticas de automação web e integração com APIs externas.
Ele também serve como exemplo prático de uso do Selenium e Google Sheets para controle de relatórios empresariais.

🛡️ Observação

As credenciais e URLs utilizadas são genéricas ou mascaradas por motivos de segurança.
Para usar este projeto, substitua os placeholders no código pelo seu ambiente real e mantenha dados sensíveis no arquivo .env.

📄 Licença

Este projeto é de uso livre para fins educacionais e demonstração de automação Python.

