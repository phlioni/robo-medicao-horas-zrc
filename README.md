# RPA - Automação de Medição de Horas (Projeto ZRC)

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)

Este robô é uma versão especializada para automatizar o processo de medição de horas exclusivamente para o **projeto ZRC**. Sua principal característica é a lógica de extração de dados com um período de faturamento customizado, que vai do dia 15 do mês anterior até o dia 14 do mês atual.

## 🚀 Funcionalidades Principais

-   **Automação Web**: Acessa a plataforma PSOffice, realiza login e navega até a tela de relatórios.
-   **Período de Datas Customizado**: Calcula e aplica dinamicamente o filtro de data customizado: do **dia 15 do mês anterior até o dia 14 do mês atual**.
-   **Extração de Dados**: Baixa o relatório de horas consolidado para o período especificado.
-   **Desbloqueio e Processamento**: Habilita a edição do arquivo Excel baixado, filtra os dados para incluir apenas projetos que começam com "ZRC" e processa apenas os apontamentos com status "Aprovado" e comentários válidos.
-   **Geração de Relatórios Detalhados**:
    -   Cria uma estrutura de pastas para o mês da medição (ex: `relatorios/medicao/Setembro-2025/`).
    -   Gera **um arquivo Excel para cada sub-projeto ZRC encontrado** (ex: `ZRC-01-relatorio-horas-Setembro-2025.xlsx`).
    -   Cada arquivo Excel contém abas separadas para cada profissional que trabalhou no projeto, com uma linha de "Total de Horas" no topo de cada aba para fácil conferência.
-   **Envio de E-mail Dedicado**: Envia um único e-mail para os stakeholders do projeto ZRC, contendo uma tabela resumo no corpo e todos os relatórios Excel detalhados (gerados na etapa anterior) em anexo.
-   **Monitoramento**: Ao final de cada execução (com sucesso ou falha), envia um e-mail de status para o administrador com o tempo de cada etapa e detalhes de eventuais erros.
-   **Agendamento**: Projetado para ser executado de forma autônoma no dia 15 de cada mês, alinhado com o fechamento do período de medição.

## 🛠️ Pré-requisitos

-   [Python](https://www.python.org/downloads/) (versão 3.8 ou superior)
-   [Git](https://git-scm.com/downloads/) (opcional)
-   Google Chrome
-   Microsoft Excel instalado

## ⚙️ Instalação

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/phlioni/robo-medicao-horas-zrc.git](https://github.com/phlioni/robo-medicao-horas-zrc.git)
    cd robo-medicao-horas-zrc
    ```

2.  **Crie e ative o ambiente virtual:**
    ```bash
    # Windows
    python -m venv venv
    venv\Scripts\activate

    # Linux / macOS
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

## 📝 Configuração

1.  **Crie o arquivo `config.py`**: Se não existir, faça uma cópia do arquivo `config.py.template` e renomeie-a para `config.py`.
2.  [cite_start]**Preencha `config.py`**: Abra o arquivo `config.py` e preencha todas as variáveis necessárias[cite: 1]:
    -   `SITE_LOGIN` e `SITE_SENHA`.
    -   `EMAIL_REMETENTE` e `EMAIL_SENHA`.
    -   [cite_start]`ZRC_DESTINATARIO_PRINCIPAL` e `ZRC_DESTINATARIOS_COPIA`[cite: 2].
    -   [cite_start]`STATUS_EMAIL_DESTINATARIO`[cite: 3].
    -   Dados da sua assinatura de e-mail.
3.  [cite_start]**Pasta de Imagens**: Certifique-se de que a pasta `images` na raiz do projeto contém os arquivos `mosten.png` e `selos.png` para a assinatura dos e-mails[cite: 3].

## ▶️ Execução

Para executar o robô, utilize os scripts de inicialização que ativam o ambiente virtual automaticamente.

#### Windows

-   **Execução Manual (Visível):** Dê um duplo clique no arquivo `run_robot.bat`. Uma janela de terminal e uma do Chrome se abrirão, mostrando o progresso.
-   **Execução Silenciosa (Invisível):** Dê um duplo clique no arquivo `lancar_robo_oculto.vbs`. O robô será executado em segundo plano, sem nenhuma janela visível.

#### Linux / macOS

1.  **Dê permissão de execução ao script:**
    ```bash
    chmod +x run_robot.sh
    ```
2.  **Execute o script:**
    ```bash
    ./run_robot.sh
    ```

## 🗓️ Agendamento

O robô foi projetado para ser executado no dia 15 de cada mês.

#### Windows (Usando o Agendador de Tarefas)

1.  Abra o Agendador de Tarefas.
2.  Crie uma nova tarefa.
3.  **Ação**: "Iniciar um programa". Aponte para o caminho absoluto do arquivo `lancar_robo_oculto.vbs` (para execução silenciosa).
4.  **Disparador (Gatilho)**: Configure o disparador para **"Mensalmente"**, nos **"Dias:" `15`**, para todos os meses.

#### Linux (Usando o Cron)

1.  Abra o editor do crontab: `crontab -e`
2.  Adicione uma linha para agendar a execução no dia 15 de cada mês (ex: às 09:00).
    ```crontab
    # Executar às 9:00 da manhã do dia 15 de cada mês
    0 9 15 * * /home/seu_usuario/caminho/para/robo-medicao-horas-zrc/run_robot.sh
    ```

## 📂 Estrutura do Projeto
.
├── images/
│   ├── mosten.png
│   └── selos.png
├── relatorios/
├── venv/
├── .gitignore
├── automacao_web.py
├── config.py
├── config.py.template
├── envio_email.py
├── excel_handler.py
├── lancar_robo_oculto.vbs
├── main.py
├── processamento_dados.py
├── README.md
├── requirements.txt
├── run_robot.bat
└── run_robot.sh