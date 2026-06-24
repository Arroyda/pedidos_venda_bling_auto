# PedidosAuto – Bling Web

Aplicação web que automatiza a criação de **pedidos de venda no Bling** a partir de uma planilha Excel. O usuário envia a planilha pela interface, acompanha o progresso em tempo real e baixa, ao final, uma planilha de resultado com os números dos pedidos criados e eventuais ajustes de estoque.

A automação é feita via **Selenium** (navegador Chrome), simulando o preenchimento do formulário de pedidos do Bling.

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Como executar](#como-executar)
- [Como usar](#como-usar)
- [Formato da planilha](#formato-da-planilha)
- [Planilha de resultado](#planilha-de-resultado)
- [Solução de problemas](#solução-de-problemas)
- [Segurança](#segurança)

---

## Funcionalidades

- Upload de planilha Excel (`.xlsx`) pela interface web.
- Login automático no Bling com as credenciais informadas no formulário.
- Criação de múltiplos pedidos em sequência, agrupando os produtos de cada pedido.
- **Numeração automática à prova de conflito**: o número do pedido é atribuído pelo próprio Bling, evitando o erro de "número já cadastrado".
- Validação de estoque por item, com ajuste da quantidade enviada quando o estoque é insuficiente.
- Acompanhamento dos logs em tempo real (durante a execução).
- Botão para cancelar a execução a qualquer momento.
- Geração de uma planilha de resultado para download, com os números dos pedidos e marcações de estoque.

---

## Estrutura do projeto

```
PedidosAuto - Bling Web/
├── app.py                 # Servidor Flask (upload, logs em tempo real, download)
├── bot.py                 # Lógica do bot (Selenium) encapsulada na classe BlingBot
├── requirements.txt       # Dependências Python
├── chromedriver.exe       # Driver do Chrome (Windows)
├── templates/
│   └── index.html         # Página principal
├── static/
│   └── style.css          # Estilos da interface
├── uploads/               # Planilhas enviadas (geradas em tempo de execução)
├── prints/                # Capturas de tela de erros (geradas em tempo de execução)
└── results/               # Pasta de resultados (não utilizada — resultado vai por download)
```

> As credenciais **não ficam mais em arquivo**: são preenchidas no formulário a cada execução e usadas apenas em memória.

---

## Requisitos

- **Python 3.10 ou superior** (o código usa sintaxe de tipos como `str | None`).
- **Google Chrome** instalado.
- **ChromeDriver** compatível com a versão do seu Chrome.
- Uma conta no **Bling** com permissão para criar pedidos de venda.

---

## Instalação

1. Tenha o Python instalado. Para conferir:

   ```bash
   python --version
   ```

2. (Recomendado) Crie e ative um ambiente virtual dentro da pasta do projeto:

   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # Linux / macOS
   source venv/bin/activate
   ```

3. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

   As dependências são: `flask`, `selenium`, `pandas`, `openpyxl`, `python-dotenv`.

### ChromeDriver

O projeto já inclui um `chromedriver.exe` (Windows). Ele precisa ser **da mesma versão do seu Google Chrome**.

- Verifique a versão do Chrome em `chrome://settings/help`.
- Caso o bot não inicie o navegador, baixe o ChromeDriver correspondente em
  <https://googlechromelabs.github.io/chrome-for-testing/> e substitua o arquivo.

---

## Como executar

Na pasta do projeto, rode:

```bash
python app.py
```

O servidor sobe em:

```
http://127.0.0.1:5000
```

Abra esse endereço no navegador.

> **Atenção à rede:** por padrão o `app.py` sobe com `host="0.0.0.0"`, o que torna a aplicação acessível por outros computadores da mesma rede (sem criptografia). Para uso apenas local, troque para `host="127.0.0.1"`. Veja [Segurança](#segurança).

---

## Como usar

1. Abra `http://127.0.0.1:5000` no navegador.
2. Preencha **todos os campos** (são obrigatórios):
   - **E-mail do Bling**
   - **Senha do Bling**
   - **Cliente** (ex.: `ENVIO FULL ML`)
   - **Loja** (ex.: `Loja Full`)
   - **Frete** (ex.: `9 - Sem Ocorrência de Transporte`)
3. Selecione a planilha `.xlsx`.
4. Clique para iniciar. Uma janela do Chrome será aberta e o bot fará login e criará os pedidos.
5. Acompanhe os logs em tempo real na própria página. Se precisar, use o botão de **cancelar**.
6. Ao terminar, **baixe a planilha de resultado**.

> Caso algum campo obrigatório fique em branco, a aplicação avisa exatamente quais campos faltam, sem iniciar o bot.

---

## Formato da planilha

A planilha precisa ter uma aba chamada exatamente:

```
Montagem de Pedidos - Envio
```

Nas primeiras linhas dessa aba, o bot procura automaticamente a linha de **cabeçalho** (identificada pela coluna `SKU PAI`), então não há problema se houver título ou banner no topo.

Colunas **obrigatórias** no cabeçalho:

| Coluna          | Significado                         |
|-----------------|-------------------------------------|
| `SKU PAI`       | Código (SKU) do produto             |
| `QNT PLANEJADA` | Quantidade planejada do produto     |

### Como os pedidos são agrupados

Cada **bloco de linhas preenchidas** vira um pedido. Uma **linha em branco** (SKU ou quantidade vazios) separa um pedido do próximo.

Exemplo:

| SKU PAI   | QNT PLANEJADA |
|-----------|---------------|
| SKU-001   | 10            |
| SKU-002   | 5             |
|           |               |  ← linha em branco separa os pedidos
| SKU-003   | 8             |

No exemplo acima são criados **dois pedidos**: o primeiro com `SKU-001` e `SKU-002`, o segundo com `SKU-003`.

---

## Planilha de resultado

Ao final, é gerada (em memória, para download) uma cópia da planilha com:

- O **número do pedido** atribuído pelo Bling, preenchido na coluna correspondente (cabeçalho contendo `PEDIDO` + `VENDA`).
- A **quantidade realmente enviada** na coluna de quantidade do pedido (cabeçalho contendo `QNT` + `PEDIDO` + `VENDA`), com destaque por cor quando houve ajuste:
  - 🟡 **Amarelo**: estoque insuficiente — a quantidade enviada foi reduzida para o disponível.
  - 🔴 **Vermelho**: produto sem estoque.

O nome do arquivo segue o padrão `NomeDaPlanilha_resultado_AAAAMMDD_HHMMSS.xlsx`.

---

## Solução de problemas

**O navegador não abre / erro ao iniciar o Chrome**
Versão do ChromeDriver incompatível com o Chrome instalado. Baixe a versão correta (ver [ChromeDriver](#chromedriver)).

**"Não encontrei a coluna 'SKU PAI'..."**
Confira se a aba se chama `Montagem de Pedidos - Envio` e se existe a coluna `SKU PAI` no cabeçalho.

**"Preencha todos os campos obrigatórios..."**
Algum campo do formulário ficou em branco. A mensagem indica quais.

**Login falha**
Verifique e-mail e senha. Lembre-se de que o login é o mesmo do site do Bling.

**Um SKU não foi encontrado**
O log informa o SKU e mostra os candidatos retornados pelo Bling. Confira se o SKU da planilha bate com o cadastrado no Bling.

**Capturas de erro**
Quando ocorre um erro ao inserir um produto, uma captura de tela é salva na pasta `prints/` para ajudar no diagnóstico.

---

## Segurança

- **Credenciais**: o e-mail e a senha são informados no formulário a cada uso e mantidos apenas em memória durante a execução — não são gravados em disco.
- **Rede**: o `app.py` está configurado com `host="0.0.0.0"`, expondo a aplicação a toda a rede local via HTTP (sem criptografia). Para uso pessoal, prefira `host="127.0.0.1"`. Se for disponibilizar para outras pessoas, coloque atrás de HTTPS.
- **Modo debug**: o servidor roda com `debug=True`, adequado para desenvolvimento, mas **não recomendado** em ambiente de produção/compartilhado.

---

## Observação sobre a abordagem técnica

A automação atual é feita por **Selenium** (controle do navegador). O Bling também oferece uma **API oficial (REST v3)**, que permite criar pedidos diretamente, sem navegador — mais rápida e mais estável, e que elimina de vez a questão do número do pedido (basta deixar o Bling numerar). Migrar para a API é uma evolução possível no futuro, caso se queira mais desempenho e robustez.
