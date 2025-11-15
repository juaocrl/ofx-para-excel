# 📘 Conversor OFX para Excel

### *Automatização da conversão de arquivos bancários (.OFX) para planilhas Excel estruturadas*

<p align="center">

<strong>Interface gráfica limpa, conversão rápida e layout
compatível com sistemas contábeis </strong>

</p>


------------------------------------------------------------------------

## 📑 **Tabela de Conteúdos**

-   [Visão Geral](#-visao-geral)
-   [Funcionalidades](#-funcionalidades)
-   [Tecnologias Usadas](#-tecnologias-usadas)
-   [Instalação](#-instalacao)
-   [Como Usar](#-como-usar)
-   [Estrutura do Projeto](#-estrutura-do-projeto)
-   [Processamento dos Arquivos OFX](#-processamento-dos-arquivos-ofx)
-   [Possíveis Melhorias](#-possiveis-melhorias)
-   [Licença](#-licenca)

------------------------------------------------------------------------

## 🔍 Visão Geral

O conversor OFX para Excel tem como objetivo simplificar a preparação
dos extratos bancários para integração com softwares contábeis.\
A planilha gerada segue um layout padronizado, permitindo que o contador
ou analista classifique os lançamentos, atribua códigos do plano de
contas e revise informações antes de realizar a importação no sistema
contábil.\
Este processo otimiza a classificação, lançamento e conciliação
bancária, reduzindo retrabalho e melhorando a organização financeira.

Apesar de ter sido desenvolvido inicialmente para atender às
necessidades específicas de uma empresa, o layout é totalmente
flexível.\
Nada impede que o projeto seja aprimorado ou adaptado para o escritório
contábil que você desejar, permitindo ajustar colunas, incluir novas
regras, alterar nomenclaturas ou estruturar a planilha de acordo com o
sistema contábil utilizado.\
Essa flexibilidade torna o conversor uma ferramenta útil não apenas para
uma empresa específica, mas para qualquer profissional ou escritório que
precise tratar extratos bancários de forma organizada, padronizada e
eficiente.

------------------------------------------------------------------------

## ✨ Funcionalidades

-   Interface gráfica intuitiva (Tkinter)
-   Processamento de múltiplos arquivos OFX
-   Padronização dos dados contábeis
-   Conversão confiável usando pandas, ofxparse e openpyxl
-   Tratamento de erros e mensagens claras

------------------------------------------------------------------------

## 🧰 Tecnologias Usadas

| Tecnologia     | Uso                                      |
|----------------|-------------------------------------------|
| **Python 3**   | Base do projeto                           |
| **Tkinter**    | Interface gráfica (GUI)                   |
| **ofxparse**   | Leitura e interpretação dos arquivos OFX  |
| **pandas**     | Manipulação de dados e DataFrame          |
| **openpyxl**   | Escrita e geração do arquivo Excel (.xlsx)|
| **Pillow (PIL)** | Ícones e imagens da interface            |


------------------------------------------------------------------------

## ⚙️ Instalação

``` bash
# 1. Baixe o repositório
git clone https://github.com/****/ofx-para-excel.git

# 2. Entre na pasta
cd ofx-para-excel

# 3. Crie o ambiente virtual
python -m venv venv

# 4. Ative o ambiente (Windows)
venv\Scripts\activate

# 5. Instale as dependências
pip install -r requirements.txt

# 6. Execute o programa
python main.py

```

------------------------------------------------------------------------

## 🖥 Como Usar

1. Abra o aplicativo.  
2. Informe:
   - **Número do Banco** → código do plano de contas da conta bancária  
   - **Nome do Banco**  
3. Clique em **Selecionar e Processar Arquivos OFX**  
4. Selecione os arquivos `.ofx`  
5. Escolha onde salvar  
6. Pronto! A planilha será gerada automaticamente   

---

## 📁 Estrutura do Projeto

```
ofx-para-excel/
│
├── main.py
├── requirements.txt
├── fundo.png
├── logo_cadasto.ico
└── README.md
```

---

## 🔎 Processamento dos Arquivos OFX

A lógica segue o padrão **contábil**, não bancário:

| Campo               | Regra |
|---------------------|-------|
| **débito**          | Valores **positivos** (entradas no banco) |
| **crédito**         | Valores **negativos** (saídas do banco) |
| **data**            | Convertida para o número serial do Excel |
| **valor**           | Formato brasileiro (vírgula) |
| **nome do emitente**| `DEB C/C {banco}` ou `CRED C/C {banco}` conforme o tipo |
| **complemento**     | Texto original do campo `memo` |
| **histórico total** | Montagem automática (tipo + banco + memo) |

---

## 🧾 Exemplo Completo de Conversão

### 1. Transação original no OFX

```
Data: 2024-10-05
Valor: -150.75
Memo: PAGAMENTO MERCADO LIVRE
```

Informações do usuário:

- Número do Banco: **111**
- Nome do Banco: **ITAÚ**

---

## 🔄 2. Processamento contábil

Valor negativo → **CRÉDITO** (saída)

| Campo                    | Resultado                          |
|--------------------------|-------------------------------------|
| débito                   | *(vazio)*                           |
| crédito                  | 111                                 |
| data                     | 45620                               |
| valor                    | -150,75                             |
| nome do emitente         | CRED C/C ITAÚ                       |
| complemento do histórico | PAGAMENTO MERCADO LIVRE             |
| histórico total          | CRED C/C ITAÚ PAGAMENTO MERCADO LIVRE |

---

## 📊 3. Resultado final no Excel

| debito | credito | data  | valor   | codigo do historico | n.documento | nome do emitente | complemento do historico | historico total                         |
|--------|---------|--------|---------|----------------------|--------------|------------------|---------------------------|-----------------------------------------|
|        | 111     | 45620  | -150,75 |                      |              | CRED C/C ITAÚ    | PAGAMENTO MERCADO LIVRE   | CRED C/C ITAÚ PAGAMENTO MERCADO LIVRE  |

---

## 🚧 Possíveis Melhorias

- Exportação CSV  
- Preview antes da exportação  
- Personalização da estrutura  
- Tema claro/escuro  

---

### 📝 4. Representação textual

    debito: 111
    credito:
    data: 45620
    valor: -150,75
    nome do emitente: DEB C/C ITAÚ
    complemento: PAGAMENTO MERCADO LIVRE
    histórico total: DEB C/C ITAÚ PAGAMENTO MERCADO LIVRE

------------------------------------------------------------------------

## 🚧 Possíveis Melhorias

-   Exportação CSV\
-   Preview antes da exportação\
-   Personalização do layout\
-   Tema escuro/claro\


------------------------------------------------------------------------

## 📬 Contato

  <strong>Discord:</small> juaocrl#2412<br>
  <strong>GitHub:</strong> <a href="https://github.com/juaocrl">juaocrl</a><br>
  <strong>LinkedIn:</strong> <a href="https://www.linkedin.com/in/joaovictorsmoura">João</a><br>
</div>

