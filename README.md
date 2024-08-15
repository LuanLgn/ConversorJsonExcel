# Conversor JSON para Excel

Este programa permite converter arquivos JSON em planilhas Excel, com a opção de selecionar as colunas a serem exportadas. Utiliza `pandas` para manipulação de dados e `tkinter` para a interface gráfica.

## Pré-requisitos

Certifique-se de ter o Python instalado. Este programa foi testado com o Python 3.8 e deve funcionar com versões mais recentes.

### Bibliotecas necessárias

- `pandas`
- `openpyxl`
- `tqdm`
- `tkinter` (geralmente incluído com a instalação padrão do Python)

## Instalação

1. **Clone o repositório ou baixe os arquivos.**

2. **Crie e ative um ambiente virtual:**

   Abra o terminal e execute os seguintes comandos ou execute o setup.bat:

   ```sh
   @echo off
   setlocal

   REM Definir o nome da pasta do ambiente virtual
   set "venv_dir=venv"

   REM Verificar se o ambiente virtual já existe
   if exist "%venv_dir%\Scripts\activate" (
       echo Ambiente virtual já existe.
   ) else (
       echo Criando ambiente virtual...
       python -m venv %venv_dir%
   )

   echo Ativando ambiente virtual...
   call %venv_dir%\Scripts\activate

   echo Instalando dependências...
   pip install -r requirements.txt

   echo Verificando a presença do conversor.py...
   if exist conversor.py (
       echo "conversor.py encontrado."
   ) else (
       echo "conversor.py não encontrado."
       pause
       exit /b 1
   )

   echo Executando o conversor.py...
   python conversor.py

   echo Configuração e execução concluídas.
   pause

   endlocal
   ```

3. **Instale as dependências:**

   Certifique-se de que você está no ambiente virtual e execute:

   ```sh
   pip install -r requirements.txt
   ```

   `requirements.txt` deve incluir:

   ```
   pandas
   openpyxl
   tqdm
   ```

## Uso

1. **Execute o programa:**

   No terminal, com o ambiente virtual ativado, execute:

   ```sh
   python conversor.py
   ```

2. **Selecione o arquivo JSON:**

   Uma janela de diálogo será exibida solicitando que você selecione o arquivo JSON que deseja converter.

3. **Escolha as colunas para exportar:**

   Uma nova janela será exibida permitindo que você selecione as colunas a serem exportadas para o arquivo Excel. Use Ctrl+Clique ou Shift+Clique para selecionar várias colunas.

4. **Salve o arquivo Excel:**

   Após selecionar as colunas, uma janela de diálogo permitirá que você escolha o local para salvar o arquivo Excel.

5. **Concluído:**

   O programa exibirá uma mensagem informando que o arquivo foi salvo com sucesso no local escolhido.

## Notas

- O programa achata estruturas JSON aninhadas para colunas planas em um DataFrame.
- Listas dentro do JSON são convertidas em strings para garantir a compatibilidade com o formato Excel.
- Se estiver com erro de importação provavelmente seu JSON está faltando algo, mas fale comigo se estiver com problemas :D
- Feito com amor pro pessoal de Implantação S2
