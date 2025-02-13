# Conversor XML para XLSX

Este projeto converte arquivos XML exportados do AutoCAD em arquivos XLSX filtrando materiais específicos, como MDF e portas. O sistema possui uma interface gráfica simples feita com Tkinter.

## Funcionalidades

Filtragem de materiais relevantes (MDF e portas)
Conversão automática de dimensões
Exportação dos dados para uma planilha Excel (XLSX)
Interface gráfica para facilitar a conversão
Login de acesso ao sistema

## Requisitos
Para executar este projeto, é necessário ter instalado:
Python 3.x
Bibliotecas: openpyxl, tkinter

### Instale as dependências com:
pip install openpyxl

## Como executar

Para rodar o conversor, execute o seguinte comando:
python conversor.py

### Como transformar em um executável
Para criar um executável do projeto, utilize o PyInstaller:
Instale o PyInstaller, caso ainda não tenha:
pip install pyinstaller
Gere o executável com o seguinte comando:
pyinstaller --onefile --windowed conversor.py

Isso criará um executável na pasta dist/. Para rodá-lo, basta acessar a pasta dist/ e executar o arquivo gerado.

## Contato
Desenvolvido por Enzo Martins. Para mais informações, acesse enzomartinsdev.com. Ou entre em contato em (31) 99521-8418

