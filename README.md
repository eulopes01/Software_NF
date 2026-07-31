# Expense PDF Merger

Aplicativo desktop desenvolvido para facilitar o controle de despesas, a organização de notas fiscais em PDF e o preenchimento automático de planilhas de controle.

---

## Sobre o projeto

O **Expense PDF Merger** surgiu a partir de uma necessidade do meu trabalho. Todos os meses eu precisava organizar as notas fiscais e comprovantes de despesas, deixando tudo na mesma ordem da fatura antes de enviar para o setor financeiro.

Para automatizar esse processo, decidi desenvolver uma aplicação desktop onde fosse possível registrar cada despesa, anexar o PDF correspondente e, ao final, gerar um único arquivo contendo todos os documentos na ordem correta, além de **lançar todos esses dados automaticamente em uma planilha do Excel**.

Optei por armazenar todos os dados localmente por questões de segurança, evitando que informações financeiras fossem enviadas para serviços externos.

O projeto foi desenvolvido em **Python**, utilizando **CustomTkinter** para a interface gráfica. Durante o desenvolvimento, também utilizei ferramentas de IA como apoio para pesquisa, solução de dúvidas e aceleração do desenvolvimento, sempre revisando e adaptando o código conforme a necessidade.

---

## Funcionalidades

- Junta vários PDFs em um único arquivo.
- **Automação de Excel:** Preenche automaticamente a planilha de controle (preservando fórmulas e formatações originais) e atualiza os dados sem duplicar lançamentos.
- Organiza automaticamente os documentos e lançamentos por data.
- Cadastro de fornecedores com preenchimento automático.
- Aceita datas digitadas sem necessidade de barras (ex.: `0108` → `01/08/2026`).
- Armazena despesas e fornecedores localmente.
- Destaque visual (linhas vermelhas) para despesas que ainda não possuem PDF anexado.
- **Alerta visual de segurança** ao entrar no "Modo de Edição", evitando confusões na alteração de registros.
- Permite adicionar, editar e excluir registros.
- Funciona totalmente offline.
- Pode ser distribuído como um executável (.exe), sem necessidade de instalar Python.

---

## Estrutura do projeto

```text
├── domain/
│   └── expense.py
├── services/
│   ├── pdf_service.py
│   └── excel_service.py
├── utils/
│   └── validators.py
├── ui/
│   └── app_window.py
├── main.py
└── requirements.txt
````
Cada pasta possui uma responsabilidade específica, facilitando a manutenção e futuras melhorias seguindo os princípios de Clean Architecture.

## Executando o projeto
Clone o repositório.

Crie um ambiente virtual:

PowerShell
```
python -m venv venv
````
Ative o ambiente:

PowerShell
````
.\venv\Scripts\Activate.ps1
````
Instale as dependências:

PowerShell
````
pip install -r requirements.txt
````
Execute o programa:

PowerShell
python main.py
Gerando o executável
PowerShell
````
python -m PyInstaller --noconsole --onefile --name "ExpenseApp" --add-data "C:\caminho\para\customtkinter;customtkinter" main.py
````
O executável será criado na pasta:

Plaintext
dist/
Armazenamento dos dados
Os dados são armazenados localmente em:

Plaintext
````
%APPDATA%\ExpenseApp
````
Nesse diretório ficam os arquivos:
````
expenses.json

vendors.json
````
Essa abordagem mantém o histórico do usuário mesmo após fechar o programa e evita que informações fiquem espalhadas pelo computador.

## Tecnologias utilizadas
````
Python

CustomTkinter

PyPDF2

openpyxl (Automação de planilhas Excel)

PyInstaller
````
Certificados
Os conhecimentos utilizados no desenvolvimento deste projeto foram adquiridos por meio de cursos e estudos práticos.

📜 Python Completo - Danki Code

➡️ Visualizar certificado:

https://shre.ink/Certificado-Python

Status
Versão: 1.0.0

Projeto funcional desenvolvido para automatizar a organização de despesas, a consolidação de documentos em PDF e o preenchimento de relatórios no Excel.
