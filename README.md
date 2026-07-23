# Expense PDF Merger

Aplicativo desktop desenvolvido para facilitar o controle de despesas e a organização de notas fiscais em PDF.

---

## Sobre o projeto

O **Expense PDF Merger** surgiu a partir de uma necessidade do meu trabalho. Todos os meses eu precisava organizar as notas fiscais e comprovantes de despesas, deixando tudo na mesma ordem da fatura antes de enviar para o setor financeiro.

Para automatizar esse processo, decidi desenvolver uma aplicação desktop onde fosse possível registrar cada despesa, anexar o PDF correspondente e, ao final, gerar um único arquivo contendo todos os documentos na ordem correta.

Optei por armazenar todos os dados localmente por questões de segurança, evitando que informações financeiras fossem enviadas para serviços externos.

O projeto foi desenvolvido em **Python**, utilizando **CustomTkinter** para a interface gráfica. Durante o desenvolvimento, também utilizei ferramentas de IA como apoio para pesquisa, solução de dúvidas e aceleração do desenvolvimento, sempre revisando e adaptando o código conforme a necessidade.

---

## Funcionalidades

- Junta vários PDFs em um único arquivo.
- Organiza automaticamente os documentos por data.
- Cadastro de fornecedores com preenchimento automático.
- Aceita datas digitadas sem necessidade de barras (ex.: `0108` → `01/08/2026`).
- Armazena despesas e fornecedores localmente.
- Destaque visual para despesas que ainda não possuem PDF anexado.
- Permite adicionar, editar e excluir registros.
- Funciona totalmente offline.
- Pode ser distribuído como um executável (.exe), sem necessidade de instalar Python.

---

## Estrutura do projeto

```text
├── domain/
│   └── expense.py
├── services/
│   └── pdf_service.py
├── utils/
│   └── validators.py
├── ui/
│   └── app_window.py
├── main.py
└── requirements.txt
```

Cada pasta possui uma responsabilidade específica, facilitando a manutenção e futuras melhorias.

---

## Executando o projeto

1. Clone o repositório.

2. Crie um ambiente virtual:

```powershell
python -m venv venv
```

3. Ative o ambiente:

```powershell
.\venv\Scripts\Activate.ps1
```

4. Instale as dependências:

```powershell
pip install -r requirements.txt
```

5. Execute o programa:

```powershell
python main.py
```

---

## Gerando o executável

```powershell
python -m PyInstaller --noconsole --onefile --name "ExpenseApp" --add-data "C:\caminho\para\customtkinter;customtkinter" main.py
```

O executável será criado na pasta:

```text
dist/
```

---

## Armazenamento dos dados

Os dados são armazenados localmente em:

```text
%APPDATA%\ExpenseApp
```

Nesse diretório ficam os arquivos:

- `expenses.json`
- `vendors.json`

Essa abordagem mantém o histórico do usuário mesmo após fechar o programa e evita que informações fiquem espalhadas pelo computador.

---

## Tecnologias utilizadas

- Python
- CustomTkinter
- PyPDF2
- PyInstaller

---

## Certificados

Os conhecimentos utilizados no desenvolvimento deste projeto foram adquiridos por meio de cursos e estudos práticos.

📜 **Python Completo - Danki Code**

➡️ **Visualizar certificado:**

**https://mineraleng-my.sharepoint.com/:b:/g/personal/gabriel_sousa_mineral_eng_br/IQDrbDrh2j89RpIUzRZ0tjIbAS2nS3PyuvdwW7V-5wkWGUY?e=BibVhn](https://drive.google.com/file/d/1YH8BqCvM4AVb1g4437YPVxnryH9IH9YU/view?usp=sharing**

---

## Status

**Versão:** 1.0.0

Projeto funcional desenvolvido para automatizar a organização de despesas e a consolidação de documentos em PDF.
