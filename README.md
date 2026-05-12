# 📋 Sistema de Cadastro de Clientes

Projeto desenvolvido em Python para gerenciamento de clientes utilizando planilhas Excel com a biblioteca OpenPyXL.

O sistema funciona via terminal e permite:

- 👤 Cadastro de clientes
- 📧 Validação de e-mails
- 🔍 Busca de clientes
- 📄 Listagem de clientes
- ❌ Remoção de clientes
- 💾 Salvamento automático em planilha `.xlsx`

---

# 🚀 Tecnologias utilizadas

- 🐍 Python
- 📗 OpenPyXL
- 📂 Excel (.xlsx)

---

# 📁 Estrutura do projeto

```txt
CadastroClientes/
│
├── main.py
├── clientes.py
├── README.md
│
└── data/
    └── tabelaclientes.xlsx
```

---

# ⚙️ Funcionalidades

- ✅ Cadastro de clientes
- ✅ Validação de e-mails
- ✅ Verificação de nomes duplicados
- ✅ Verificação de e-mails duplicados
- ✅ Busca de clientes pelo nome
- ✅ Listagem completa dos clientes
- ✅ Exclusão de clientes
- ✅ Salvamento automático em planilha Excel

---

# 📄 Como os dados são armazenados

Os clientes são armazenados em uma planilha Excel (`.xlsx`) contendo:

| Nome | Email |
|------|------|
| Cliente | email@email.com |

---

# 🧩 Organização dos arquivos

## 📌 `main.py`

Responsável pelo fluxo principal do sistema:
- menu interativo
- navegação entre opções
- carregamento da planilha
- inicialização do sistema

---

## 📌 `clientes.py`

Contém toda a lógica do sistema:
- cadastro de clientes
- validação de e-mails
- verificação de duplicidade
- busca de clientes
- exclusão de clientes
- salvamento dos dados

---

# 🖥️ Menu do sistema

```txt
MENU
1 - Cadastrar clientes.
2 - Lista de clientes.
3 - Buscar clientes.
4 - Sair.
5 - Deletar cliente.
```

---

# ▶️ Como executar o projeto

### 1️⃣ Instale o Python

Baixe em:
https://www.python.org/

---

### 2️⃣ Instale a biblioteca OpenPyXL

```bash
pip install openpyxl
```

---

### 3️⃣ Execute o sistema

```bash
python main.py
```

---

# 🧠 Conceitos praticados

Este projeto foi desenvolvido com foco em aprendizado de:

- 🧩 Modularização
- 📂 Manipulação de arquivos Excel
- 🔄 Loops
- ⚙️ Funções
- 🧠 Validações
- 📋 CRUD básico
- 💾 Persistência de dados
- 🐍 Estruturação de projetos Python

---

# 🔮 Melhorias futuras

- 🔒 Validação avançada de e-mails
- 📞 Cadastro de telefone
- 📊 Exportação de relatórios
- 🔎 Busca avançada
- 🗄️ Migração para banco de dados SQL

---

# 👨‍💻 Autor

**Murilo Beluci Nunes**
