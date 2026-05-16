# 📇 Sistema de Gerenciamento de Clientes

Projeto desenvolvido em Python para gerenciamento de clientes utilizando planilhas Excel com a biblioteca OpenPyXL.

O sistema funciona via terminal e permite o cadastro, consulta e gerenciamento de clientes de forma simples e organizada.

## ✨ Funcionalidades

- 👤 Cadastro de clientes
- 📧 Validação de e-mails
- 📞 Cadastro e validação de telefone
- 🔍 Busca de clientes
- 📄 Listagem de clientes
- ❌ Remoção de clientes
- 🚫 Verificação de dados duplicados
- 💾 Salvamento automático em planilha `.xlsx`

---

# 🚀 Tecnologias utilizadas

- 🐍 Python
- 📗 OpenPyXL
- 📂 Excel (`.xlsx`)

---

# 📁 Estrutura do projeto

```txt
Sistema-Gerenciamento-Clientes/
│
├── main.py
├── clientes.py
├── README.md
│
└── data/
    └── tabelaclientes.xlsx
```

---

# ⚙️ Funcionalidades do sistema

- ✅ Cadastro de clientes
- ✅ Validação de e-mails
- ✅ Cadastro de telefone
- ✅ Validação de telefone
- ✅ Verificação de nomes duplicados
- ✅ Verificação de e-mails duplicados
- ✅ Verificação de telefones duplicados
- ✅ Busca de clientes pelo nome
- ✅ Listagem completa dos clientes
- ✅ Exclusão de clientes
- ✅ Salvamento automático em planilha Excel
- ✅ Cancelamento de operações

---

# 📄 Como os dados são armazenados

Os clientes são armazenados em uma planilha Excel (`.xlsx`) contendo:

| Nome | Email | Telefone |
|------|--------|-----------|
| Cliente | email@email.com | 11999999999 |

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
- validação de telefone
- verificação de duplicidade
- busca de clientes
- exclusão de clientes
- salvamento dos dados

---

# 🖥️ Menu do sistema

```txt
MENU
1 - Cadastrar clientes
2 - Listar clientes
3 - Buscar cliente
4 - Sair
5 - Deletar cliente
```

---

# ▶️ Como executar o projeto

## 1️⃣ Instale o Python

Baixe em:

https://www.python.org/

---

## 2️⃣ Instale a biblioteca OpenPyXL

```bash
pip install openpyxl
```

---

## 3️⃣ Execute o sistema

```bash
python main.py
```

---

# 🧠 Conceitos praticados

Este projeto foi desenvolvido com foco no aprendizado de:

- 🧩 Modularização
- 📂 Manipulação de arquivos Excel
- 🔄 Estruturas de repetição (`while`)
- ⚙️ Funções
- 🧠 Validação de dados
- 📋 CRUD básico
- 💾 Persistência de dados
- 🐍 Estruturação de projetos Python
- 📊 Manipulação de planilhas com OpenPyXL

---

# 🔮 Melhorias futuras

- ✏️ Edição de clientes
- 🔒 Validação avançada de e-mails
- 📊 Exportação de relatórios
- 🔎 Busca avançada
- 📑 Ordenação de clientes
- 🗄️ Migração para banco de dados SQL
- 🌐 Interface gráfica/web

---

# 👨‍💻 Autor

**Murilo Beluci Nunes**
