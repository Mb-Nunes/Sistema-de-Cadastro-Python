def linha():
    print("_" * 30)

def pausar():
    input("Pressione Enter para continuar...")
    linha()

def nome_existe(ws, nome):
    for nome_ws, _, _ in list(ws.values)[1:]:
        if nome == nome_ws:
            return True
    return False

def email_existe(ws, email):
    for _, email_ws, _ in list(ws.values)[1:]:
        if email == email_ws:
            return True
    return False

def telefone_existe(ws, telefone):
    for _, _, telefone_ws in list(ws.values)[1:]:
        if telefone == telefone_ws:
            return True
    return False

def pedir_nome(ws):
    while True:
        nome = input("Digite o nome do cliente: ").title()

        if nome.lower() == "cancelar":
            return None

        if nome_existe(ws, nome):
            print("Nome já cadastrado!")
        else:
            return nome

def pedir_email(ws):
    while True:
        email = input("Digite o e-mail do cliente: ")

        if email.lower() == "cancelar":
            return None

        if "@" not in email or "." not in email or email.count("@") != 1:
            print("Email inválido")
            continue

        if email_existe(ws, email):
            print("Email já cadastrado!")
            continue

        return email

def pedir_telefone(ws):
    while True:
        telefone = input("Digite o número do cliente com o DDD: ")

        if telefone.lower() == "cancelar":
            return None

        if not telefone.isdigit():
            print(
                "Digite apenas números."
            )
            continue

        if telefone_existe(ws, telefone):
            print("Número já cadastrado!")
            continue
        else:
            return telefone

def perguntar_continuar():
    while True:
        continuar = input("Deseja cadastrar outro cliente? (s/n): ").lower()
        if continuar in ["s", "n"]:
            return continuar == "s"
        print("Resposta inválida.")

def cadastrar_clientes(ws, wb, arquivo):
    linha()
    print("CADASTRO DE CLIENTE")

    while True:
        nome = pedir_nome(ws)
        if nome is None:
            print("Cadastro cancelado.")
            linha()
            return

        email = pedir_email(ws)
        if email is None:
            print("Cadastro cancelado.")
            linha()
            return

        telefone = pedir_telefone(ws)
        if telefone is None:
            print("Cadastro cancelado.")
            linha()
            return

        ws.append([nome, email, telefone])

        if not perguntar_continuar():
            break

    wb.save(arquivo)
    print("\nClientes salvos com sucesso!")
    linha()

def listar_clientes(ws):
    linha()
    print("LISTA DE CLIENTES\n")

    clientes = list(ws.values)[1:]

    if not clientes:
        print("Nenhum cliente cadastrado.")
    else:
        for nome, contato, telefone in clientes:
            print(f"{nome} / {contato} / {telefone}")

    linha()

def buscar_cliente(ws):
    linha()
    print("BUSCAR CLIENTES\n")

    busca = input("Digite um nome que deseja pesquisar: ")

    encontrou = False

    for nome, contato, telefone in list(ws.values)[1:]:
        if busca.lower() in nome.lower():
            print(f"{nome} / {contato} / {telefone}")
            encontrou = True

    if not encontrou:
        print("Nome não encontrado.")

    linha()

def fechar(wb, arquivo):
    wb.save(arquivo)
    linha()
    print("Arquivo salvo!\nFechando...")


def deletar_cliente(ws, wb, arquivo):
    linha()
    print("LISTA DE CLIENTES\n")

    clientes = list(ws.values)[1:]

    for nome, contato, telefone in clientes:
        print(f"{nome} / {contato} / {telefone}")

    while True:
        nome_delete = input("\nDigite o nome que deseja deletar: ").lower()

        if nome_delete == "cancelar":
            return

        encontrou = False

        for i, (nome, contato, telefone) in enumerate(clientes, start=2):
            if nome_delete == nome.lower():
                ws.delete_rows(i)
                print(f"{nome} removido com sucesso!")
                linha()
                wb.save(arquivo)
                encontrou = True
                break

        if not encontrou:
            print("Nome não encontrado.")
            linha()

        break
