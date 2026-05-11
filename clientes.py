def nome_existe(ws, nome):
    for nome_ws, _ in list(ws.values)[1:]:
        if nome == nome_ws:
            return True
    return False

def email_existe(ws, email):
    for _, email_ws in list(ws.values)[1:]:
        if email == email_ws:
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

def perguntar_continuar():
    while True:
        continuar = input("Deseja cadastrar outro cliente? (s/n): ").lower()
        if continuar in ["s", "n"]:
            return continuar == "s"
        print("Resposta inválida.")

def cadastrar_clientes(ws, wb, arquivo):
    print("_______________________________")
    print("CADASTRO DE CLIENTE")

    while True:
        nome = pedir_nome(ws)
        if nome is None:
            print("Cadastro cancelado.")
            print("_______________________________")
            return

        email = pedir_email(ws)
        if email is None:
            print("Cadastro cancelado.")
            print("_______________________________")

            return

        ws.append([nome, email])

        if not perguntar_continuar():
            break

    wb.save(arquivo)
    print("\nClientes salvos com sucesso!")
    print("_______________________________")

def listar_clientes(ws):
    print("_______________________________")
    print("LISTA DE CLIENTES\n")

    clientes = list(ws.values)[1:]

    if not clientes:
        print("Nenhum cliente cadastrado.")
    else:
        for nome, contato in clientes:
            print(f"{nome} / {contato}")

    print("_______________________________")

def buscar_cliente(ws):
    print("____________________________")
    print("BUSCAR CLIENTES\n")

    busca = input("Digite um nome que deseja pesquisar: ")

    encontrou = False

    for nome, contato in list(ws.values)[1:]:
        if busca.lower() in nome.lower():
            print(f"{nome} / {contato}")
            encontrou = True

    if not encontrou:
        print("Nome não encontrado.")

    print("____________________________")

def fechar(wb, arquivo):
    wb.save(arquivo)
    print("_______________________________")
    print("Arquivo salvo!\nFechando...")


def deletar_cliente(ws, wb, arquivo):
    print("_______________________________")
    print("LISTA DE CLIENTES\n")

    clientes = list(ws.values)[1:]

    for nome, contato in clientes:
        print(f"{nome} / {contato}")

    while True:
        nome_delete = input("\nDigite o nome que deseja deletar: ").lower()

        if nome_delete == "cancelar":
            return

        encontrou = False

        for i, (nome, contato) in enumerate(clientes, start=2):
            if nome_delete == nome.lower():
                ws.delete_rows(i)
                print(f"{nome} removido com sucesso!")
                print("_______________________________")
                wb.save(arquivo)
                encontrou = True
                break

        if not encontrou:
            print("Nome não encontrado.")
            print("_______________________________")

        break
