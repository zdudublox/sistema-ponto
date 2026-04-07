import pandas as pd
import json
from datetime import datetime

ARQ_CSV = "registros.csv"
ARQ_FUNC = "funcionarios.json"

codigos = {
    "7152": "Correção de anúncios",
    "1234": "Atendimento"
}

def carregar_funcionarios():
    try:
        with open(ARQ_FUNC, "r") as f:
            return json.load(f)
    except:
        return {}

def salvar_funcionarios(dados):
    with open(ARQ_FUNC, "w") as f:
        json.dump(dados, f, indent=4)

def cadastrar_funcionario():
    dados = carregar_funcionarios()

    nome = input("Nome: ").strip().lower()
    tipo = input("Tipo (clt/estagio): ").strip().lower()

    if tipo not in ["clt", "estagio"]:
        print("Tipo inválido!")
        return

    dados[nome] = {"tipo": tipo}
    salvar_funcionarios(dados)

    print("Funcionário cadastrado!")

def entrada():
    dados = carregar_funcionarios()
    nome = input("Nome: ").strip().lower()

    if nome not in dados:
        print("Funcionário não cadastrado!")
        return

    agora = datetime.now()
    data = agora.strftime("%d/%m/%Y")
    hora = agora.strftime("%H:%M")

    df = pd.read_csv(ARQ_CSV, dtype=str)

    df.loc[len(df)] = [nome, data, hora, "", "", "", "", "", ""]
    df.to_csv(ARQ_CSV, index=False)

    print("Entrada registrada!")

def inicio_almoco():
    nome = input("Nome: ").strip().lower()
    df = pd.read_csv(ARQ_CSV, dtype=str)

    for i in range(len(df)-1, -1, -1):
        nome_csv = str(df.loc[i, "nome"]).strip().lower()
        inicio = str(df.loc[i, "inicio_almoco"]).strip()

        if nome_csv == nome and (inicio == "" or inicio == "nan"):
            df.loc[i, "inicio_almoco"] = datetime.now().strftime("%H:%M")
            df.to_csv(ARQ_CSV, index=False)
            print("Início do almoço registrado!")
            return

    print("Registro não encontrado!")

def fim_almoco():
    nome = input("Nome: ").strip().lower()
    df = pd.read_csv(ARQ_CSV, dtype=str)

    for i in range(len(df)-1, -1, -1):
        nome_csv = str(df.loc[i, "nome"]).strip().lower()
        fim = str(df.loc[i, "fim_almoco"]).strip()

        if nome_csv == nome and (fim == "" or fim == "nan"):
            df.loc[i, "fim_almoco"] = datetime.now().strftime("%H:%M")
            df.to_csv(ARQ_CSV, index=False)
            print("Fim do almoço registrado!")
            return

    print("Registro não encontrado!")

def saida():
    dados = carregar_funcionarios()
    nome = input("Nome: ").strip().lower()

    if nome not in dados:
        print("Funcionário não cadastrado!")
        return

    df = pd.read_csv(ARQ_CSV, dtype=str)

    for i in range(len(df)-1, -1, -1):
        nome_csv = str(df.loc[i, "nome"]).strip().lower()
        saida_csv = str(df.loc[i, "saida"]).strip()

        if nome_csv == nome and (saida_csv == "" or saida_csv == "nan"):
            agora = datetime.now()
            hora_saida = agora.strftime("%H:%M")

            df.loc[i, "saida"] = hora_saida

            codigo = input("Código: ")
            atividade = codigos.get(codigo, "Desconhecido")
            df.loc[i, "atividade"] = atividade

            entrada = datetime.strptime(df.loc[i, "entrada"], "%H:%M")
            saida = datetime.strptime(hora_saida, "%H:%M")

            # cálculo do almoço
            inicio_a = df.loc[i, "inicio_almoco"]
            fim_a = df.loc[i, "fim_almoco"]

            if inicio_a and fim_a and inicio_a != "nan" and fim_a != "nan":
                inicio_a = datetime.strptime(inicio_a, "%H:%M")
                fim_a = datetime.strptime(fim_a, "%H:%M")
                almoco = (fim_a - inicio_a).seconds / 3600

                # REGRA EMPRESA
                if almoco > 2:
                    usar_banco = input("Almoço passou de 2h. Usar banco de horas? (s/n): ").lower()

                    if usar_banco != "s":
                        almoco = 2
            else:
                almoco = 0

            horas = (saida - entrada).seconds / 3600 - almoco
            if horas < 0:
                horas = 0

            carga = 8 if dados[nome]["tipo"] == "clt" else 6
            saldo = horas - carga

            df.loc[i, "horas"] = str(round(horas, 2))
            df.loc[i, "saldo"] = str(round(saldo, 2))

            df.to_csv(ARQ_CSV, index=False)

            print(f"Horas: {round(horas,2)} | Saldo: {round(saldo,2)}")
            return

    print("Nenhuma entrada encontrada!")

def menu():
    while True:
        print("\n1 - Cadastrar funcionário")
        print("2 - Entrada")
        print("3 - Início do almoço")
        print("4 - Fim do almoço")
        print("5 - Saída")
        print("6 - Sair")

        op = input("Escolha: ")

        if op == "1":
            cadastrar_funcionario()
        elif op == "2":
            entrada()
        elif op == "3":
            inicio_almoco()
        elif op == "4":
            fim_almoco()
        elif op == "5":
            saida()
        elif op == "6":
            break

menu()