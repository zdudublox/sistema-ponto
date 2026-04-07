from flask import Flask, request, render_template_string, jsonify
import pandas as pd
from datetime import datetime
import json
import os

app = Flask(__name__)

ARQ_CSV = "registros.csv"
ARQ_FUNC = "funcionarios.json"
ARQ_CODIGOS = "codigos.xlsx"  # Nome da planilha de códigos
PLANILHA_EXCEL = "registros_funcionarios.xlsx"  # Planilha separada

# Inicialização do CSV de registros
if not os.path.exists(ARQ_CSV):
    columns = ["nome", "data", "entrada", "inicio_almoco", "fim_almoco", "saida", "atividade", "horas", "saldo"]
    pd.DataFrame(columns=columns).to_csv(ARQ_CSV, index=False)

def carregar_tabela_codigos():
    """Lê a planilha de códigos e retorna um dicionário {codigo: atividade}"""
    try:
        df_cods = pd.read_excel(ARQ_CODIGOS, dtype=str)
        return dict(zip(df_cods['codigo'], df_cods['atividade']))
    except Exception as e:
        print(f"Erro ao carregar planilha de códigos: {e}")
        return {}

def carregar_funcionarios():
    try:
        with open(ARQ_FUNC, "r") as f:
            return json.load(f)
    except:
        return {}

def salvar_funcionarios(dados):
    with open(ARQ_FUNC, "w") as f:
        json.dump(dados, f, indent=4)

def registrar_excel(nome, entrada, saida, atividade, saldo=None, detalhes=""):
    """Adiciona o registro à planilha Excel no formato desejado."""
    agora = datetime.now()
    dia = agora.strftime("%d/%m/%Y")
    
    # Mapeamento de dias da semana para português
    dias_pt = {
        "Monday": "Segunda-Feira",
        "Tuesday": "Terça-Feira",
        "Wednesday": "Quarta-Feira",
        "Thursday": "Quinta-Feira",
        "Friday": "Sexta-Feira",
        "Saturday": "Sábado",
        "Sunday": "Domingo"
    }
    dia_semana = dias_pt[agora.strftime("%A")]
    
    df_novo = pd.DataFrame([{
        "Nome": nome,
        "Dia": dia,
        "Dia da semana": dia_semana,
        "Horário entrada": entrada,
        "Atividade do dia": atividade,
        "Horário saída": saida,
        "Detalhes": detalhes,
        "Saldo de horas": round(saldo,2) if saldo is not None else ""
    }])
    
    if os.path.exists(PLANILHA_EXCEL):
        df_exist = pd.read_excel(PLANILHA_EXCEL)
        df_novo = pd.concat([df_exist, df_novo], ignore_index=True)
    
    df_novo.to_excel(PLANILHA_EXCEL, index=False)

@app.route("/")
def index():
    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Sistema de Ponto Ollintekers</title>
        <meta charset="UTF-8">
        <style>
            :root { --bg: #0f172a; --card: #1e293b; --primary: #38bdf8; --text: #f8fafc; }
            body { font-family: 'Segoe UI', sans-serif; background: var(--bg); color: var(--text); display: flex; justify-content: center; padding: 20px; }
            .container { width: 100%; max-width: 450px; }
            .card { background: var(--card); padding: 20px; border-radius: 15px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); margin-bottom: 20px; }
            h2 { color: var(--primary); margin: 0 0 15px 0; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 1px; }
            input, select { width: 100%; padding: 12px; margin: 8px 0; border-radius: 8px; border: 1px solid #334155; background: #0f172a; color: white; box-sizing: border-box; }
            .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
            button { width: 100%; padding: 12px; border-radius: 8px; border: none; font-weight: bold; cursor: pointer; transition: 0.2s; color: white; }
            .btn-blue { background: #0284c7; }
            .btn-green { background: #16a34a; }
            .btn-orange { background: #ea580c; }
            .btn-red { background: #dc2626; grid-column: span 2; margin-top: 10px; }
            button:hover { filter: brightness(1.2); }
            #area-saida { display: none; grid-column: span 2; background: #00000033; padding: 15px; border-radius: 8px; margin-top: 10px; border: 1px dashed #475569; }
            #status { display: none; padding: 15px; border-radius: 8px; text-align: center; font-weight: bold; margin-top: 15px; border-left: 5px solid; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1 style="text-align:center">Portal de Ponto</h1>
            
            <div class="card">
                <h2>Cadastro</h2>
                <input type="text" id="reg_nome" placeholder="Nome Completo">
                <select id="reg_tipo">
                    <option value="clt">CLT (8h)</option>
                    <option value="estagio">Estagiario (6h)</option>
                </select>
                <button class="btn-blue" onclick="enviar('cadastrar')">Cadastrar</button>
            </div>

            <div class="card">
                <h2>Registro</h2>
                <input type="text" id="ponto_nome" placeholder="Nome do Funcionario">
                <div class="grid">
                    <button class="btn-green" onclick="enviar('entrada')">Entrada</button>
                    <button class="btn-orange" onclick="enviar('inicio_almoco')">Ini. Almoco</button>
                    <button class="btn-orange" onclick="enviar('fim_almoco')">Fim Almoco</button>
                    <button class="btn-red" id="btn-pre-saida" onclick="mostrarSaida()">Registrar Saida</button>
                    
                    <div id="area-saida">
                        <label style="font-size: 11px;">Codigo Atividade (Consulte a Planilha):</label>
                        <input type="text" id="ponto_codigo" placeholder="Ex: 7152">
                        <label style="font-size: 11px;">Almoco > 2h, usar banco?</label>
                        <select id="usar_banco">
                            <option value="s">Sim</option>
                            <option value="n">Nao</option>
                        </select>
                        <button class="btn-red" style="grid-column: auto;" onclick="enviar('saida')">Confirmar Saida</button>
                    </div>
                </div>
            </div>
            <div id="status"></div>
        </div>

        <script>
            function mostrarSaida() {
                document.getElementById('area-saida').style.display = 'block';
                document.getElementById('btn-pre-saida').style.display = 'none';
            }
            async function enviar(rota) {
                const statusDiv = document.getElementById('status');
                const payload = {
                    nome: rota === 'cadastrar' ? document.getElementById('reg_nome').value : document.getElementById('ponto_nome').value,
                    tipo: document.getElementById('reg_tipo').value,
                    codigo: document.getElementById('ponto_codigo').value,
                    usar_banco: document.getElementById('usar_banco').value
                };
                if (!payload.nome) { alert("Nome obrigatorio."); return; }
                try {
                    const response = await fetch('/' + rota, {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });
                    const data = await response.json();
                    statusDiv.style.display = 'block';
                    statusDiv.innerText = data.msg;
                    statusDiv.style.backgroundColor = data.ok ? '#064e3b' : '#450a0a';
                    statusDiv.style.borderColor = data.ok ? '#22c55e' : '#ef4444';
                    if(data.ok && rota === 'saida') {
                        document.getElementById('area-saida').style.display = 'none';
                        document.getElementById('btn-pre-saida').style.display = 'block';
                        document.getElementById('ponto_codigo').value = '';
                    }
                } catch (e) { alert("Erro de conexao."); }
            }
        </script>
    </body>
    </html>
    """)

@app.route("/cadastrar", methods=["POST"])
def cadastrar():
    data = request.json
    nome = data['nome'].strip().lower()
    f = carregar_funcionarios()
    f[nome] = {"tipo": data['tipo']}
    salvar_funcionarios(f)
    return jsonify({"ok": True, "msg": "Cadastrado."})

@app.route("/entrada", methods=["POST"])
def entrada():
    nome = request.json['nome'].strip().lower()
    if nome not in carregar_funcionarios(): return jsonify({"ok": False, "msg": "Nao cadastrado."})
    df = pd.read_csv(ARQ_CSV, dtype=str).fillna("")
    for i in range(len(df)-1, -1, -1):
        if df.loc[i, "nome"] == nome and df.loc[i, "saida"] == "":
            return jsonify({"ok": False, "msg": "Ja existe uma entrada aberta!"})
    agora = datetime.now()
    df.loc[len(df)] = [nome, agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M"), "", "", "", "", "", ""]
    df.to_csv(ARQ_CSV, index=False)
    return jsonify({"ok": True, "msg": "Entrada registrada."})

@app.route("/inicio_almoco", methods=["POST"])
def ini_almoco():
    nome = request.json['nome'].strip().lower()
    df = pd.read_csv(ARQ_CSV, dtype=str).fillna("")
    for i in range(len(df)-1, -1, -1):
        if df.loc[i, "nome"] == nome and df.loc[i, "saida"] == "":
            df.loc[i, "inicio_almoco"] = datetime.now().strftime("%H:%M")
            df.to_csv(ARQ_CSV, index=False)
            return jsonify({"ok": True, "msg": "Inicio almoco registrado."})
    return jsonify({"ok": False, "msg": "Entrada nao encontrada."})

@app.route("/fim_almoco", methods=["POST"])
def fim_almoco():
    nome = request.json['nome'].strip().lower()
    df = pd.read_csv(ARQ_CSV, dtype=str).fillna("")
    for i in range(len(df)-1, -1, -1):
        if df.loc[i, "nome"] == nome and df.loc[i, "saida"] == "" and df.loc[i, "inicio_almoco"] != "":
            df.loc[i, "fim_almoco"] = datetime.now().strftime("%H:%M")
            df.to_csv(ARQ_CSV, index=False)
            return jsonify({"ok": True, "msg": "Fim almoco registrado."})
    return jsonify({"ok": False, "msg": "Inicio de almoco nao encontrado."})

@app.route("/saida", methods=["POST"])
def saida():
    req = request.json
    nome, codigo_digitado, usar_banco = req['nome'].strip().lower(), req['codigo'].strip(), req['usar_banco']
    
    tabela_atividades = carregar_tabela_codigos()
    if codigo_digitado not in tabela_atividades:
        return jsonify({"ok": False, "msg": "CODIGO INVALIDO! Consulte a planilha de atividades."})
    
    atividade_nome = tabela_atividades[codigo_digitado]
    funcs = carregar_funcionarios()
    df = pd.read_csv(ARQ_CSV, dtype=str).fillna("")
    
    for i in range(len(df)-1, -1, -1):
        if df.loc[i, "nome"] == nome and df.loc[i, "saida"] == "":
            if df.loc[i, "inicio_almoco"] != "" and df.loc[i, "fim_almoco"] == "":
                return jsonify({"ok": False, "msg": "Finalize o almoco antes de sair!"})

            agora = datetime.now()
            hora_s = agora.strftime("%H:%M")
            
            ent_dt = datetime.strptime(df.loc[i, "entrada"], "%H:%M")
            sai_dt = datetime.strptime(hora_s, "%H:%M")
            
            almoco = 0
            ia, fa = df.loc[i, "inicio_almoco"], df.loc[i, "fim_almoco"]
            if ia != "" and fa != "":
                almoco = (datetime.strptime(fa, "%H:%M") - datetime.strptime(ia, "%H:%M")).total_seconds() / 3600
                if almoco > 2 and usar_banco != "s": almoco = 2
            
            total = (sai_dt - ent_dt).total_seconds() / 3600 - almoco
            carga = 8 if funcs[nome]["tipo"] == "clt" else 6
            saldo = total - carga

            df.loc[i, "saida"] = hora_s
            df.loc[i, "atividade"] = atividade_nome
            df.loc[i, "horas"] = str(round(total, 2))
            df.loc[i, "saldo"] = str(round(saldo, 2))
            df.to_csv(ARQ_CSV, index=False)

            # --- REGISTRO NO EXCEL ---
            detalhes = f"Almoço: {ia}-{fa}" if ia and fa else ""
            registrar_excel(nome=nome, entrada=df.loc[i, "entrada"], saida=hora_s, atividade=atividade_nome, saldo=saldo, detalhes=detalhes)

            return jsonify({"ok": True, "msg": f"Saida registrada: {atividade_nome}. Saldo: {round(saldo,2)}"})
                
    return jsonify({"ok": False, "msg": "Entrada nao encontrada."})

if __name__ == "__main__":
    app.run(debug=True)
