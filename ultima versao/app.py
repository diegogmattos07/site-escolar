from flask import (
    Flask, render_template, request,
    redirect, url_for, session, flash, jsonify, Response
)
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from datetime import date
import csv
import io
import sqlite3
import os
import shutil
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "altere-esta-chave-antes-de-publicar-o-sistema")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
BANCO = "sistema_escolar.db"


MATERIAS = [
    "Língua Portuguesa",
    "Matemática",
    "Ciências",
    "História",
    "Geografia",
    "Língua Inglesa",
    "Arte",
    "Educação Física",
    "Ensino Religioso",
    "Tecnologia e Inovação"
]

SERIES_VALIDAS = ["1", "2", "3", "4", "5"]

# Acima desse percentual de ausência o aluno entra em alerta.
LIMITE_AUSENCIA = 25

DIAS_SEMANA = [
    "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
    "sexta-feira", "sábado", "domingo"
]

MESES = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]


# =====================================================
# CONEXÃO COM O BANCO DE DADOS
# =====================================================

def conectar():
    conexao = sqlite3.connect(BANCO)
    conexao.row_factory = sqlite3.Row
    conexao.execute("PRAGMA foreign_keys = ON")
    return conexao


def criar_banco():
    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS professores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            senha TEXT NOT NULL,
            materia TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alunos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            serie INTEGER NOT NULL,
            materia TEXT NOT NULL,
            ano INTEGER NOT NULL,
            responsavel TEXT,
            telefone TEXT,
            observacoes TEXT,
            nota1 REAL,
            nota2 REAL,
            nota3 REAL,
            nota4 REAL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS frequencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            aluno_id INTEGER NOT NULL,
            data TEXT NOT NULL,
            presente INTEGER NOT NULL,
            UNIQUE(aluno_id, data),
            FOREIGN KEY (aluno_id)
                REFERENCES alunos(id)
                ON DELETE CASCADE
        )
    """)

    # Atualiza bancos criados antes da existência do ano letivo:
    # os alunos que já estavam cadastrados ficam no ano atual.
    colunas = [
        coluna["name"]
        for coluna in cursor.execute(
            "PRAGMA table_info(alunos)"
        ).fetchall()
    ]

    if "ano" not in colunas:
        cursor.execute(
            f"ALTER TABLE alunos "
            f"ADD COLUMN ano INTEGER NOT NULL "
            f"DEFAULT {int(date.today().year)}"
        )

    # Novos campos opcionais do aluno (bancos antigos são atualizados).
    for coluna_nova in ["responsavel", "telefone", "observacoes"]:
        if coluna_nova not in colunas:
            cursor.execute(
                f"ALTER TABLE alunos ADD COLUMN {coluna_nova} TEXT"
            )

    # Perfil de acesso: o primeiro usuário cadastrado é administrador.
    colunas_prof = [c["name"] for c in cursor.execute("PRAGMA table_info(professores)").fetchall()]
    if "perfil" not in colunas_prof:
        cursor.execute("ALTER TABLE professores ADD COLUMN perfil TEXT NOT NULL DEFAULT 'professor'")
    cursor.execute("""CREATE TABLE IF NOT EXISTS eventos (
        id INTEGER PRIMARY KEY AUTOINCREMENT, titulo TEXT NOT NULL,
        descricao TEXT, data TEXT NOT NULL, tipo TEXT NOT NULL DEFAULT 'Evento',
        professor_id INTEGER NOT NULL, ano INTEGER NOT NULL,
        FOREIGN KEY(professor_id) REFERENCES professores(id) ON DELETE CASCADE
    )""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS comunicacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, aluno_id INTEGER,
        assunto TEXT NOT NULL, mensagem TEXT NOT NULL, canal TEXT NOT NULL,
        data_envio TEXT NOT NULL, professor_id INTEGER NOT NULL,
        FOREIGN KEY(aluno_id) REFERENCES alunos(id) ON DELETE SET NULL,
        FOREIGN KEY(professor_id) REFERENCES professores(id) ON DELETE CASCADE
    )""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS auditoria (
        id INTEGER PRIMARY KEY AUTOINCREMENT, professor_id INTEGER,
        acao TEXT NOT NULL, detalhes TEXT, criado_em TEXT NOT NULL,
        FOREIGN KEY(professor_id) REFERENCES professores(id) ON DELETE SET NULL
    )""")
    # Em instalações antigas, promove o primeiro usuário ao perfil administrador.
    if cursor.execute("SELECT COUNT(*) FROM professores WHERE perfil='administrador'").fetchone()[0] == 0:
        primeiro = cursor.execute("SELECT id FROM professores ORDER BY id LIMIT 1").fetchone()
        if primeiro:
            cursor.execute("UPDATE professores SET perfil='administrador' WHERE id=?", (primeiro['id'],))
    conexao.commit()
    conexao.close()


def registrar_atividade(acao, detalhes=""):
    if "professor_id" not in session:
        return
    conexao = conectar()
    conexao.execute("INSERT INTO auditoria (professor_id, acao, detalhes, criado_em) VALUES (?, ?, ?, ?)",
                    (session["professor_id"], acao, detalhes[:500], datetime.now().isoformat(timespec="seconds")))
    conexao.commit(); conexao.close()


def administrador_obrigatorio(funcao):
    @wraps(funcao)
    def verificar(*args, **kwargs):
        if "professor_id" not in session:
            return redirect(url_for("login"))
        if session.get("perfil") != "administrador":
            flash("Esta área é exclusiva para administradores.")
            return redirect(url_for("painel"))
        return funcao(*args, **kwargs)
    return verificar


# =====================================================
# AUTENTICAÇÃO
# =====================================================

def login_obrigatorio(funcao):
    @wraps(funcao)
    def verificar_login(*args, **kwargs):
        if "professor_id" not in session:
            return redirect(url_for("login"))

        return funcao(*args, **kwargs)

    return verificar_login


# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================

def ano_selecionado():
    return session.get("ano", date.today().year)


def obter_anos_disponiveis():
    ano_hoje = date.today().year

    anos = set(range(ano_hoje - 1, ano_hoje + 6))
    anos.add(ano_selecionado())

    conexao = conectar()

    registros = conexao.execute("""
        SELECT DISTINCT ano
        FROM alunos
    """).fetchall()

    conexao.close()

    for registro in registros:
        anos.add(registro["ano"])

    return sorted(anos)


def calcular_situacao(aluno):
    notas = [
        aluno["nota1"],
        aluno["nota2"],
        aluno["nota3"],
        aluno["nota4"]
    ]

    notas_preenchidas = [
        nota for nota in notas if nota is not None
    ]

    if not notas_preenchidas:
        return None, "Em andamento"

    media = sum(notas_preenchidas) / len(notas_preenchidas)

    if len(notas_preenchidas) < 4:
        situacao = "Em andamento"
    elif media >= 6:
        situacao = "Aprovado"
    else:
        situacao = "Reprovado"

    return media, situacao


def calcular_estatisticas(alunos, frequencia=None):
    """
    Resume uma lista de alunos: total, aprovados, reprovados,
    em andamento, média geral e alunos com ausência alta.
    """
    medias = []
    aprovados = 0
    reprovados = 0
    andamento = 0
    em_risco = 0

    for aluno in alunos:

        media, situacao = calcular_situacao(aluno)

        if media is not None:
            medias.append(media)

        if situacao == "Aprovado":
            aprovados += 1
        elif situacao == "Reprovado":
            reprovados += 1
        else:
            andamento += 1

        if frequencia:
            freq = frequencia.get(aluno["id"])

            if (
                freq
                and freq["total"] > 0
                and freq["percentual"] > LIMITE_AUSENCIA
            ):
                em_risco += 1

    return {
        "total": len(alunos),
        "aprovados": aprovados,
        "reprovados": reprovados,
        "andamento": andamento,
        "em_risco": em_risco,
        "media": sum(medias) / len(medias) if medias else None
    }


def limpar_csv(texto):
    """Evita que o Excel trate um texto como fórmula."""
    texto = texto or ""

    if texto[:1] in ["=", "+", "-", "@"]:
        return "'" + texto

    return texto


def obter_alunos(materia, ano, serie=None):
    conexao = conectar()

    if serie:
        alunos = conexao.execute("""
            SELECT *
            FROM alunos
            WHERE materia = ? AND ano = ? AND serie = ?
            ORDER BY nome
        """, (materia, ano, serie)).fetchall()
    else:
        alunos = conexao.execute("""
            SELECT *
            FROM alunos
            WHERE materia = ? AND ano = ?
            ORDER BY serie, nome
        """, (materia, ano)).fetchall()

    conexao.close()
    return alunos


def obter_resumo_frequencia(alunos):
    """
    Retorna um dicionário com o resumo de frequência de cada aluno:
    {aluno_id: {"faltas": X, "total": Y, "percentual": Z}}
    """
    resumo = {}

    conexao = conectar()

    for aluno in alunos:

        linha = conexao.execute("""
            SELECT
                COUNT(*) AS total,
                COALESCE(
                    SUM(CASE WHEN presente = 0 THEN 1 ELSE 0 END),
                    0
                ) AS faltas
            FROM frequencia
            WHERE aluno_id = ?
        """, (aluno["id"],)).fetchone()

        total = linha["total"]
        faltas = linha["faltas"]

        if total > 0:
            percentual = (faltas / total) * 100
        else:
            percentual = 0

        resumo[aluno["id"]] = {
            "faltas": faltas,
            "total": total,
            "percentual": percentual
        }

    conexao.close()

    return resumo


def validar_data(data_texto):
    try:
        data_objeto = date.fromisoformat(data_texto)
        return data_objeto.isoformat()
    except (ValueError, TypeError):
        return None


def converter_nota(valor):
    valor = valor.strip().replace(",", ".")

    if valor == "":
        return None

    try:
        nota = float(valor)
    except ValueError:
        raise ValueError("Digite uma nota válida.")

    if nota < 0 or nota > 10:
        raise ValueError("As notas devem estar entre 0 e 10.")

    return nota


# =====================================================
# ESTILO E ESTRUTURA DAS PÁGINAS
# =====================================================




# =====================================================
# FUNÇÃO PARA RENDERIZAR AS PÁGINAS
# =====================================================

def renderizar_pagina(aba, **dados):
    return render_template(
        "index.html",
        mostrar_login=False,
        mostrar_cadastro_professor=False,
        aba=aba,
        materia=session.get("materia"),
        professor_nome=session.get("professor_nome"),
        perfil=session.get("perfil", "professor"),
        ano_letivo=ano_selecionado(),
        limite_ausencia=LIMITE_AUSENCIA,
        anos_disponiveis=obter_anos_disponiveis(),
        calcular_situacao=calcular_situacao,
        **dados
    )


# =====================================================
# LOGIN
# =====================================================

@app.route("/")
def inicio():
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        email = request.form.get("email", "").strip().lower()
        senha = request.form.get("senha", "")

        conexao = conectar()

        professor = conexao.execute("""
            SELECT *
            FROM professores
            WHERE email = ?
        """, (email,)).fetchone()

        conexao.close()

        if professor and check_password_hash(
            professor["senha"], senha
        ):

            session.clear()

            session["professor_id"] = professor["id"]
            session["professor_nome"] = professor["nome"]
            session["materia"] = professor["materia"]
            session["perfil"] = professor["perfil"]
            session["ano"] = date.today().year

            registrar_atividade("Login", "Acesso ao sistema")
            return redirect(url_for("painel"))

        flash("E-mail ou senha incorretos.")

    return render_template(
        "index.html",
        mostrar_login=True,
        mostrar_cadastro_professor=False
    )


# =====================================================
# CADASTRO DE PROFESSOR
# =====================================================

@app.route("/cadastro-professor", methods=["GET", "POST"])
def cadastro_professor():

    if request.method == "POST":

        nome = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip().lower()
        senha = request.form.get("senha", "")
        materia = request.form.get("materia", "")

        if not nome or not email or not senha:
            flash("Preencha todos os campos.")
        elif materia not in MATERIAS:
            flash("Selecione uma matéria válida.")
        else:

            senha_hash = generate_password_hash(senha)

            conexao = conectar()

            try:
                total_usuarios = conexao.execute("SELECT COUNT(*) FROM professores").fetchone()[0]
                perfil_inicial = "administrador" if total_usuarios == 0 else "professor"
                conexao.execute("""
                    INSERT INTO professores (nome, email, senha, materia, perfil)
                    VALUES (?, ?, ?, ?, ?)
                """, (nome, email, senha_hash, materia, perfil_inicial))

                conexao.commit()

                flash("Conta criada com sucesso! Faça seu login.")
                return redirect(url_for("login"))

            except sqlite3.IntegrityError:
                flash("Este e-mail já está cadastrado.")

            finally:
                conexao.close()

    return render_template(
        "index.html",
        mostrar_login=False,
        mostrar_cadastro_professor=True,
        materias=MATERIAS
    )


# =====================================================
# MUDAR O ANO LETIVO
# =====================================================

@app.route("/mudar-ano", methods=["POST"])
@login_obrigatorio
def mudar_ano():

    ano_texto = request.form.get("ano", "")
    voltar = request.form.get("voltar", "")
    serie = request.form.get("serie", "")

    try:
        ano = int(ano_texto)
    except ValueError:
        ano = None

    if ano is None or ano < 2000 or ano > 2100:
        flash("Ano inválido.")
    else:
        session["ano"] = ano

    parametros = {}

    if serie in SERIES_VALIDAS:
        parametros["serie"] = serie

    if voltar == "/listar":
        return redirect(url_for("listar_alunos", **parametros))

    if voltar == "/chamada":
        return redirect(url_for("chamada", **parametros))

    if voltar == "/ranking":
        return redirect(url_for("ranking"))

    if voltar == "/cadastrar":
        return redirect(url_for("cadastrar_aluno"))

    return redirect(url_for("painel"))


# =====================================================
# CADASTRAR ALUNO
# =====================================================

@app.route("/cadastrar", methods=["GET", "POST"])
@login_obrigatorio
def cadastrar_aluno():

    if request.method == "POST":

        nome = request.form.get("nome", "").strip()
        serie = request.form.get("serie", "")

        if not nome:
            flash("Digite o nome do aluno.")
        elif serie not in SERIES_VALIDAS:
            flash("Selecione uma série válida.")
        else:

            conexao = conectar()

            conexao.execute("""
                INSERT INTO alunos
                (nome, serie, materia, ano,
                 nota1, nota2, nota3, nota4)
                VALUES (?, ?, ?, ?, NULL, NULL, NULL, NULL)
            """, (
                nome,
                int(serie),
                session["materia"],
                ano_selecionado()
            ))

            conexao.commit()
            conexao.close()

            flash("Aluno cadastrado com sucesso!")
            return redirect(url_for("cadastrar_aluno"))

    return renderizar_pagina("cadastrar")


# =====================================================
# LISTAR ALUNOS
# =====================================================

@app.route("/listar")
@login_obrigatorio
def listar_alunos():

    serie = request.args.get("serie", "")

    alunos = []
    frequencia = {}
    estatisticas = None

    if serie in SERIES_VALIDAS:
        alunos = obter_alunos(
            session["materia"],
            ano_selecionado(),
            int(serie)
        )

        frequencia = obter_resumo_frequencia(alunos)
        estatisticas = calcular_estatisticas(alunos, frequencia)

    return renderizar_pagina(
        "listar",
        alunos=alunos,
        frequencia=frequencia,
        estatisticas=estatisticas,
        serie_selecionada=serie
    )


# =====================================================
# EDITAR NOTAS (envio normal do formulário)
# =====================================================

@app.route("/editar/<int:aluno_id>", methods=["POST"])
@login_obrigatorio
def editar_aluno(aluno_id):

    conexao = conectar()

    aluno = conexao.execute("""
        SELECT *
        FROM alunos
        WHERE id = ? AND materia = ?
    """, (
        aluno_id,
        session["materia"]
    )).fetchone()

    if not aluno:
        conexao.close()
        flash("Aluno não encontrado.")
        return redirect(url_for("listar_alunos"))

    try:

        notas = []

        for i in range(1, 5):

            valor = request.form.get(f"nota{i}", "")
            nota = converter_nota(valor)
            notas.append(nota)

        conexao.execute("""
            UPDATE alunos
            SET nota1 = ?,
                nota2 = ?,
                nota3 = ?,
                nota4 = ?
            WHERE id = ? AND materia = ?
        """, (
            notas[0],
            notas[1],
            notas[2],
            notas[3],
            aluno_id,
            session["materia"]
        ))

        conexao.commit()
        flash("Notas atualizadas com sucesso!")

    except ValueError as erro:
        flash(str(erro))

    except Exception:
        flash("Não foi possível atualizar as notas.")

    finally:
        conexao.close()

    return redirect(
        url_for("listar_alunos", serie=aluno["serie"])
    )


# =====================================================
# SALVAMENTO AUTOMÁTICO DAS NOTAS
# =====================================================

@app.route("/salvar-notas/<int:aluno_id>", methods=["POST"])
@login_obrigatorio
def salvar_notas_automatico(aluno_id):

    conexao = conectar()

    try:

        aluno = conexao.execute("""
            SELECT *
            FROM alunos
            WHERE id = ? AND materia = ?
        """, (
            aluno_id,
            session["materia"]
        )).fetchone()

        if not aluno:
            return jsonify(
                ok=False,
                erro="Aluno não encontrado."
            ), 404

        notas = []

        for i in range(1, 5):
            valor = request.form.get(f"nota{i}", "")
            notas.append(converter_nota(valor))

        conexao.execute("""
            UPDATE alunos
            SET nota1 = ?,
                nota2 = ?,
                nota3 = ?,
                nota4 = ?
            WHERE id = ? AND materia = ?
        """, (
            notas[0],
            notas[1],
            notas[2],
            notas[3],
            aluno_id,
            session["materia"]
        ))

        conexao.commit()

        atualizado = conexao.execute("""
            SELECT *
            FROM alunos
            WHERE id = ?
        """, (aluno_id,)).fetchone()

        media, situacao = calcular_situacao(atualizado)

        return jsonify(
            ok=True,
            media=None if media is None else round(media, 2),
            situacao=situacao
        )

    except ValueError as erro:
        return jsonify(ok=False, erro=str(erro)), 400

    except Exception:
        return jsonify(
            ok=False,
            erro="Não foi possível salvar."
        ), 500

    finally:
        conexao.close()


# =====================================================
# PAINEL INICIAL
# =====================================================

@app.route("/inicio")
@login_obrigatorio
def painel():

    ano = ano_selecionado()

    alunos = obter_alunos(session["materia"], ano)
    frequencia = obter_resumo_frequencia(alunos)

    geral = calcular_estatisticas(alunos, frequencia)

    por_serie = []

    for numero in range(1, 6):

        da_serie = [
            aluno for aluno in alunos if aluno["serie"] == numero
        ]

        if da_serie:
            dados = calcular_estatisticas(da_serie, frequencia)
            dados["serie"] = numero
            por_serie.append(dados)

    atencao = []

    for aluno in alunos:

        media, situacao = calcular_situacao(aluno)
        freq = frequencia[aluno["id"]]

        motivos = []

        if media is not None and media < 6:
            if situacao == "Reprovado":
                motivos.append("Média abaixo de 6")
            else:
                motivos.append("Média parcial abaixo de 6")

        if freq["total"] > 0 and freq["percentual"] > LIMITE_AUSENCIA:
            motivos.append(f"Ausência de {freq['percentual']:.0f}%")

        if motivos:
            atencao.append({
                "nome": aluno["nome"],
                "serie": aluno["serie"],
                "motivos": motivos
            })

    hoje = date.today()

    data_extenso = (
        f"{DIAS_SEMANA[hoje.weekday()]}, {hoje.day} de "
        f"{MESES[hoje.month - 1]} de {hoje.year}"
    )

    partes_nome = session.get("professor_nome", "").split()
    primeiro_nome = partes_nome[0] if partes_nome else "professor"

    return renderizar_pagina(
        "inicio",
        geral=geral,
        por_serie=por_serie,
        atencao=atencao[:12],
        total_atencao=len(atencao),
        data_extenso=data_extenso,
        primeiro_nome=primeiro_nome
    )


# =====================================================
# ATUALIZAR DADOS DO ALUNO
# =====================================================

@app.route("/atualizar-aluno/<int:aluno_id>", methods=["POST"])
@login_obrigatorio
def atualizar_aluno(aluno_id):

    nome = request.form.get("nome", "").strip()
    serie = request.form.get("serie", "")
    responsavel = request.form.get("responsavel", "").strip()[:120]
    telefone = request.form.get("telefone", "").strip()[:30]
    observacoes = request.form.get("observacoes", "").strip()[:500]

    conexao = conectar()

    aluno = conexao.execute("""
        SELECT *
        FROM alunos
        WHERE id = ? AND materia = ?
    """, (
        aluno_id,
        session["materia"]
    )).fetchone()

    if not aluno:
        conexao.close()
        flash("Aluno não encontrado.")
        return redirect(url_for("listar_alunos"))

    if not nome:
        conexao.close()
        flash("O nome do aluno não pode ficar vazio.")
        return redirect(
            url_for("listar_alunos", serie=aluno["serie"])
        )

    if serie not in SERIES_VALIDAS:
        conexao.close()
        flash("Selecione uma série válida.")
        return redirect(
            url_for("listar_alunos", serie=aluno["serie"])
        )

    conexao.execute("""
        UPDATE alunos
        SET nome = ?,
            serie = ?,
            responsavel = ?,
            telefone = ?,
            observacoes = ?
        WHERE id = ? AND materia = ?
    """, (
        nome,
        int(serie),
        responsavel,
        telefone,
        observacoes,
        aluno_id,
        session["materia"]
    ))

    conexao.commit()
    conexao.close()

    if int(serie) != aluno["serie"]:
        flash(f"{nome} agora está no {serie}º ano.")
    else:
        flash(f"Dados de {nome} atualizados com sucesso.")

    return redirect(url_for("listar_alunos", serie=serie))


# =====================================================
# EXPORTAR PLANILHA (CSV PARA O EXCEL)
# =====================================================

@app.route("/exportar")
@login_obrigatorio
def exportar():

    ano = ano_selecionado()

    alunos = obter_alunos(session["materia"], ano)
    frequencia = obter_resumo_frequencia(alunos)

    def formatar(numero):
        if numero is None:
            return ""
        return str(round(numero, 2)).replace(".", ",")

    saida = io.StringIO()
    escritor = csv.writer(saida, delimiter=";")

    escritor.writerow([
        "Nome", "Série", "1º Bimestre", "2º Bimestre",
        "3º Bimestre", "4º Bimestre", "Média", "Situação",
        "Faltas", "Ausência (%)", "Responsável", "Telefone",
        "Observações"
    ])

    for aluno in alunos:

        media, situacao = calcular_situacao(aluno)
        freq = frequencia[aluno["id"]]

        escritor.writerow([
            limpar_csv(aluno["nome"]),
            f"{aluno['serie']}º ano",
            formatar(aluno["nota1"]),
            formatar(aluno["nota2"]),
            formatar(aluno["nota3"]),
            formatar(aluno["nota4"]),
            formatar(media),
            situacao,
            freq["faltas"],
            formatar(freq["percentual"]),
            limpar_csv(aluno["responsavel"]),
            limpar_csv(aluno["telefone"]),
            limpar_csv(aluno["observacoes"])
        ])

    # O BOM no início faz o Excel abrir com acentos corretos.
    conteudo = "\ufeff" + saida.getvalue()

    return Response(
        conteudo,
        mimetype="text/csv; charset=utf-8",
        headers={
            "Content-Disposition":
                f"attachment; filename=alunos_{ano}.csv"
        }
    )


# =====================================================
# EXCLUIR ALUNO
# =====================================================

@app.route("/excluir/<int:aluno_id>", methods=["POST"])
@login_obrigatorio
def excluir_aluno(aluno_id):

    origem = request.form.get("origem", "listar")
    serie = request.form.get("serie", "")
    data_texto = validar_data(request.form.get("data", ""))

    conexao = conectar()

    aluno = conexao.execute("""
        SELECT *
        FROM alunos
        WHERE id = ? AND materia = ?
    """, (
        aluno_id,
        session["materia"]
    )).fetchone()

    if not aluno:
        conexao.close()
        flash("Aluno não encontrado.")
    else:
        # As faltas do aluno são apagadas junto (ON DELETE CASCADE).
        conexao.execute("""
            DELETE FROM alunos
            WHERE id = ? AND materia = ?
        """, (
            aluno_id,
            session["materia"]
        ))

        conexao.commit()
        conexao.close()

        flash(f"Aluno {aluno['nome']} excluído com sucesso.")

    parametros = {}

    if serie in SERIES_VALIDAS:
        parametros["serie"] = serie

    if origem == "chamada":

        if data_texto:
            parametros["data"] = data_texto

        return redirect(url_for("chamada", **parametros))

    return redirect(url_for("listar_alunos", **parametros))


# =====================================================
# LISTA DE CHAMADA
# =====================================================

@app.route("/chamada", methods=["GET", "POST"])
@login_obrigatorio
def chamada():

    if request.method == "POST":

        serie = request.form.get("serie", "")
        data_texto = validar_data(
            request.form.get("data", "")
        )

        if serie not in SERIES_VALIDAS or not data_texto:
            flash("Série ou data inválida.")
            return redirect(url_for("chamada"))

        alunos = obter_alunos(
            session["materia"],
            ano_selecionado(),
            int(serie)
        )

        ids_presentes = set(
            request.form.getlist("presente")
        )

        conexao = conectar()

        for aluno in alunos:

            presente = 1 if str(aluno["id"]) in ids_presentes else 0

            conexao.execute("""
                INSERT INTO frequencia
                (aluno_id, data, presente)
                VALUES (?, ?, ?)
                ON CONFLICT(aluno_id, data)
                DO UPDATE SET presente = excluded.presente
            """, (
                aluno["id"],
                data_texto,
                presente
            ))

        conexao.commit()
        conexao.close()

        flash("Chamada salva com sucesso!")

        return redirect(url_for(
            "chamada",
            serie=serie,
            data=data_texto
        ))

    serie = request.args.get("serie", "")
    data_texto = validar_data(
        request.args.get("data", date.today().isoformat())
    )

    if not data_texto:
        data_texto = date.today().isoformat()

    alunos = []
    presentes = []

    if serie in SERIES_VALIDAS:

        alunos = obter_alunos(
            session["materia"],
            ano_selecionado(),
            int(serie)
        )

        conexao = conectar()

        for aluno in alunos:

            registro = conexao.execute("""
                SELECT presente
                FROM frequencia
                WHERE aluno_id = ? AND data = ?
            """, (
                aluno["id"],
                data_texto
            )).fetchone()

            if registro is None:

                # Na primeira chamada, todos começam presentes.
                conexao.execute("""
                    INSERT INTO frequencia
                    (aluno_id, data, presente)
                    VALUES (?, ?, 1)
                """, (
                    aluno["id"],
                    data_texto
                ))

                presentes.append(aluno["id"])

            elif registro["presente"] == 1:

                presentes.append(aluno["id"])

        conexao.commit()
        conexao.close()

    return renderizar_pagina(
        "chamada",
        alunos=alunos,
        serie_selecionada=serie,
        data_selecionada=data_texto,
        presentes=presentes
    )


# =====================================================
# RANKING
# =====================================================

@app.route("/ranking")
@login_obrigatorio
def ranking():

    alunos = obter_alunos(
        session["materia"],
        ano_selecionado()
    )

    ranking_alunos = []

    for aluno in alunos:

        notas = [
            aluno["nota1"],
            aluno["nota2"],
            aluno["nota3"],
            aluno["nota4"]
        ]

        # O ranking considera apenas alunos com as quatro notas.
        if all(nota is not None for nota in notas):

            media, situacao = calcular_situacao(aluno)

            ranking_alunos.append({
                "nome": aluno["nome"],
                "serie": aluno["serie"],
                "media": media,
                "situacao": situacao
            })

    ranking_alunos.sort(
        key=lambda aluno: aluno["media"],
        reverse=True
    )

    return renderizar_pagina(
        "ranking",
        ranking=ranking_alunos
    )


# =====================================================
# CALENDÁRIO ESCOLAR
# =====================================================

@app.route("/calendario", methods=["GET", "POST"])
@login_obrigatorio
def calendario():
    conexao = conectar()
    if request.method == "POST":
        titulo = request.form.get("titulo", "").strip()
        data_evento = validar_data(request.form.get("data", ""))
        tipo = request.form.get("tipo", "Evento")
        descricao = request.form.get("descricao", "").strip()
        if not titulo or not data_evento:
            flash("Informe um título e uma data válida.")
        else:
            conexao.execute("INSERT INTO eventos (titulo, descricao, data, tipo, professor_id, ano) VALUES (?, ?, ?, ?, ?, ?)",
                (titulo[:120], descricao[:1000], data_evento, tipo[:40], session["professor_id"], ano_selecionado()))
            conexao.commit(); registrar_atividade("Novo evento", titulo)
            flash("Evento adicionado ao calendário.")
        conexao.close()
        return redirect(url_for("calendario"))
    eventos = conexao.execute("SELECT * FROM eventos WHERE ano = ? AND professor_id = ? ORDER BY data", (ano_selecionado(), session["professor_id"])).fetchall()
    conexao.close()
    return renderizar_pagina("calendario", eventos=eventos)

@app.route("/calendario/excluir/<int:evento_id>", methods=["POST"])
@login_obrigatorio
def excluir_evento(evento_id):
    conexao = conectar(); conexao.execute("DELETE FROM eventos WHERE id = ? AND professor_id = ?", (evento_id, session["professor_id"])); conexao.commit(); conexao.close()
    registrar_atividade("Exclusão de evento", str(evento_id)); flash("Evento removido.")
    return redirect(url_for("calendario"))

# =====================================================
# BOLETIM INDIVIDUAL EM PDF
# =====================================================

@app.route("/boletim/<int:aluno_id>")
@login_obrigatorio
def boletim_pdf(aluno_id):
    conexao = conectar()
    aluno = conexao.execute("SELECT * FROM alunos WHERE id = ? AND materia = ?", (aluno_id, session["materia"])).fetchone()
    conexao.close()
    if not aluno:
        flash("Aluno não encontrado."); return redirect(url_for("listar_alunos"))
    media, situacao = calcular_situacao(aluno)
    buffer = BytesIO(); pdf = canvas.Canvas(buffer, pagesize=A4); largura, altura = A4
    pdf.setTitle("Boletim escolar - " + aluno["nome"])
    pdf.setFont("Helvetica-Bold", 20); pdf.drawString(50, altura-65, "BOLETIM ESCOLAR")
    pdf.setFont("Helvetica", 11); pdf.drawString(50, altura-100, "Aluno: " + aluno["nome"])
    pdf.drawString(50, altura-120, f"Série: {aluno['serie']}º ano    Ano letivo: {aluno['ano']}")
    pdf.drawString(50, altura-140, "Disciplina: " + aluno["materia"])
    y=altura-190; pdf.setFont("Helvetica-Bold", 11); pdf.drawString(60,y,"Período"); pdf.drawString(250,y,"Nota")
    pdf.line(50,y-8,540,y-8); pdf.setFont("Helvetica", 11)
    for i in range(1,5):
        y-=28; nota=aluno[f"nota{i}"]; pdf.drawString(60,y,f"{i}º bimestre"); pdf.drawString(250,y,"—" if nota is None else f"{nota:.1f}")
    y-=45; pdf.setFont("Helvetica-Bold", 12); pdf.drawString(60,y,"Média: " + (f"{media:.2f}" if media is not None else "Em andamento"))
    y-=25; pdf.drawString(60,y,"Situação: " + situacao)
    pdf.setFont("Helvetica", 9); pdf.drawString(50,45,"Documento gerado pelo Sistema Escolar")
    pdf.save(); buffer.seek(0); registrar_atividade("Emissão de boletim", aluno["nome"])
    return Response(buffer.getvalue(), mimetype="application/pdf", headers={"Content-Disposition": f"attachment; filename=boletim_{aluno_id}.pdf"})

# =====================================================
# EXPORTAÇÃO XLSX
# =====================================================

@app.route("/exportar-excel")
@login_obrigatorio
def exportar_excel():
    alunos = obter_alunos(session["materia"], ano_selecionado()); wb=Workbook(); ws=wb.active; ws.title="Alunos"
    ws.append(["Aluno", "Série", "Disciplina", "1º Bimestre", "2º Bimestre", "3º Bimestre", "4º Bimestre", "Média", "Situação", "Responsável", "Telefone"])
    for a in alunos:
        media,situacao=calcular_situacao(a); ws.append([a["nome"], f"{a['serie']}º ano", a["materia"], a["nota1"],a["nota2"],a["nota3"],a["nota4"],media,situacao,a["responsavel"],a["telefone"]])
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        letra=col[0].column_letter; ws.column_dimensions[letra].width=min(max(max(len(str(c.value or "")) for c in col)+2,12),35)
    buf=BytesIO(); wb.save(buf); buf.seek(0); registrar_atividade("Exportação Excel", str(len(alunos))+" alunos")
    return Response(buf.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition":f"attachment; filename=relatorio_escolar_{ano_selecionado()}.xlsx"})

# =====================================================
# REGISTRO DE COMUNICAÇÃO COM RESPONSÁVEIS
# =====================================================

@app.route("/comunicacao", methods=["GET", "POST"])
@login_obrigatorio
def comunicacao():
    conexao=conectar()
    if request.method == "POST":
        aluno_id=request.form.get("aluno_id", type=int); assunto=request.form.get("assunto","").strip(); mensagem=request.form.get("mensagem","").strip(); canal=request.form.get("canal","Registro")
        aluno=conexao.execute("SELECT * FROM alunos WHERE id=? AND materia=?",(aluno_id,session["materia"])).fetchone()
        if not aluno or not assunto or not mensagem:
            flash("Confira o aluno, assunto e mensagem.")
        else:
            conexao.execute("INSERT INTO comunicacoes (aluno_id, assunto, mensagem, canal, data_envio, professor_id) VALUES (?, ?, ?, ?, ?, ?)",(aluno_id,assunto[:120],mensagem[:2000],canal,datetime.now().isoformat(timespec="minutes"),session["professor_id"]))
            conexao.commit(); registrar_atividade("Registro de comunicação", aluno["nome"]+": "+assunto); flash("Comunicação registrada no histórico. O sistema não envia mensagens automaticamente.")
        conexao.close(); return redirect(url_for("comunicacao"))
    alunos=conexao.execute("SELECT id,nome,serie,responsavel,telefone FROM alunos WHERE materia=? AND ano=? ORDER BY nome",(session["materia"],ano_selecionado())).fetchall()
    historico=conexao.execute("SELECT c.*, a.nome AS aluno_nome FROM comunicacoes c LEFT JOIN alunos a ON a.id=c.aluno_id WHERE c.professor_id=? ORDER BY c.id DESC LIMIT 100",(session["professor_id"],)).fetchall()
    conexao.close(); return renderizar_pagina("comunicacao", alunos=alunos, historico=historico)

# =====================================================
# ADMINISTRAÇÃO: AUDITORIA E BACKUP
# =====================================================

@app.route("/auditoria")
@login_obrigatorio
@administrador_obrigatorio
def auditoria():
    conexao=conectar(); registros=conexao.execute("SELECT a.*, p.nome FROM auditoria a LEFT JOIN professores p ON p.id=a.professor_id ORDER BY a.id DESC LIMIT 200").fetchall(); conexao.close()
    return renderizar_pagina("auditoria", registros=registros)

@app.route("/backup")
@login_obrigatorio
@administrador_obrigatorio
def backup():
    import tempfile
    origem = sqlite3.connect(BANCO)
    with tempfile.NamedTemporaryFile(suffix=".db") as tmp:
        copia = sqlite3.connect(tmp.name)
        origem.backup(copia)
        copia.close()
        origem.close()
        tmp.seek(0)
        conteudo = tmp.read()
    registrar_atividade("Backup do banco", "Download manual")
    return Response(
        conteudo,
        mimetype="application/octet-stream",
        headers={"Content-Disposition": f"attachment; filename=backup_escolar_{date.today().isoformat()}.db"}
    )


# =====================================================
# SAIR DO SISTEMA
# =====================================================

@app.route("/sair")
def sair():

    registrar_atividade("Logout", "Saída do sistema")
    session.clear()

    flash("Você saiu do sistema.")

    return redirect(url_for("login"))


# =====================================================
# INICIALIZAÇÃO
# =====================================================

if __name__ == "__main__":

    criar_banco()

    app.run(debug=True)
