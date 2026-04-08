import os
import sqlite3
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify

app = Flask(__name__)
app.secret_key = os.urandom(24)
DB_PATH = os.path.join(os.path.dirname(__file__), "controle.db")
DURACAO_REUNIAO = 90  # minutos


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS salas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE
        );

        CREATE TABLE IF NOT EXISTS notebooks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            localizacao TEXT DEFAULT 'Desconhecida',
            responsavel TEXT DEFAULT '',
            atualizado_em TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS reservas_sala (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sala_id INTEGER NOT NULL,
            responsavel TEXT NOT NULL,
            data TEXT NOT NULL DEFAULT (date('now','localtime')),
            hora_inicio TEXT NOT NULL,
            hora_fim TEXT NOT NULL,
            motivo TEXT DEFAULT '',
            link_reuniao TEXT DEFAULT '',
            FOREIGN KEY (sala_id) REFERENCES salas(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS historico_notebook (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            notebook_id INTEGER NOT NULL,
            localizacao_anterior TEXT,
            localizacao_nova TEXT,
            responsavel TEXT,
            movido_em TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (notebook_id) REFERENCES notebooks(id) ON DELETE CASCADE
        );
    """)
    conn.commit()
    conn.close()


def sala_status(conn, sala_id):
    """Retorna a reserva ativa AGORA para a sala (data de hoje + hora atual)."""
    hoje = date.today().isoformat()
    agora = datetime.now().strftime("%H:%M")
    return conn.execute("""
        SELECT * FROM reservas_sala
        WHERE sala_id = ? AND data = ? AND hora_inicio <= ? AND hora_fim > ?
    """, (sala_id, hoje, agora, agora)).fetchone()


# ──────────────── PAGINA INICIAL ────────────────

@app.route("/")
def index():
    conn = get_db()
    salas = conn.execute("SELECT * FROM salas ORDER BY nome").fetchall()
    notebooks = conn.execute("SELECT * FROM notebooks ORDER BY nome").fetchall()
    hoje = date.today().isoformat()
    agora = datetime.now().strftime("%H:%M")
    limite = (datetime.now() + timedelta(minutes=DURACAO_REUNIAO)).strftime("%H:%M")

    salas_info = []
    for s in salas:
        reserva = sala_status(conn, s["id"])
        proxima_res = None
        if reserva is None:
            proxima_res = conn.execute("""
                SELECT * FROM reservas_sala
                WHERE sala_id = ? AND data = ? AND hora_inicio > ? AND hora_inicio <= ?
                ORDER BY hora_inicio LIMIT 1
            """, (s["id"], hoje, agora, limite)).fetchone()
        salas_info.append({
            "id": s["id"],
            "nome": s["nome"],
            "ocupada": reserva is not None,
            "responsavel": reserva["responsavel"] if reserva else None,
            "motivo": reserva["motivo"] if reserva else None,
            "ate": reserva["hora_fim"] if reserva else None,
            "link_reuniao": reserva["link_reuniao"] if reserva else None,
            "proxima": {
                "responsavel": proxima_res["responsavel"],
                "hora_inicio": proxima_res["hora_inicio"],
                "motivo": proxima_res["motivo"],
            } if proxima_res else None,
        })

    proximas = conn.execute("""
        SELECT rs.*, s.nome as sala_nome
        FROM reservas_sala rs
        JOIN salas s ON s.id = rs.sala_id
        WHERE (rs.data = ? AND rs.hora_fim > ?) OR rs.data > ?
        ORDER BY rs.data, rs.hora_inicio
        LIMIT 15
    """, (hoje, agora, hoje)).fetchall()

    nomes_salas = {s["nome"] for s in salas}
    notebooks_livres = [nb for nb in notebooks if nb["localizacao"] not in nomes_salas]
    notebooks_em_sala = [nb for nb in notebooks if nb["localizacao"] in nomes_salas]

    conn.close()
    return render_template("index.html",
                           salas=salas_info,
                           notebooks=notebooks,
                           notebooks_livres=notebooks_livres,
                           notebooks_em_sala=notebooks_em_sala,
                           proximas=proximas,
                           agora=agora,
                           hoje=hoje)


# ──────────────── RESERVAR / LIBERAR SALA ────────────────

@app.route("/reservar/<int:sala_id>", methods=["POST"])
def reservar_sala(sala_id):
    responsavel = request.form.get("responsavel", "").strip()
    motivo = request.form.get("motivo", "").strip()
    link_reuniao = request.form.get("link_reuniao", "").strip()
    hora_inicio_form = request.form.get("hora_inicio", "").strip()
    data_form = request.form.get("data", "").strip()
    dia_inteiro = request.form.get("dia_inteiro") == "1"

    if not responsavel:
        flash("Informe seu nome.", "danger")
        return redirect(url_for("index"))

    data_reserva = data_form if data_form else date.today().isoformat()

    if dia_inteiro:
        inicio = "00:00"
        fim = "23:59"
    elif hora_inicio_form:
        try:
            inicio_dt = datetime.strptime(hora_inicio_form, "%H:%M")
            inicio = hora_inicio_form
            fim = (inicio_dt + timedelta(minutes=DURACAO_REUNIAO)).strftime("%H:%M")
        except ValueError:
            flash("Hora inválida.", "danger")
            return redirect(url_for("index"))
    else:
        agora = datetime.now()
        inicio = agora.strftime("%H:%M")
        fim = (agora + timedelta(minutes=DURACAO_REUNIAO)).strftime("%H:%M")

    conn = get_db()
    conflito = conn.execute("""
        SELECT * FROM reservas_sala
        WHERE sala_id = ? AND data = ? AND hora_inicio < ? AND hora_fim > ?
    """, (sala_id, data_reserva, fim, inicio)).fetchone()

    if conflito:
        flash(f"Sala já está ocupada nesse horário em {data_reserva}!", "danger")
    else:
        conn.execute("""
            INSERT INTO reservas_sala (sala_id, responsavel, data, hora_inicio, hora_fim, motivo, link_reuniao)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (sala_id, responsavel, data_reserva, inicio, fim, motivo, link_reuniao))
        conn.commit()
        try:
            data_fmt = datetime.strptime(data_reserva, "%Y-%m-%d").strftime("%d/%m")
        except ValueError:
            data_fmt = data_reserva
        flash(f"Sala reservada para {data_fmt} das {inicio} às {fim}.", "success")
    conn.close()
    return redirect(url_for("index"))


@app.route("/liberar/<int:sala_id>", methods=["POST"])
def liberar_sala(sala_id):
    conn = get_db()
    hoje = date.today().isoformat()
    agora = datetime.now().strftime("%H:%M")
    conn.execute("""
        DELETE FROM reservas_sala
        WHERE sala_id = ? AND data = ? AND hora_inicio <= ? AND hora_fim > ?
    """, (sala_id, hoje, agora, agora))
    conn.commit()
    conn.close()
    flash("Sala liberada.", "info")
    return redirect(url_for("index"))


@app.route("/reserva/<int:reserva_id>/cancelar", methods=["POST"])
def cancelar_reserva(reserva_id):
    conn = get_db()
    conn.execute("DELETE FROM reservas_sala WHERE id = ?", (reserva_id,))
    conn.commit()
    conn.close()
    flash("Reserva cancelada.", "info")
    return redirect(request.referrer or url_for("index"))


@app.route("/reserva/<int:reserva_id>/editar", methods=["POST"])
def editar_reserva(reserva_id):
    responsavel = request.form.get("responsavel", "").strip()
    motivo = request.form.get("motivo", "").strip()
    link_reuniao = request.form.get("link_reuniao", "").strip()
    hora_inicio_form = request.form.get("hora_inicio", "").strip()
    data_form = request.form.get("data", "").strip()
    dia_inteiro = request.form.get("dia_inteiro") == "1"

    if not responsavel:
        flash("Informe seu nome.", "danger")
        return redirect(request.referrer or url_for("index"))

    conn = get_db()
    reserva = conn.execute("SELECT * FROM reservas_sala WHERE id = ?", (reserva_id,)).fetchone()
    if not reserva:
        flash("Reserva não encontrada.", "danger")
        conn.close()
        return redirect(url_for("index"))

    data_reserva = data_form if data_form else reserva["data"]

    if dia_inteiro:
        inicio = "00:00"
        fim = "23:59"
    elif hora_inicio_form:
        try:
            inicio_dt = datetime.strptime(hora_inicio_form, "%H:%M")
            inicio = hora_inicio_form
            fim = (inicio_dt + timedelta(minutes=DURACAO_REUNIAO)).strftime("%H:%M")
        except ValueError:
            flash("Hora inválida.", "danger")
            conn.close()
            return redirect(request.referrer or url_for("index"))
    else:
        inicio = reserva["hora_inicio"]
        fim = reserva["hora_fim"]

    # Conflito (excluindo a própria reserva)
    conflito = conn.execute("""
        SELECT * FROM reservas_sala
        WHERE sala_id = ? AND data = ? AND hora_inicio < ? AND hora_fim > ? AND id != ?
    """, (reserva["sala_id"], data_reserva, fim, inicio, reserva_id)).fetchone()

    if conflito:
        flash("Conflito com outra reserva nesse horário!", "danger")
    else:
        conn.execute("""
            UPDATE reservas_sala
            SET responsavel = ?, data = ?, hora_inicio = ?, hora_fim = ?, motivo = ?, link_reuniao = ?
            WHERE id = ?
        """, (responsavel, data_reserva, inicio, fim, motivo, link_reuniao, reserva_id))
        conn.commit()
        flash("Reserva atualizada.", "success")
    conn.close()
    return redirect(request.referrer or url_for("index"))


# ──────────────── AGENDA ────────────────

@app.route("/agenda")
def agenda():
    conn = get_db()
    salas = conn.execute("SELECT * FROM salas ORDER BY nome").fetchall()

    data_inicio_str = request.args.get("semana", date.today().isoformat())
    try:
        data_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d").date()
    except ValueError:
        data_inicio = date.today()

    # Começar na segunda-feira da semana
    data_inicio = data_inicio - timedelta(days=data_inicio.weekday())
    dias = [data_inicio + timedelta(days=i) for i in range(7)]
    data_fim = dias[-1]

    semana_anterior = (data_inicio - timedelta(days=7)).isoformat()
    semana_proxima = (data_inicio + timedelta(days=7)).isoformat()

    reservas = conn.execute("""
        SELECT rs.*, s.nome as sala_nome
        FROM reservas_sala rs
        JOIN salas s ON s.id = rs.sala_id
        WHERE rs.data >= ? AND rs.data <= ?
        ORDER BY rs.data, rs.hora_inicio
    """, (data_inicio.isoformat(), data_fim.isoformat())).fetchall()
    conn.close()

    # Organizar: agenda_data[data_str][sala_nome] = [lista de reservas]
    agenda_data = {}
    for d in dias:
        d_str = d.isoformat()
        agenda_data[d_str] = {}
        for s in salas:
            agenda_data[d_str][s["nome"]] = []

    for r in reservas:
        d_str = r["data"]
        sala_nome = r["sala_nome"]
        if d_str in agenda_data and sala_nome in agenda_data[d_str]:
            agenda_data[d_str][sala_nome].append(dict(r))

    DIAS_SEMANA = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

    return render_template("agenda.html",
                           salas=salas,
                           dias=dias,
                           dias_semana=DIAS_SEMANA,
                           agenda_data=agenda_data,
                           semana_anterior=semana_anterior,
                           semana_proxima=semana_proxima,
                           hoje=date.today().isoformat())


# ──────────────── GERENCIAR SALAS ────────────────

@app.route("/salas")
def gerenciar_salas():
    conn = get_db()
    salas = conn.execute("SELECT * FROM salas ORDER BY nome").fetchall()
    conn.close()
    return render_template("salas.html", salas=salas)


@app.route("/salas/nova", methods=["POST"])
def nova_sala():
    nome = request.form.get("nome", "").strip()
    if not nome:
        flash("Nome da sala é obrigatório.", "danger")
        return redirect(url_for("gerenciar_salas"))
    conn = get_db()
    try:
        conn.execute("INSERT INTO salas (nome) VALUES (?)", (nome,))
        conn.commit()
        flash(f"Sala '{nome}' cadastrada.", "success")
    except sqlite3.IntegrityError:
        flash(f"Sala '{nome}' já existe.", "warning")
    conn.close()
    return redirect(url_for("gerenciar_salas"))


@app.route("/salas/<int:sala_id>/excluir", methods=["POST"])
def excluir_sala(sala_id):
    conn = get_db()
    conn.execute("DELETE FROM salas WHERE id = ?", (sala_id,))
    conn.commit()
    conn.close()
    flash("Sala excluída.", "info")
    return redirect(url_for("gerenciar_salas"))


# ──────────────── NOTEBOOKS ────────────────

@app.route("/notebooks")
def gerenciar_notebooks():
    conn = get_db()
    notebooks = conn.execute("SELECT * FROM notebooks ORDER BY nome").fetchall()
    conn.close()
    return render_template("notebooks.html", notebooks=notebooks)


@app.route("/notebooks/novo", methods=["POST"])
def novo_notebook():
    nome = request.form.get("nome", "").strip()
    localizacao = request.form.get("localizacao", "").strip()
    if not nome:
        flash("Nome do notebook é obrigatório.", "danger")
        return redirect(url_for("gerenciar_notebooks"))
    conn = get_db()
    try:
        conn.execute("""
            INSERT INTO notebooks (nome, localizacao, atualizado_em)
            VALUES (?, ?, ?)
        """, (nome, localizacao, datetime.now().strftime("%d/%m/%Y %H:%M")))
        conn.commit()
        flash(f"Notebook '{nome}' cadastrado.", "success")
    except sqlite3.IntegrityError:
        flash(f"Notebook '{nome}' já existe.", "warning")
    conn.close()
    return redirect(url_for("gerenciar_notebooks"))


@app.route("/notebooks/<int:nb_id>/mover", methods=["POST"])
def mover_notebook(nb_id):
    nova_loc = request.form.get("localizacao", "").strip()
    responsavel = request.form.get("responsavel", "").strip()
    if not nova_loc:
        flash("Informe a nova localização.", "danger")
        return redirect(url_for("gerenciar_notebooks"))

    conn = get_db()
    nb = conn.execute("SELECT * FROM notebooks WHERE id = ?", (nb_id,)).fetchone()
    if nb:
        conn.execute("""
            INSERT INTO historico_notebook (notebook_id, localizacao_anterior, localizacao_nova, responsavel)
            VALUES (?, ?, ?, ?)
        """, (nb_id, nb["localizacao"], nova_loc, responsavel))
        conn.execute("""
            UPDATE notebooks SET localizacao = ?, responsavel = ?, atualizado_em = ? WHERE id = ?
        """, (nova_loc, responsavel, datetime.now().strftime("%d/%m/%Y %H:%M"), nb_id))
        conn.commit()
        flash(f"Notebook movido para '{nova_loc}'.", "success")
    conn.close()
    return redirect(url_for("gerenciar_notebooks"))


@app.route("/notebooks/<int:nb_id>/historico")
def historico_notebook(nb_id):
    conn = get_db()
    nb = conn.execute("SELECT * FROM notebooks WHERE id = ?", (nb_id,)).fetchone()
    historico = conn.execute("""
        SELECT * FROM historico_notebook
        WHERE notebook_id = ?
        ORDER BY movido_em DESC
    """, (nb_id,)).fetchall()
    conn.close()
    return render_template("historico_notebook.html", notebook=nb, historico=historico)


@app.route("/notebooks/<int:nb_id>/excluir", methods=["POST"])
def excluir_notebook(nb_id):
    conn = get_db()
    conn.execute("DELETE FROM notebooks WHERE id = ?", (nb_id,))
    conn.commit()
    conn.close()
    flash("Notebook excluído.", "info")
    return redirect(url_for("gerenciar_notebooks"))


# ──────────────── MOVER NOTEBOOK VIA DRAG ────────────────

@app.route("/notebooks/<int:nb_id>/mover-para-sala", methods=["POST"])
def mover_notebook_para_sala(nb_id):
    data = request.get_json()
    sala_nome = data.get("sala_nome", "").strip() if data else ""
    if not sala_nome:
        return jsonify({"ok": False, "msg": "Sala não informada."}), 400

    conn = get_db()
    nb = conn.execute("SELECT * FROM notebooks WHERE id = ?", (nb_id,)).fetchone()
    if nb:
        conn.execute("""
            INSERT INTO historico_notebook (notebook_id, localizacao_anterior, localizacao_nova, responsavel)
            VALUES (?, ?, ?, ?)
        """, (nb_id, nb["localizacao"], sala_nome, ""))
        conn.execute("""
            UPDATE notebooks SET localizacao = ?, responsavel = '', atualizado_em = ? WHERE id = ?
        """, (sala_nome, datetime.now().strftime("%d/%m/%Y %H:%M"), nb_id))
        conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ──────────────── TIRAR NOTEBOOK DA SALA ────────────────

@app.route("/notebooks/<int:nb_id>/tirar-da-sala", methods=["POST"])
def tirar_notebook_da_sala(nb_id):
    conn = get_db()
    nb = conn.execute("SELECT * FROM notebooks WHERE id = ?", (nb_id,)).fetchone()
    if nb:
        conn.execute("""
            INSERT INTO historico_notebook (notebook_id, localizacao_anterior, localizacao_nova, responsavel)
            VALUES (?, ?, ?, ?)
        """, (nb_id, nb["localizacao"], "Livre", ""))
        conn.execute("""
            UPDATE notebooks SET localizacao = 'Livre', responsavel = '', atualizado_em = ? WHERE id = ?
        """, (datetime.now().strftime("%d/%m/%Y %H:%M"), nb_id))
        conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ──────────────── LIMPAR EXPIRADAS ────────────────

@app.route("/limpar")
def limpar_expiradas():
    conn = get_db()
    hoje = date.today().isoformat()
    agora = datetime.now().strftime("%H:%M")
    conn.execute("""
        DELETE FROM reservas_sala
        WHERE data < ? OR (data = ? AND hora_fim <= ?)
    """, (hoje, hoje, agora))
    conn.commit()
    conn.close()
    return redirect(url_for("index"))


# ──────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    print("=" * 50)
    print("  Controle de Salas e Notebooks")
    print("  Acesse: http://localhost:333")
    print("=" * 50)
    app.run(host="0.0.0.0", port=333, debug=True)
