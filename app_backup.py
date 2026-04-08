import os
import sqlite3
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify

app = Flask(__name__)
app.secret_key = os.urandom(24)
DB_PATH = os.path.join(os.path.dirname(__file__), "controle.db")


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
            nome TEXT NOT NULL UNIQUE,
            capacidade INTEGER DEFAULT 0,
            descricao TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS notebooks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            descricao TEXT DEFAULT '',
            localizacao TEXT DEFAULT 'Desconhecida',
            responsavel TEXT DEFAULT '',
            atualizado_em TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS reservas_sala (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sala_id INTEGER NOT NULL,
            responsavel TEXT NOT NULL,
            data TEXT NOT NULL,
            hora_inicio TEXT NOT NULL,
            hora_fim TEXT NOT NULL,
            motivo TEXT DEFAULT '',
            criado_em TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (sala_id) REFERENCES salas(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS reservas_notebook (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            notebook_id INTEGER NOT NULL,
            responsavel TEXT NOT NULL,
            data TEXT NOT NULL,
            hora_inicio TEXT NOT NULL,
            hora_fim TEXT NOT NULL,
            motivo TEXT DEFAULT '',
            criado_em TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (notebook_id) REFERENCES notebooks(id) ON DELETE CASCADE
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


# ──────────────── PAGINA INICIAL ────────────────

@app.route("/")
def index():
    conn = get_db()
    hoje = date.today().isoformat()
    agora = datetime.now().strftime("%H:%M")

    salas = conn.execute("SELECT * FROM salas ORDER BY nome").fetchall()
    notebooks = conn.execute("SELECT * FROM notebooks ORDER BY nome").fetchall()

    reservas_hoje = conn.execute("""
        SELECT rs.*, s.nome as sala_nome
        FROM reservas_sala rs
        JOIN salas s ON s.id = rs.sala_id
        WHERE rs.data = ?
        ORDER BY rs.hora_inicio
    """, (hoje,)).fetchall()

    reservas_nb_hoje = conn.execute("""
        SELECT rn.*, n.nome as notebook_nome
        FROM reservas_notebook rn
        JOIN notebooks n ON n.id = rn.notebook_id
        WHERE rn.data = ?
        ORDER BY rn.hora_inicio
    """, (hoje,)).fetchall()

    conn.close()
    return render_template("index.html",
                           salas=salas,
                           notebooks=notebooks,
                           reservas_hoje=reservas_hoje,
                           reservas_nb_hoje=reservas_nb_hoje,
                           hoje=hoje,
                           agora=agora)


# ──────────────── SALAS ────────────────

@app.route("/salas")
def listar_salas():
    conn = get_db()
    salas = conn.execute("SELECT * FROM salas ORDER BY nome").fetchall()
    conn.close()
    return render_template("salas.html", salas=salas)


@app.route("/salas/nova", methods=["POST"])
def nova_sala():
    nome = request.form.get("nome", "").strip()
    capacidade = request.form.get("capacidade", 0, type=int)
    descricao = request.form.get("descricao", "").strip()
    if not nome:
        flash("Nome da sala é obrigatório.", "danger")
        return redirect(url_for("listar_salas"))
    conn = get_db()
    try:
        conn.execute("INSERT INTO salas (nome, capacidade, descricao) VALUES (?, ?, ?)",
                      (nome, capacidade, descricao))
        conn.commit()
        flash(f"Sala '{nome}' cadastrada.", "success")
    except sqlite3.IntegrityError:
        flash(f"Sala '{nome}' já existe.", "warning")
    conn.close()
    return redirect(url_for("listar_salas"))


@app.route("/salas/<int:sala_id>/excluir", methods=["POST"])
def excluir_sala(sala_id):
    conn = get_db()
    conn.execute("DELETE FROM salas WHERE id = ?", (sala_id,))
    conn.commit()
    conn.close()
    flash("Sala excluída.", "info")
    return redirect(url_for("listar_salas"))


# ──────────────── RESERVAS DE SALA ────────────────

@app.route("/reservas/sala")
def reservas_sala():
    conn = get_db()
    salas = conn.execute("SELECT * FROM salas ORDER BY nome").fetchall()
    data_filtro = request.args.get("data", date.today().isoformat())
    reservas = conn.execute("""
        SELECT rs.*, s.nome as sala_nome
        FROM reservas_sala rs
        JOIN salas s ON s.id = rs.sala_id
        WHERE rs.data = ?
        ORDER BY rs.hora_inicio
    """, (data_filtro,)).fetchall()
    conn.close()
    return render_template("reservas_sala.html",
                           salas=salas, reservas=reservas, data_filtro=data_filtro)


@app.route("/reservas/sala/nova", methods=["POST"])
def nova_reserva_sala():
    sala_id = request.form.get("sala_id", type=int)
    responsavel = request.form.get("responsavel", "").strip()
    data = request.form.get("data", "").strip()
    hora_inicio = request.form.get("hora_inicio", "").strip()
    hora_fim = request.form.get("hora_fim", "").strip()
    motivo = request.form.get("motivo", "").strip()

    if not all([sala_id, responsavel, data, hora_inicio, hora_fim]):
        flash("Preencha todos os campos obrigatórios.", "danger")
        return redirect(url_for("reservas_sala"))

    if hora_fim <= hora_inicio:
        flash("Hora fim deve ser maior que hora início.", "danger")
        return redirect(url_for("reservas_sala", data=data))

    conn = get_db()
    conflito = conn.execute("""
        SELECT * FROM reservas_sala
        WHERE sala_id = ? AND data = ?
        AND hora_inicio < ? AND hora_fim > ?
    """, (sala_id, data, hora_fim, hora_inicio)).fetchone()

    if conflito:
        flash("Conflito de horário! Essa sala já está reservada nesse período.", "danger")
    else:
        conn.execute("""
            INSERT INTO reservas_sala (sala_id, responsavel, data, hora_inicio, hora_fim, motivo)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (sala_id, responsavel, data, hora_inicio, hora_fim, motivo))
        conn.commit()
        flash("Reserva de sala criada.", "success")
    conn.close()
    return redirect(url_for("reservas_sala", data=data))


@app.route("/reservas/sala/<int:reserva_id>/excluir", methods=["POST"])
def excluir_reserva_sala(reserva_id):
    conn = get_db()
    reserva = conn.execute("SELECT data FROM reservas_sala WHERE id = ?", (reserva_id,)).fetchone()
    data = reserva["data"] if reserva else date.today().isoformat()
    conn.execute("DELETE FROM reservas_sala WHERE id = ?", (reserva_id,))
    conn.commit()
    conn.close()
    flash("Reserva excluída.", "info")
    return redirect(url_for("reservas_sala", data=data))


# ──────────────── NOTEBOOKS ────────────────

@app.route("/notebooks")
def listar_notebooks():
    conn = get_db()
    notebooks = conn.execute("SELECT * FROM notebooks ORDER BY nome").fetchall()
    conn.close()
    return render_template("notebooks.html", notebooks=notebooks)


@app.route("/notebooks/novo", methods=["POST"])
def novo_notebook():
    nome = request.form.get("nome", "").strip()
    descricao = request.form.get("descricao", "").strip()
    localizacao = request.form.get("localizacao", "").strip()
    if not nome:
        flash("Nome do notebook é obrigatório.", "danger")
        return redirect(url_for("listar_notebooks"))
    conn = get_db()
    try:
        conn.execute("""
            INSERT INTO notebooks (nome, descricao, localizacao, atualizado_em)
            VALUES (?, ?, ?, ?)
        """, (nome, descricao, localizacao, datetime.now().strftime("%d/%m/%Y %H:%M")))
        conn.commit()
        flash(f"Notebook '{nome}' cadastrado.", "success")
    except sqlite3.IntegrityError:
        flash(f"Notebook '{nome}' já existe.", "warning")
    conn.close()
    return redirect(url_for("listar_notebooks"))


@app.route("/notebooks/<int:nb_id>/mover", methods=["POST"])
def mover_notebook(nb_id):
    nova_loc = request.form.get("localizacao", "").strip()
    responsavel = request.form.get("responsavel", "").strip()
    if not nova_loc:
        flash("Informe a nova localização.", "danger")
        return redirect(url_for("listar_notebooks"))

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
        flash(f"Notebook '{nb['nome']}' movido para '{nova_loc}'.", "success")
    conn.close()
    return redirect(url_for("listar_notebooks"))


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
    return redirect(url_for("listar_notebooks"))


# ──────────────── RESERVAS DE NOTEBOOK ────────────────

@app.route("/reservas/notebook")
def reservas_notebook():
    conn = get_db()
    notebooks = conn.execute("SELECT * FROM notebooks ORDER BY nome").fetchall()
    data_filtro = request.args.get("data", date.today().isoformat())
    reservas = conn.execute("""
        SELECT rn.*, n.nome as notebook_nome
        FROM reservas_notebook rn
        JOIN notebooks n ON n.id = rn.notebook_id
        WHERE rn.data = ?
        ORDER BY rn.hora_inicio
    """, (data_filtro,)).fetchall()
    conn.close()
    return render_template("reservas_notebook.html",
                           notebooks=notebooks, reservas=reservas, data_filtro=data_filtro)


@app.route("/reservas/notebook/nova", methods=["POST"])
def nova_reserva_notebook():
    notebook_id = request.form.get("notebook_id", type=int)
    responsavel = request.form.get("responsavel", "").strip()
    data = request.form.get("data", "").strip()
    hora_inicio = request.form.get("hora_inicio", "").strip()
    hora_fim = request.form.get("hora_fim", "").strip()
    motivo = request.form.get("motivo", "").strip()

    if not all([notebook_id, responsavel, data, hora_inicio, hora_fim]):
        flash("Preencha todos os campos obrigatórios.", "danger")
        return redirect(url_for("reservas_notebook"))

    if hora_fim <= hora_inicio:
        flash("Hora fim deve ser maior que hora início.", "danger")
        return redirect(url_for("reservas_notebook", data=data))

    conn = get_db()
    conflito = conn.execute("""
        SELECT * FROM reservas_notebook
        WHERE notebook_id = ? AND data = ?
        AND hora_inicio < ? AND hora_fim > ?
    """, (notebook_id, data, hora_fim, hora_inicio)).fetchone()

    if conflito:
        flash("Conflito! Esse notebook já está reservado nesse horário.", "danger")
    else:
        conn.execute("""
            INSERT INTO reservas_notebook (notebook_id, responsavel, data, hora_inicio, hora_fim, motivo)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (notebook_id, responsavel, data, hora_inicio, hora_fim, motivo))
        conn.commit()
        flash("Reserva de notebook criada.", "success")
    conn.close()
    return redirect(url_for("reservas_notebook", data=data))


@app.route("/reservas/notebook/<int:reserva_id>/excluir", methods=["POST"])
def excluir_reserva_notebook(reserva_id):
    conn = get_db()
    reserva = conn.execute("SELECT data FROM reservas_notebook WHERE id = ?", (reserva_id,)).fetchone()
    data = reserva["data"] if reserva else date.today().isoformat()
    conn.execute("DELETE FROM reservas_notebook WHERE id = ?", (reserva_id,))
    conn.commit()
    conn.close()
    flash("Reserva excluída.", "info")
    return redirect(url_for("reservas_notebook", data=data))


# ──────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    print("=" * 50)
    print("  Sistema de Controle de Salas e Notebooks")
    print("  Acesse: http://localhost:333")
    print("=" * 50)
    app.run(host="0.0.0.0", port=333, debug=True)
