from flask import Flask, render_template, request, redirect, url_for, make_response
import sqlite3
from collections import defaultdict
import win32print
import win32ui
import tkinter as tk
from tkinter import simpledialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app = Flask(__name__)

conn = sqlite3.connect('gestao_cliches.db')
cursor = conn.cursor()

#criaçao de tabelas do banco de dados
cursor.execute('''
               CREATE TABLE IF NOT EXISTS clientes(
               id_cliente INTEGER PRIMARY KEY AUTOINCREMENT,
               codigo TEXT,
               nome VARCHAR,
               ncaixa TEXT,
               cnpj TEXT UNIQUE
               )
               ''')
conn.commit()

cursor.execute('''
               CREATE TABLE IF NOT EXISTS produtos(
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               codigo VARCHAR,
               codigo_cliente,
               descricao VARCHAR,
               os VARCHAR,
               clicheria VARCHAR,
               qtde NUMBER,
               data VARCHAR
               )
               ''')
conn.commit()
cursor.execute('''
               CREATE TABLE IF NOT EXISTS estoque(
               codigoest NUMBER,
               item,
               obs VARCHAR,
               cod_local
               )
               ''')
conn.commit()
cursor.execute('''
               CREATE TABLE IF NOT EXISTS movestoque(
               item,
               data VARCHAR,
               cod_local
               )
               ''')
conn.commit()
conn.close()


@app.route('/')
def index():
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()
    cursor.execute('Select COUNT(nome) from clientes')
    total_clientes = cursor.fetchone()[0]
    conn.close()
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()
    cursor.execute('Select COUNT(ncaixa) from clientes')
    totaldecaixa = cursor.fetchone()[0]
    conn.close()
    return render_template('index.html', total_clientes=total_clientes, totaldecaixa=totaldecaixa)

@app.route('/novo_cliente', methods=['GET', 'POST'])
def novo_cliente():
    if request.method == 'POST':
        codigo = request.form['codigo']
        nome = request.form['nome']
        ncaixa = request.form['ncaixa']
        cnpj = request.form['cnpj']
        conn = sqlite3.connect ('gestao_cliches.db')
        cursor = conn.cursor()
        cursor.execute('''
                   INSERT INTO clientes (codigo, nome, ncaixa, cnpj)
                   values(?, ?, ?, ?)
                   ''', (codigo, nome, ncaixa, cnpj))
        conn.commit()
        conn.close()

        return redirect(url_for('index'))
    
    return render_template('novo_cliente.html')

##@app.route('/limpar_pacientes')
##def limpar_pacientes():
    ##conn = sqlite3.connect('gestao_hospitalar.db')
    ##cursor = conn.cursor()
    ##cursor.execute('DELETE  FROM pacientes')
    ##conn.commit()
    ##conn.close()
    ##return redirect(url_for('index'))


@app.route('/lista_clientes')
def lista_clientes():
    busca = request.args.get('q','').strip()

    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()
    if busca:
        cursor.execute('''
          SELECT codigo, nome, ncaixa FROM clientes 
          WHERE nome LIKE ? OR CAST(codigo AS TEXT) LIKE ?
          ORDER BY nome''',(f'%{busca}%',f'%{busca}%',))
    else:
        cursor.execute('SELECT codigo, nome, ncaixa FROM clientes ORDER BY nome')
    rows = cursor.fetchall()
    conn.close()

    pacientes_por_letra = defaultdict(list)

    for (codigo,nome,ncaixa) in rows:
        inicial = nome[0].upper()
        pacientes_por_letra[inicial].append((codigo, nome, ncaixa))

    return render_template('lista_clientes.html', pacientes_por_letra=sorted(pacientes_por_letra.items()))



@app.route('/lista_prod')
def lista_prod():
    busca = request.args.get('q','').strip()

    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()
    if busca:
        cursor.execute('''
          SELECT codigo, descricao FROM produtos 
          WHERE descricao LIKE ? OR CAST(codigo AS TEXT) LIKE ?
          ORDER BY descricao''',(f'%{busca}%',f'%{busca}%',))
    else:
        cursor.execute('SELECT codigo, descricao FROM produtos ORDER BY descricao')
    rows = cursor.fetchall()
    conn.close()

    produtos_por_letra = defaultdict(list)

    for (codigo,descricao) in rows:
        inicial = descricao[0].upper()
        produtos_por_letra[inicial].append((codigo, descricao))

    return render_template('lista_prod.html', produtos_por_letra=sorted(produtos_por_letra.items()))

@app.route('/cadastro_produto', methods=['GET', 'POST'])
def cadastro_produto():
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()

    if request.method == 'POST':
        paciente_nome = request.form.get('paciente_nome')
        codigo_cliente = paciente_nome.split(" - ")[0].strip()

        codigo = request.form['codigo']
        descricao = request.form['descricao']
        os = request.form['os']
        clicheria = request.form['clicheria']
        qtde = request.form['qtde']
        data = request.form['data']
        seq_caixa = request.form['seq_caixa']
        cursor.execute('''
            INSERT INTO produtos (codigo_cliente, codigo, descricao, os, clicheria, qtde, data, seq_caixa)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (codigo_cliente, codigo, descricao, os, clicheria, qtde, data, seq_caixa))
        id = cursor.lastrowid
        cursor.execute('''
            INSERT INTO estoque (item, obs, cod_local,data)
            VALUES(?, 'Cliche novo - '||?, ?, ?)
                       ''',(codigo,clicheria, 2,data)) 
        cursor.execute('''
            INSERT INTO movestoque (item, data, obs_mov, cod_local)
            VALUES(?, ?, 'Cliche novo - '||?, ?)
                       ''',(f"{codigo} - {descricao}", data, clicheria, 2))   
        conn.commit()
        print(">>> ID inserido:", id)
        buscar_e_imprimir_atual(id)
        print(">>> buscar_e_imprimir chamada")
        ##conn.close()

    cursor.execute('SELECT codigo, nome FROM clientes ORDER BY nome')
    dados = cursor.fetchall()
    conn.close()
    return render_template('cadastro_produto.html', dados=dados)

def imprimir_etiqueta_ppla(codigo_cliente, nome_cliente, codigo, descricao, os, clicheria, caixa, qtde, data, seq_caixa): 
    data_formatada = f"{data[8:10]}/{data[5:7]}/{data[2:4]}"
    etiqueta = f"""

L
m
e
K1504
L
C0005
H12D11
m
111100004000070{codigo_cliente}-{nome_cliente[:39]}
111100003200070{codigo}-{descricao}
121100002350200QTDE:{qtde}
111100001000040Clicheria:{clicheria}
111100000200100OS:{os}
111100000200500Data: {data_formatada}
121100001600050Caixa: {caixa} - {seq_caixa}
1E3100001000650{codigo}
Q0001
E

"""

    nome_impressora = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(nome_impressora)

    try:
        win32print.StartDocPrinter(hprinter, 1, ("Etiqueta PPLA", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, etiqueta.encode('utf8'))
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
    finally:
        win32print.ClosePrinter(hprinter)


def buscar_e_imprimir_atual(id_produto):
    conn = sqlite3.connect('gestao_cliches.db')
    cur = conn.cursor()
    cur.execute("SELECT p.codigo_cliente, c.nome, p.codigo, p.descricao, p.os, p.clicheria, c.ncaixa, p.qtde, p.data, p.seq_caixa FROM produtos p join clientes c ON p.codigo_cliente = c.codigo WHERE id = ?", (id_produto,))
    dados = cur.fetchone()
    conn.close()
    print(">>> Resultado da busca:", dados) 
    if dados:
                print(">>> Dados encontrados:", dados)
                imprimir_etiqueta_ppla(*dados)
    else:
                print(">>> Nenhum dado encontrado para o ID:", id_produto)

@app.route('/imprimir', methods=['GET', 'POST'])
def imprimir_etq():
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()

    if request.method == 'POST':
        produto_nome = request.form.get('produto_nome')
        if not produto_nome or " - " not in produto_nome:
            conn.close()
            return "Código do cliente inválido", 400

        codigo = produto_nome.split(" - ")[0].strip()
        seq_caixa = request.form.get('seq_caixa')
        cursor.execute('''
            UPDATE produtos
            SET seq_caixa = ?
            WHERE codigo = ?
                ''', (seq_caixa, codigo))
        conn.commit()
        buscar_e_imprimir(codigo)
  
    cursor.execute('SELECT codigo, descricao FROM produtos ORDER BY descricao')
    dadosprod = cursor.fetchall()
    conn.close()
    conn.commit
    return render_template('imprimir.html', dadosprod=dadosprod)

def imprimir_etq(codigo_cliente, nome_cliente, codigo, descricao, os, clicheria, caixa, qtde, data, seq_caixa): 
    data_formatada = f"{data[8:10]}/{data[5:7]}/{data[2:4]}" 
    etiqueta = f"""
   
L
m
e
K1504
L
C0005
H12D11
m
111100004000050{codigo_cliente}-{nome_cliente[:39]}
111100003200050{codigo}-{descricao}
121100002350100QTDE:{qtde}
111100001000040Clicheria:{clicheria}
111100000200100OS:{os}
111100000200500Data: {data_formatada}
121100001600050Caixa: {caixa} - {seq_caixa}
1E3100001000650{codigo}
Q0001
E
"""

    etiqueta_cp850 = etiqueta.encode("cp850")

    nome_impressora = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(nome_impressora)

    try:
        win32print.StartDocPrinter(hprinter, 1, ("Etiqueta PPLA", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, etiqueta_cp850)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
    finally:
        win32print.ClosePrinter(hprinter)

def buscar_e_imprimir(id_produtos):
    conn = sqlite3.connect('gestao_cliches.db')
    cur = conn.cursor()

    cur.execute("SELECT p.codigo_cliente, c.nome, p.codigo, p.descricao, p.os, p.clicheria, c.ncaixa, p.qtde, p.data, p.seq_caixa FROM produtos p join clientes c ON p.codigo_cliente = c.codigo WHERE p.codigo = ?", (id_produtos,))
    dados = cur.fetchone()
    conn.close()
    if dados:
        print("Chamando a impressora com:", dados)
        imprimir_etq(*dados)

@app.route('/cadastro_caixa', methods=['GET', 'POST'])
def cadastro_caixa():
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()


    if request.method == 'POST':
        paciente_nome = request.form.get('paciente_nome')
        if not paciente_nome or " - " not in paciente_nome:
            conn.close()
            mensagem = "Código do cliente inválido"
        else:
            codigo_cliente = paciente_nome.split(" - ")[0].strip()
            ncaixa = request.form.get('ncaixa')

            cursor.execute('''
            UPDATE clientes
            SET ncaixa = ?
            WHERE codigo = ?
            ''', (ncaixa, codigo_cliente))

            conn.commit()
            mensagem = "Caixa cadastrada com sucesso"

    cursor.execute('SELECT codigo, nome FROM clientes ORDER BY nome')
    dados = cursor.fetchall()
    mensagem = cursor.fetchall()
    conn.close()

    return render_template('cadastro_caixa.html', dados=dados, mensagem=mensagem)

@app.route('/movestoque', methods=['GET', 'POST'])
def movestoque():
    conn = sqlite3.connect('gestao_cliches.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    produtos_enviados = []
    if request.method == 'POST':
        estnome = request.form.get('estnome')
        if estnome:
         cod_local = estnome.split(" - ")[0].strip()
        else:
            cod_local = None
        for i in range(10):
            obs_mov = request.form[f'obs_mov_{i}']
            produto_nome = request.form.get(f'produto_nome_{i}')
      
            
            # Ignora se estiver vazio
            if not produto_nome or not estnome:
                continue
            
            try:
                produto_codigo = produto_nome.split(" - ")[0].strip()
                cursor.execute('SELECT melhorias_qualidade FROM produtos WHERE codigo = ?', (produto_codigo,))
                result = cursor.fetchone()
         
                melhorias_qualidade = result[0] if result else None
                print(f"[DADOS] produto_codigo={produto_codigo}, melhorias_qualidade={melhorias_qualidade}")

                if cod_local == "2" and melhorias_qualidade:
                    produtos_enviados.append(produto_codigo)
                
                # Grava a movimentação
                cursor.execute('''
                    INSERT INTO movestoque (item, cod_local, data, obs_mov)
                    VALUES (?, ?, datetime('now', '-3 hours'), ?)
                ''', (produto_nome,cod_local, obs_mov))

                             # Atualiza estoque ou deleta dependendo do local
                if cod_local == "4":
                        cursor.execute('DELETE FROM estoque WHERE item = ?', (produto_codigo,))
                        cursor.execute('DELETE FROM produtos WHERE codigo = ?', (produto_codigo,))           
                else:
                    cursor.execute('''
                        UPDATE estoque
                        SET cod_local = ?, obs = ?
                        WHERE item = ?
                    ''', (cod_local, obs_mov, produto_codigo))

            except Exception as e:
                print(f"Erro ao processar linha {i}: {e}")  # Log para debug       
        conn.commit()
        
        if cod_local == "2" and produtos_enviados:
         print("[DEBUG FINAL] Produtos a enviar:", produtos_enviados)
         buscar_e_enviar(produtos_enviados)

        conn.close()
        
        return redirect('/movestoque')

    # GET → carrega os dados
    cursor.execute('SELECT codigo, descricao FROM produtos ORDER BY descricao')
    produtos = cursor.fetchall()
    

    cursor.execute('SELECT id, desclocal FROM locais ORDER BY id')
    locais = cursor.fetchall()
     

    conn.close()
    return render_template('movestoque.html', produtos=produtos, locais=locais)
  
def buscar_e_enviar(lista_codigos):
    if not lista_codigos:
        print("Nenhum código informado para envio.")
        return

    # Conecta ao banco SQLite
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()

    # Cria os placeholders (?, ?, ?, ..., ?)
    placeholders = ','.join(['?'] * len(lista_codigos))
    query = f'''
        SELECT codigo, descricao, melhorias_qualidade
        FROM produtos
        WHERE codigo IN ({placeholders})
    '''
    cursor.execute(query, lista_codigos)
    resultados = cursor.fetchall()
    conn.close()

    if not resultados:
        print("Nenhum resultado encontrado para os códigos informados.")
        return

    # Monta corpo do e-mail
    corpo_email = "<h2>Relatório de Melhorias</h2><ul>"
    for codigo, descricao, melhorias in resultados:
        corpo_email += f"<li>{codigo} - {descricao} - {melhorias}</li>"
    corpo_email += "</ul>"

    # Configurações do e-mail
    remetente = "xxx@google.com"
    senha = "asdfg"
    destinatario = "xxx@google.com"

    # Usa o primeiro item como parte do assunto
    assunto = f"Melhorias - {resultados[0][0]} - {resultados[0][1]}"

    # Monta e-mail
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo_email, 'html'))
    #configurar o email aqui
    try:
        with smtplib.SMTP_SSL() as servidor:
            servidor.login(remetente, senha)
            servidor.send_message(msg)
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")




@app.route('/melhorias', methods=['GET', 'POST'])
def melhorias():
    conn = sqlite3.connect('gestao_cliches.db')
    cursor = conn.cursor()

    if request.method == 'POST':
        produto_nome = request.form.get('produto_nome')
        if not produto_nome or " - " not in produto_nome:
            conn.close()
            return "Código do produto inválido", 400

        codigo = produto_nome.split(" - ")[0].strip()
        melhorias_qualidade = request.form.get('melhorias_qualidade')
        cursor.execute('''
            UPDATE produtos
            SET melhorias_qualidade = ?
            WHERE codigo = ?
                ''', (melhorias_qualidade, codigo))
        conn.commit()

        return redirect('/melhorias')
    
    cursor.execute('SELECT codigo, descricao FROM produtos ORDER BY descricao')
    dadosprod = cursor.fetchall()
    conn.close()
    conn.commit
    
    return render_template('melhorias.html', dadosprod=dadosprod)
    

    
        
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
