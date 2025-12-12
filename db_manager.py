import sqlite3
import hashlib
import json
import os

# #####################################################################
# --- CONFIGURAÇÃO E INICIALIZAÇÃO ---
# #####################################################################

CONFIG_FILE = 'config_cliente.json'
DB_NAME_DEFAULT = 'automacao_db.sqlite'

def obter_nome_banco():
    """
    Define o nome do arquivo do banco de dados dinamicamente.
    Lê 'config_cliente.json' para personalizar o banco por cliente.
    """
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                cliente = data.get('cliente', '').strip()
                
                if cliente:
                    # Sanitiza o nome (remove caracteres especiais)
                    safe_name = "".join([c for c in cliente if c.isalnum() or c in (' ', '-', '_')]).strip()
                    safe_name = safe_name.replace(' ', '_').lower()
                    return f'automacao_db_{safe_name}.sqlite'
        except Exception:
            pass # Falha silenciosa, usa o padrão
            
    return DB_NAME_DEFAULT

# Define o nome do banco globalmente
DB_NAME = obter_nome_banco()

def setup_database():
    """
    Inicializa o banco de dados.
    Cria as tabelas 'cidades' e 'usuarios' e insere o administrador padrão.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # 1. Tabela de Cidades
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS cidades (
                id INTEGER PRIMARY KEY,
                nome TEXT UNIQUE NOT NULL
            )
        """)

        # 2. Tabela de Usuários
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL
            )
        """)

        # 3. Cria usuário 'admin' padrão se não existir (Senha: @admin@)
        cursor.execute("SELECT * FROM usuarios WHERE username = 'admin'")
        if not cursor.fetchone():
            senha_hash = hashlib.sha256("@admin@".encode()).hexdigest()
            cursor.execute("INSERT INTO usuarios (username, password_hash) VALUES (?, ?)", ('admin', senha_hash))
            print(f"🔧 [DB] Banco '{DB_NAME}' inicializado. Admin padrão criado.")
            
        conn.commit()
    except Exception as e:
        print(f"❌ [DB] Erro crítico na inicialização: {e}")
    finally:
        conn.close()

# #####################################################################
# --- MÓDULO: CIDADES (CRUD) ---
# #####################################################################

def adicionar_cidade(nome_cidade):
    """
    Adiciona uma nova cidade.
    Retorna: (Mensagem, Sucesso[bool])
    """
    nome_cidade = nome_cidade.strip().upper()
    if not nome_cidade:
        return "❌ Erro: O nome da cidade não pode ser vazio.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO cidades (nome) VALUES (?)", (nome_cidade,))
        conn.commit()
        return f"✅ Cidade '{nome_cidade}' adicionada com sucesso!", True
    except sqlite3.IntegrityError:
        return f"⚠️ A cidade '{nome_cidade}' já está cadastrada.", False
    except Exception as e:
        return f"❌ Erro técnico: {e}", False
    finally:
        conn.close()

def buscar_nomes_cidades():
    """Retorna uma lista de strings com os nomes de todas as cidades (Ordem Alfabética)."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT nome FROM cidades ORDER BY nome ASC")
        return [row[0] for row in cursor.fetchall()]
    except Exception:
        return []
    finally:
        conn.close()

def listar_cidades():
    """Retorna uma lista de tuplas (id, nome) de todas as cidades."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, nome FROM cidades ORDER BY nome ASC")
        return cursor.fetchall()
    except Exception:
        return []
    finally:
        conn.close()

def buscar_nome_cidade_por_id(cidade_id):
    """Retorna o nome da cidade baseado no ID, ou None se não existir."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT nome FROM cidades WHERE id = ?", (cidade_id,))
        resultado = cursor.fetchone()
        return resultado[0] if resultado else None
    except Exception:
        return None
    finally:
        conn.close()

def excluir_cidade(cidade_id):
    """
    Exclui uma cidade pelo ID.
    Retorna: (Mensagem, Sucesso[bool])
    """
    try:
        cidade_id = int(cidade_id)
    except ValueError:
        return "❌ Erro: O ID deve ser um número inteiro.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # Busca o nome antes para confirmar na mensagem
        cursor.execute("SELECT nome FROM cidades WHERE id = ?", (cidade_id,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return f"⚠️ Cidade com ID {cidade_id} não encontrada.", False
            
        nome_cidade = resultado[0]
        
        cursor.execute("DELETE FROM cidades WHERE id = ?", (cidade_id,))
        conn.commit()
        return f"✅ Cidade '{nome_cidade}' (ID {cidade_id}) excluída.", True
        
    except Exception as e:
        return f"❌ Erro ao excluir: {e}", False
    finally:
        conn.close()

# #####################################################################
# --- MÓDULO: USUÁRIOS (AUTH & CRUD) ---
# #####################################################################

def verificar_credenciais(username, password):
    """
    Verifica se o usuário e senha (hash) conferem.
    Retorna True/False.
    """
    username = username.strip().lower()
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT password_hash FROM usuarios WHERE username = ?", (username,))
        resultado = cursor.fetchone()
        
        if resultado and resultado[0] == password_hash:
            return True
        return False
    except Exception:
        return False
    finally:
        conn.close()

def adicionar_usuario(username, password):
    """
    Cadastra novo usuário (senha é salva como hash SHA256).
    """
    username = username.strip().lower()
    if not username or not password:
        return "❌ Usuário e Senha são obrigatórios.", False
        
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO usuarios (username, password_hash) VALUES (?, ?)", (username, password_hash))
        conn.commit()
        return f"✅ Usuário '{username}' cadastrado com sucesso!", True
    except sqlite3.IntegrityError:
        return f"⚠️ O usuário '{username}' já existe.", False
    except Exception as e:
        return f"❌ Erro técnico: {e}", False
    finally:
        conn.close()

def listar_usuarios():
    """Retorna lista de todos os usernames."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT username FROM usuarios ORDER BY username ASC")
        return [row[0] for row in cursor.fetchall()]
    except Exception:
        return []
    finally:
        conn.close()

def excluir_usuario(username):
    """
    Exclui um usuário. Impede a exclusão do 'admin'.
    """
    username = username.strip().lower()
    
    if username == "admin":
        return "⛔ AÇÃO NEGADA: Não é permitido excluir o administrador principal.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT username FROM usuarios WHERE username = ?", (username,))
        if not cursor.fetchone():
             return f"⚠️ Usuário '{username}' não encontrado.", False
             
        cursor.execute("DELETE FROM usuarios WHERE username = ?", (username,))
        conn.commit()
        return f"✅ Usuário '{username}' removido.", True
    except Exception as e:
        return f"❌ Erro técnico: {e}", False
    finally:
        conn.close()