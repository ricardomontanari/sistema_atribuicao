import sqlite3
import hashlib
import json
import os

# #####################################################################
# --- CONFIGURAÇÃO DO NOME DO BANCO (DINÂMICO) ---
# #####################################################################

CONFIG_FILE = 'config_cliente.json'

def obter_nome_banco():
    """
    Define o nome do arquivo do banco de dados.
    1. Tenta ler 'config_cliente.json' para buscar o nome do cliente.
    2. Se encontrar, gera 'automacao_db_NOMECLIENTE.sqlite'.
    3. Se não encontrar, usa o padrão 'automacao_db.sqlite'.
    """
    nome_padrao = 'automacao_db.sqlite'
    
    # Procura o arquivo de config no mesmo diretório do executável/script
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                cliente = data.get('cliente', '').strip()
                
                if cliente:
                    # Limpa o nome para evitar caracteres inválidos em arquivos
                    # Ex: "São Paulo" vira "sao_paulo"
                    safe_name = "".join([c for c in cliente if c.isalnum() or c in (' ', '-', '_')]).strip()
                    safe_name = safe_name.replace(' ', '_').lower()
                    
                    db_name = f'automacao_db_{safe_name}.sqlite'
                    print(f"🔧 Configuração detectada! Usando banco: {db_name}")
                    return db_name
                    
        except Exception as e:
            print(f"⚠️ Erro ao ler config do cliente: {e}")
            
    return nome_padrao

# Define a constante globalmente para ser usada nas funções abaixo
DB_NAME = obter_nome_banco()

# #####################################################################
# --- INICIALIZAÇÃO DO BANCO DE DADOS ---
# #####################################################################

def setup_database():
    """
    Inicializa o banco de dados, cria as tabelas necessárias e
    garante que exista um usuário 'admin' padrão.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Tabela 1: Cidades 
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cidades (
            id INTEGER PRIMARY KEY,
            nome TEXT UNIQUE NOT NULL
        )
    """)

    # Tabela 2: Usuários
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL
        )
    """)
    
    # Cria usuário admin padrão se não existir (senha: admin)
    try:
        senha_hash = hashlib.sha256("admin".encode()).hexdigest()
        cursor.execute("INSERT OR IGNORE INTO usuarios (username, password) VALUES (?, ?)", ("admin", senha_hash))
        conn.commit()
    except Exception as e:
        print(f"Erro ao criar admin padrão: {e}")
    
    conn.commit()
    conn.close()
    
# #####################################################################
# --- FUNÇÕES: CIDADES ---
# #####################################################################

def adicionar_cidade(nome_cidade):
    """Adiciona uma nova cidade."""
    nome_cidade = nome_cidade.strip().upper()
    if not nome_cidade:
        return "❌ Erro: O nome da cidade não pode ser vazio.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO cidades (nome) VALUES (?)", (nome_cidade,))
        conn.commit()
        return f"✅ Cidade '{nome_cidade}' cadastrada com sucesso!", True
    except sqlite3.IntegrityError:
        return f"⚠️ Cidade '{nome_cidade}' já existe.", False
    except Exception as e:
        return f"❌ Erro: {e}", False
    finally:
        conn.close()

def buscar_nomes_cidades():
    """Retorna lista de nomes (strings) para o ComboBox."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT nome FROM cidades ORDER BY nome ASC")
        return [cidade[0] for cidade in cursor.fetchall()]
    except Exception:
        return []
    finally:
        conn.close()

def listar_cidades():
    """Retorna lista de tuplas (id, nome) para a Listagem."""
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
    """Busca o nome de uma cidade pelo ID (útil para mensagens de confirmação)."""
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
    """Exclui uma cidade pelo ID."""
    try:
        cidade_id = int(cidade_id)
    except ValueError:
        return "❌ Erro: O ID deve ser um número inteiro.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT nome FROM cidades WHERE id = ?", (cidade_id,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return f"⚠️ Cidade com ID {cidade_id} não encontrada.", False
            
        nome_cidade = resultado[0]
        cursor.execute("DELETE FROM cidades WHERE id = ?", (cidade_id,))
        conn.commit()
        return f"✅ Cidade '{nome_cidade}' excluída.", True
        
    except Exception as e:
        return f"❌ Erro: {e}", False
    finally:
        conn.close()

# #####################################################################
# --- FUNÇÕES: USUÁRIOS ---
# #####################################################################

def verificar_credenciais(username, password):
    """Verifica se o usuário e senha (hash) conferem no banco."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # Gera o hash da senha informada para comparar com o banco
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        
        cursor.execute("SELECT username FROM usuarios WHERE username = ? AND password = ?", (username, pwd_hash))
        return cursor.fetchone() is not None
    except Exception:
        return False
    finally:
        conn.close()

def adicionar_usuario(username, password):
    """Cadastra um novo usuário com senha hash."""
    username = username.strip().lower() # Padroniza minusculo para login
    if not username or not password:
        return "❌ Usuário e Senha são obrigatórios.", False
    
    if len(password) < 4:
        return "⚠️ A senha deve ter pelo menos 4 caracteres.", False

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        cursor.execute("INSERT INTO usuarios (username, password) VALUES (?, ?)", (username, pwd_hash))
        conn.commit()
        return f"✅ Usuário '{username}' cadastrado!", True
    except sqlite3.IntegrityError:
        return f"⚠️ O usuário '{username}' já existe.", False
    except Exception as e:
        return f"❌ Erro: {e}", False
    finally:
        conn.close()

def listar_usuarios():
    """Retorna lista de todos os usernames cadastrados."""
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
    """Exclui um usuário, mas impede a exclusão do 'admin'."""
    username = username.strip().lower()
    
    if username == "admin":
        return "⛔ AÇÃO NEGADA: Você não pode excluir o administrador principal.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # Verifica se existe antes de tentar apagar
        cursor.execute("SELECT username FROM usuarios WHERE username = ?", (username,))
        if not cursor.fetchone():
             return f"⚠️ Usuário '{username}' não encontrado.", False
             
        cursor.execute("DELETE FROM usuarios WHERE username = ?", (username,))
        conn.commit()
        return f"✅ Usuário '{username}' excluído com sucesso.", True
    except Exception as e:
        return f"❌ Erro: {e}", False
    finally:
        conn.close()