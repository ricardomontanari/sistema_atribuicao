import sqlite3

# Nome do arquivo do banco de dados SQLite
DB_NAME = 'automacao_db.sqlite'

# #####################################################################
# --- INICIALIZAÇÃO DO BANCO DE DADOS ---
# #####################################################################

def setup_database():
    """
    Inicializa o banco de dados e cria a tabela 'cidades' se ela não existir.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Tabela 1: Cidades 
    # A restrição UNIQUE garante que não haja nomes de cidades duplicados
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cidades (
            id INTEGER PRIMARY KEY,
            nome TEXT UNIQUE NOT NULL
        )
    """)
    
    conn.commit()
    conn.close()
    
# #####################################################################
# --- FUNÇÕES DE CADASTRO (INSERT) ---
# #####################################################################

def adicionar_cidade(nome_cidade):
    """Adiciona uma nova cidade à tabela de cidades. Retorna mensagem e status (True/False)."""
    nome_cidade = nome_cidade.strip().upper()
    if not nome_cidade:
        return "❌ Erro: O nome da cidade não pode ser vazio.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # Usamos '?' para prevenção contra injeção SQL
        cursor.execute("INSERT INTO cidades (nome) VALUES (?)", (nome_cidade,))
        conn.commit()
        return f"✅ Cidade '{nome_cidade}' cadastrada com sucesso!", True
    except sqlite3.IntegrityError:
        return f"⚠️ Cidade '{nome_cidade}' já existe no banco de dados.", False
    except Exception as e:
        return f"❌ Erro ao adicionar cidade: {e}", False
    finally:
        conn.close()

# #####################################################################
# --- FUNÇÕES DE BUSCA (SELECT) ---
# #####################################################################

def buscar_nomes_cidades():
    """Retorna uma lista de nomes de cidades ordenadas alfabeticamente (para ComboBox)."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT nome FROM cidades ORDER BY nome ASC")
        # Retorna uma lista de strings (nomes), desempacotando as tuplas
        return [cidade[0] for cidade in cursor.fetchall()]
    except Exception:
        return []
    finally:
        conn.close()

def listar_cidades():
    """Retorna uma lista de tuplas (id, nome) de todas as cidades (para ListBox)."""
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
    """Busca o nome de uma cidade usando seu ID."""
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

# #####################################################################
# --- FUNÇÕES DE EXCLUSÃO (DELETE) ---
# #####################################################################

def excluir_cidade(cidade_id):
    """Exclui uma cidade da tabela usando seu ID."""
    try:
        cidade_id = int(cidade_id)
    except ValueError:
        return "❌ Erro: O ID deve ser um número inteiro.", False
        
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        # Busca o nome antes de deletar para uma mensagem de confirmação melhor
        cursor.execute("SELECT nome FROM cidades WHERE id = ?", (cidade_id,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return f"⚠️ Cidade com ID {cidade_id} não encontrada.", False
            
        nome_cidade = resultado[0]
        
        cursor.execute("DELETE FROM cidades WHERE id = ?", (cidade_id,))
        conn.commit()
        
        return f"✅ Cidade '{nome_cidade}' (ID {cidade_id}) excluída com sucesso.", True
        
    except Exception as e:
        return f"❌ Erro ao excluir cidade (ID {cidade_id}): {e}", False
    finally:
        conn.close()