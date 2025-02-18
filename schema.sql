CREATE TABLE modelos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL,
    conteudo TEXT NOT NULL,
    data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE recibos_gerados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    modelo_id INTEGER,
    cliente_nome TEXT NOT NULL,
    valor DECIMAL(10,2) NOT NULL,
    numero_documento TEXT NOT NULL,
    data_geracao DATETIME DEFAULT CURRENT_TIMESTAMP,
    documento_blob BLOB,
    FOREIGN KEY (modelo_id) REFERENCES modelos(id)
);

CREATE TABLE receipt_sequence (
    id INTEGER PRIMARY KEY,
    last_number INTEGER NOT NULL DEFAULT 0
);

-- Initialize with first record
INSERT INTO receipt_sequence (id, last_number) VALUES (1, 0);
