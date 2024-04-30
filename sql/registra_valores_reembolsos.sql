INSERT INTO clientes_financeiro_reembolsos
(cliente_id, descricao, valor, mes, ano, created_at, updated_at)
VALUES (%s, %s, %s, %s, %s, NOW(), NOW())