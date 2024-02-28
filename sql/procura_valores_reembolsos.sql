SELECT
descricao, valor
FROM clientes_financeiro_reembolsos WHERE
cliente_id = %s AND mes = %s AND ano = %s