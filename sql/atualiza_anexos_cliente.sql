UPDATE clientes_financeiro_valores 
SET anexo_enviado = 1 
WHERE cliente_id = %s AND mes = %s AND ano = %s