UPDATE clientes_financeiro_valores SET 
percentual_human = %s,
economia_mensal = %s, 
economia_liquida = %s,
total_fatura = %s 
WHERE 
cliente_id = %s AND mes = %s AND ano = %s