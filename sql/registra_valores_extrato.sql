INSERT INTO clientes_financeiro_valores 
(cliente_id, cod_empresa, convenio_farmacia, adiant_salarial, numero_empregados, 
numero_estagiarios, trabalhando, salario_contri_empregados, 
salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, irrf, 
salarios_pagar, mes, ano, anexo_enviado, relatorio_enviado, created_at, updated_at)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NOW(), NOW())