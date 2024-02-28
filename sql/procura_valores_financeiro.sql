SELECT 
cliente_id, 
convenio_farmacia,
adiant_salarial,
numero_empregados,
numero_estagiarios,
trabalhando,
salario_contri_empregados,
salario_contri_contribuintes,
soma_salarios_provdt,
inss,
fgts,
irrf,
salarios_pagar,
vale_transporte,
assinat_eletronica,
vale_refeicao,
mensal_ponto_elet,
saude_seguranca_trabalho,
percentual_human,
economia_mensal,
total_fatura
FROM clientes_financeiro_valores WHERE 
cliente_id = %s AND mes = %s AND ano = %s