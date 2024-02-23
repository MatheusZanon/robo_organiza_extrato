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
mensal_ponto_elet,
assinat_eletronica,
vale_transporte,
vale_refeicao,
saude_seguranca_trabalho,
percentual_human,
economia_mensal,
total_fatura
FROM clientes_financeiro_valores WHERE 
cliente_id = %s AND cod_empresa = %s AND mes = %s AND ano = %s