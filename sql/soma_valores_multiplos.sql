SELECT 
cliente_id, 
ROUND(SUM(convenio_farmacia), 2) AS convenio_farmacia,
ROUND(SUM(adiant_salarial), 2) AS adiant_salarial,
SUM(numero_empregados) AS numero_empregados,
SUM(numero_estagiarios) AS numero_estagiarios,
SUM(trabalhando) AS trabalhando,
ROUND(SUM(salario_contri_empregados), 2) AS salario_contri_empregados,
ROUND(SUM(salario_contri_contribuintes), 2) AS salario_contri_contribuintes,
ROUND(SUM(soma_salarios_provdt), 2) AS soma_salarios_provdt,
ROUND(SUM(inss), 2) AS inss,
ROUND(SUM(fgts), 2) AS fgts,
ROUND(SUM(irrf), 2) AS irrf,
ROUND(SUM(salarios_pagar), 2) AS salarios_pagar,
ROUND(SUM(mensal_ponto_elet), 2) AS mensal_ponto_elet,
ROUND(SUM(assinat_eletronica), 2) AS assinat_eletronica,
ROUND(SUM(vale_transporte), 2) AS vale_transporte,
ROUND(SUM(vale_refeicao), 2) AS vale_refeicao,
ROUND(SUM(saude_seguranca_trabalho), 2) AS saude_seguranca_trabalho,
ROUND(SUM(percentual_human), 2) AS percentual_human,
ROUND(SUM(economia_mensal), 2) AS economia_mensal,
ROUND(SUM(total_fatura), 2) AS total_fatura
FROM clientes_financeiro_valores WHERE cliente_id = %s AND mes = %s and ano = %s
GROUP BY cliente_id;