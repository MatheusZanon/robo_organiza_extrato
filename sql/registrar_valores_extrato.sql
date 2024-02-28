INSERT INTO clientes_financeiro_valores 
(cliente_id, cod_empresa, convenio_farmacia, adiant_salarial, numero_empregados, 
numero_estagiarios, trabalhando, salario_contri_empregados, 
salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, irrf, 
salarios_pagar, vale_transporte, assinat_eletronica, vale_refeicao, mensal_ponto_elet, saude_seguranca_trabalho, mes, ano)
VALUES (%s, %s, %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s, %s, %s, %s,%s, %s, %s, %s, %s)