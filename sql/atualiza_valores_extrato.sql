UPDATE clientes_financeiro_valores
SET convenio_farmacia=%s, 
adiant_salarial=%s, 
numero_empregados=%s, 
numero_estagiarios=%s, 
trabalhando=%s, 
salario_contri_empregados=%s, 
salario_contri_contribuintes=%s,
soma_salarios_provdt=%s, 
inss=%s, 
fgts=%s, 
irrf=%s, 
salarios_pagar=%s, 
updated_at=NOW()
WHERE cliente_id = %s AND mes = %s AND ano = %s