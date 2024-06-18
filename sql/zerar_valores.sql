UPDATE clientes_financeiro_valores
SET
    convenio_farmacia = 0,
    adiant_salarial = 0,
    numero_empregados = 0,
    numero_estagiarios = 0,
    trabalhando = 0,
    salario_contri_empregados = 0,
    salario_contri_contribuintes = 0,
    soma_salarios_provdt = 0,
    inss = 0,
    fgts = 0,
    irrf = 0,
    salarios_pagar = 0,
    percentual_human = 0,
    economia_mensal = 0,
    economia_liquida = 0,
    total_fatura = 0,
    anexo_enviado = 0
WHERE
    mes = %s AND ano = %s AND cliente_id = %s