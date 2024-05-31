UPDATE clientes_financeiro_valores
SET
    convenio_farmacia = NULL,
    adiant_salarial = NULL,
    numero_empregados = NULL,
    numero_estagiarios = NULL,
    trabalhando = NULL,
    salario_contri_empregados = NULL,
    salario_contri_contribuintes = NULL,
    soma_salarios_provdt = NULL,
    inss = NULL,
    fgts = NULL,
    irrf = NULL,
    salarios_pagar = NULL,
    percentual_human = NULL,
    economia_mensal = NULL,
    economia_liquida = NULL,
    total_fatura = NULL,
    anexo_enviado = 0
WHERE
    mes = %s AND ano = %s AND cliente_id = %s