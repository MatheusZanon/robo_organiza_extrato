SELECT 
id, 
nome_razao_social, 
cnpj, 
cpf, 
email, 
telefone_celular, 
regiao,
is_active  
FROM clientes_financeiro WHERE nome_razao_social = %s