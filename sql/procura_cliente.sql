SELECT 
id, 
nome_razao_social, 
cnpj, 
cpf, 
email, 
telefone_celular, 
regiao  
FROM clientes_financeiro WHERE nome_razao_social = %s