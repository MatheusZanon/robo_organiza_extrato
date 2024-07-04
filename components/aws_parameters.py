import boto3

def get_regin_name():
    ssm = boto3.client('ssm', region_name='sa-east-1') #definindo região padrão
    try:
        parameter = ssm.get_parameter(Name='/human/REGION_NAME', WithDecryption=True)
        return parameter['Parameter']['Value']
    except ssm.exceptions.ParameterNotFound:
        return 'sa-east-1'
    
def get_ssm_parameter(name, default=None):
    ssm = boto3.client('ssm', get_regin_name())
    try:
        parameter = ssm.get_parameter(Name=name, WithDecryption=True)
        return parameter['Parameter']['Value']
    except ssm.exceptions.ParameterNotFound:
        return default
