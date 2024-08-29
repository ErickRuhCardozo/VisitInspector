import re
import requests


class EinQueryer:
    """
    This class is responsible for querying an EIN on an online service and
    return the appropriate info about the queried Establishment.
    """


    SERVICE_URL = 'https://api-publica.speedio.com.br/buscarcnpj'


    @staticmethod
    def query(ein: str) -> dict | None:
        ein = re.sub(r'\D', '', ein)
        res = requests.get(EinQueryer.SERVICE_URL, {'cnpj': ein})
        
        if not res.ok:
            return None
        
        res = res.json()
        establishment = dict()
        establishment['region'] = res['BAIRRO'] if 'BAIRRO' in res else '?'

        if 'NOME FANTASIA' in res and res['NOME FANTASIA'] == '':
            establishment['name'] = res['RAZAO SOCIAL'] if 'RAZAO SOCIAL' in res else '???'
        else:
            establishment['name'] = res['NOME FANTASIA'] if 'NOME FANTASIA' in res else '???'

        establishment['ein'] = EinQueryer.format_ein(res['CNPJ'])

        if 'LOGRADOURO' in res and 'NUMERO' in res:
            establishment['address'] = f'{res['LOGRADOURO']}, {res['NUMERO']}'
        else:
            establishment['address'] = '???'

        return establishment
    

    @staticmethod
    def format_ein(ein: str) -> str:
        return f'{ein[:8]}/{ein[8:12]}-{ein[12:]}'