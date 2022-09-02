import requests as rq
import json


def consultar_endereco(dado):
    session = rq.Session()

    url = f'https://buscacepinter.correios.com.br/app/endereco/carrega-cep-endereco.php'

    headers = {
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36',
        'Host': 'buscacepinter.correios.com.br',
        'Origin': 'https://buscacepinter.correios.com.br',
        'Referer': 'https://buscacepinter.correios.com.br/app/endereco/index.php'
        }

    data = {
        'pagina': '/app/endereco/index.php',
        'cepaux': '',
        'mensagem_alerta':'',
        'endereco': dado,
        'tipoCEP': 'ALL'
    }

    endereco = {}

    try: 
        res = session.post(url, headers=headers, data=data)
        dados = json.loads(res.text)['dados'][0]

        endereco['estado'] = dados['uf']
        endereco['cidade'] = dados['localidade']
        endereco['logradouro'] = dados['logradouroDNEC']
        endereco['bairro'] = dados['bairro']
        endereco['cep'] = dados['cep']
    except IndexError:
        print(json.loads(res.text)['mensagem'])

    return endereco


cep = '52090725'

if len(cep) == 10:
    cep1 = ''.join(cep.split('.'))
    cep = ''.join(cep1.split('-'))

ende = consultar_endereco(cep)

print(ende['estado'])
print(ende['logradouro'])
print(ende['bairro'])
print(ende['cidade'])
print(ende['cep'])

