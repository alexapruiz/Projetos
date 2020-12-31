#!/bin/env python
## Script para Login ig ##
## Desenvolvido por Guilherme Jose Ferreira
## 
## Email: leodmx@brturbo.com.br
# Instruções:
#       - Coloque seu email e senha do IG ali nas linhas 19 e 20.
#       - Execute esse script sempre após estabelecer uma conexão.
#       - Tá pronto para navegar =)
## EDITE AQUI
email = 'usuario@ig.com.br'
senha = 'senha'

## A partir daqui nenhuma alteração é necessária
import httplib, urllib, sys
target_server = 'auth.ig.com.br'
target_script = '/servlets/autentica'
post_vars = urllib.urlencode({
                'metodo' : 'Gateway',
                'username': email,
                'password': senha,
                'cpssg': '',
                'cpkey': '',
                'url': 'http://www.ig.com.br',
                'action': 'login'})
http_headers = {'Content-type': 'application/x-www-form-urlencoded',
                'Accept': 'text/plain'}

print u'Conectando-se a http://' + target_server
conexao = httplib.HTTPConnection(target_server)
print u'Autenticando...'
conexao.request ('POST', target_script, post_vars, http_headers)
resposta = conexao.getresponse()
print u'Reposta do servidor: ', resposta.status, resposta.reason
conexao.close()
sys.exit(0)

### FIM =)
