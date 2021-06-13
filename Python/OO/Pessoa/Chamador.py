from pessoa_fisica import PessoaFisica
from pessoa_juridica import PessoaJuridica

PF1 = PessoaFisica('199.954.008-55', nome='Alex Ruiz', idade=42)
PF2 = PessoaFisica('111.222.333-44', nome='Vitor Panin Ruiz', idade=13)

PF1.setCPF = '123.456.789-00'
print(PF1.getCPF())
print(PF2.getCPF())

#print(PF1.getNome())
#print(PF1.getIdade())
#print('')
#print(PF2.getCPF())
#print(PF2.getNome())
#print(PF2.getIdade())

#PJ1 = PessoaJuridica('64.614.527/0001-99', nome='Empresa X', idade=22)
#print(PJ1.getCNPJ())
#print(PJ1.getNome())
#print(PJ1.getIdade())