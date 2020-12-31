import pyad.adquery

q = pyad.adquery.ADQuery()

q.execute_query(attributes= ["distinguishedName", "descrição"],
    where_clause = "objectClass = '*'",
    base_dn = "OU=caixa,DC=corp.caixa.gov.br,DC=gov"
)

for row in q.get_results():
    print(row)