import ConectarSqlite



def acesso(matT,cred):
    sql = ("SELECT acesso FROM tecnicos WHERE mat_gm = '"+matT+"';")
    result = (ConectarSqlite.conectarBd(sql))

    acesso = result[0][0]
    n = acesso.count(cred)

    if n > 0:
        return True
    else:
        return False
