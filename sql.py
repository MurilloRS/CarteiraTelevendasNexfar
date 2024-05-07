# Consulta SQL para verificar a existência da combinação (cd_clien, cd_vend)
query = "SELECT * FROM dbo.clientelev WHERE cd_clien = ? AND cd_vend = ?"
insert = "INSERT INTO dbo.clientelev (cd_clien,cd_vend) VALUES (?, ?);"
delete = "DELETE FROM dbo.clientelev WHERE cd_clien = ? AND cd_vend = ?"
cliente_query = "SELECT * FROM dbo.cliente WHERE cd_clien = ?"
vendedor_query = "SELECT * FROM dbo.vendedor where cd_vend = ?"
carteira = """select
                    c.cd_clien,
                    c.cd_vend,
                    '' as '           ',
                    cl.cd_clien codCliente,
                    cl.nome cliente,
                    ec.estado,
                    vd.cd_vend codVendedor,
                    vd.nome vendedor,
                    eq.descricao equipe
                from
                    clientelev c
                    left join cliente cl on cl.cd_clien = c.cd_clien
                    left join vendedor vd on vd.cd_vend = c.cd_vend
                    left join end_cli ec on ec.cd_clien = cl.cd_clien and ec.tp_end = 'FA'
                    left join equipe eq on eq.EquipeId = vd.EquipeID"""