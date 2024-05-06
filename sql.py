carteiraTelev = "SELECT cd_clien FROM dbo.clientelev where cd_vend = ?"

Telev = "select cd_vend, nome nome_televendas from dbo.vendedor telev where categ = '05'"

cliente = 'select cd_clien,nome_res nome_cliente from dbo.cliente where cd_vend in (?)'