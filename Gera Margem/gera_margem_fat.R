# FATURAMENTO ---------------------------------------------------------------------------------------------------------------------------
GERA_MARGEM_FATLIQ <- read_excel("D:/R Stuff/Gera Margem/input/FATLIQ.xlsx")
fat_liq <- data.table(GERA_MARGEM_FATLIQ)
setnames(fat_liq,
         c("Exercício/período","Centro de lucro", "X__1", "Grupamento Marketing 2", "X__2", "Nº cliente", "X__3","Matriz","X__4", "Volume Real", "Vendas líquidas","Abat. e Descontos", "Encargos Financeiros", "Receita Unitária"), 
         c("Exercicio_Periodo", "Centro_Lucro", "Centro_Lucro_Nome", "Produto", "Produto_Nome", "Cliente", "Cliente_Nome","Matriz","Matriz Nome", "Volume", "Vendas_Liquidas","Abat_Descontos", "Encargos_Financeiros", "Receita_Unitaria"))

fat_liq[Exercicio_Periodo == "JAN 2017", Exercicio_Periodo := "01/01/2017"]
fat_liq[Exercicio_Periodo == "FEV 2017", Exercicio_Periodo := "01/02/2017"]
fat_liq[Exercicio_Periodo == "MAR 2017", Exercicio_Periodo := "01/03/2017"]
fat_liq[Exercicio_Periodo == "ABR 2017", Exercicio_Periodo := "01/04/2017"]
fat_liq[Exercicio_Periodo == "MAI 2017", Exercicio_Periodo := "01/05/2017"]

fat_liq[, Exercicio_Periodo := as.Date(Exercicio_Periodo, "%d/%m/%Y")]
setkey(fat_liq, Centro_Lucro, Exercicio_Periodo, Produto, Cliente)

fat_liq[, Vendas_LiquidasCum := cumsum(Vendas_Liquidas), by = .(Centro_Lucro, Produto, Cliente)]
fat_liq[, VolumeCum := cumsum(Volume), by = .(Centro_Lucro, Produto, Cliente)]
fat_liq[, Abat_DescontosCum := cumsum(Abat_Descontos), by = .(Centro_Lucro, Produto, Cliente)]
fat_liq[, Encargos_FinanceirosCum := cumsum(Encargos_Financeiros), by = .(Centro_Lucro, Produto, Cliente)]

str(fat_liq)

rm(GERA_MARGEM_FATLIQ, GERA_MARGEM_KE5Z_DRE, GERA_MARGEM_KE5Z_EXPORTACAO)