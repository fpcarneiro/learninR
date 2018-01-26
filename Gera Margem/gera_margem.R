install.packages('data.table')
library(data.table)
library(readxl)

tokeepNum <- sapply(head(GERA_MARGEM,1),is.numeric)
tokeepInt <- sapply(head(GERA_MARGEM,1),is.integer)
tokeep <-  which(tokeepNum & !tokeepInt)
teste <- GERA_MARGEM[ , tokeep, with=FALSE]
campos <- names(GERA_MARGEM[ , tokeep, with=FALSE])

teste <- getCumSum2(GERA_MARGEM, 1,5, campos)

getCumSum2 <- function(dt, from, til, campos = c("Resultado_Operacional", "Despesas_Operacionais")) {
  return (dt[Mes >= from & Mes <= til, setNames(lapply(.SD, cumsum), paste(campos, "Acumulado", sep="_")), by = .(Centro_Lucro, Ano), .SDcols = campos])
}


GERA_MARGEM <- read_excel("D:/R Stuff/Gera Margem/input/GERA_MARGEM.xlsx", sheet = "DRE")

# DRE ---------------------------------------------------------------
path <- "input"
GERA_MARGEM <- data.table(GERA_MARGEM)
#ke5z_dre <- fread(file.path(path, "KE5Z_DRE.csv"), encoding = "Latin-1")
#ke5z_dre[,"Exercício/período"]
#str(ke5z_dre)
#head(ke5z_dre)

#setnames(ke5z_dre,
#         c("Exercício/período","Centro de lucro", "Receita Bruta", "Encargos", "Receita Operacional (A)", "CPV (B)", "Resultado Bruto (A - B)", "Despesas Operacionais (C)", "Resultado Operacional (A - B - C)"), 
#         c("Exercicio_Periodo", "Centro_Lucro", "Receita_Bruta", "Encargos", "Receita_Operacional", "CPV", "Resultado_Bruto", "Despesas_Operacionais", "Resultado_Operacional"))

GERA_MARGEM[Periodo == "JAN 2017", Periodo := "01/01/2017"]
GERA_MARGEM[Periodo == "FEV 2017", Periodo := "01/02/2017"]
GERA_MARGEM[Periodo == "MAR 2017", Periodo := "01/03/2017"]
GERA_MARGEM[Periodo == "ABR 2017", Periodo := "01/04/2017"]
GERA_MARGEM[Periodo == "MAI 2017", Periodo := "01/05/2017"]
GERA_MARGEM[Periodo == "JUN 2017", Periodo := "01/06/2017"]
GERA_MARGEM[Periodo == "JUL 2017", Periodo := "01/07/2017"]
GERA_MARGEM[Periodo == "AGO 2017", Periodo := "01/08/2017"]
GERA_MARGEM[Periodo == "SET 2017", Periodo := "01/09/2017"]
GERA_MARGEM[Periodo == "OUT 2017", Periodo := "01/10/2017"]
GERA_MARGEM[Periodo == "NOV 2017", Periodo := "01/11/2017"]
GERA_MARGEM[Periodo == "DEZ 2017", Periodo := "01/12/2017"]
GERA_MARGEM[Periodo == "JAN 2018", Periodo := "01/01/2018"]
GERA_MARGEM[Periodo == "FEV 2018", Periodo := "01/02/2018"]
GERA_MARGEM[Periodo == "MAR 2018", Periodo := "01/03/2018"]


GERA_MARGEM[, Periodo := as.Date(Periodo, "%d/%m/%Y")]

GERA_MARGEM[is.na(Receita_Bruta), Receita_Bruta := 0]
GERA_MARGEM[is.na(Encargos), Encargos := 0]
GERA_MARGEM[is.na(Receita_Operacional), Receita_Operacional := 0]
GERA_MARGEM[is.na(CPV), CPV:= 0]
GERA_MARGEM[is.na(Resultado_Bruto), Resultado_Bruto:= 0]
GERA_MARGEM[is.na(Despesas_Operacionais), Despesas_Operacionais := 0]
GERA_MARGEM[is.na(Resultado_Operacional), Resultado_Operacional := 0]
GERA_MARGEM[is.na(Receita_Intersegmento), Receita_Intersegmento := 0]
GERA_MARGEM[is.na(Custo_Intersegmento), Custo_Intersegmento := 0]
GERA_MARGEM[is.na(Resultado_Bruto_Intersegmento), Resultado_Bruto_Intersegmento := 0]
GERA_MARGEM[is.na(Receita_Exportacao), Receita_Exportacao := 0]

GERA_MARGEM[, Ano := as.integer(format(GERA_MARGEM$Periodo,"%Y"))]
GERA_MARGEM[, Mes := as.integer(format(GERA_MARGEM$Periodo,"%m"))]

setkey(GERA_MARGEM, Centro_Lucro, Ano, Mes, Periodo)

#GERA_MARGEM[, Receita_BrutaCum := cumsum(Receita_Bruta), by = .(Centro_Lucro, Ano)]
#GERA_MARGEM[, EncargosCum := cumsum(Encargos), by = .(Centro_Lucro, Ano)]
#GERA_MARGEM[, Receita_OperacionalCum := cumsum(Receita_Operacional), by = .(Centro_Lucro, Ano)]
#GERA_MARGEM[, CPVCum := cumsum(CPV), by = .(Centro_Lucro, Ano)]
#GERA_MARGEM[, Resultado_BrutoCum := cumsum(Resultado_Bruto), by = .(Centro_Lucro, Ano)]
#GERA_MARGEM[, Despesas_OperacionaisCum := cumsum(Despesas_Operacionais), by = .(Centro_Lucro, Ano)]
#GERA_MARGEM[, Resultado_OperacionalCum := cumsum(Resultado_Operacional), by = .(Centro_Lucro, Ano)]

GERA_MARGEM[Mes >= 2 & Mes <=4, Resultado_OperacionalCum2 := cumsum(Resultado_Operacional), by = .(Centro_Lucro, Ano)]


getCumSum <- function(dt, from, til, campos = c("Resultado_Operacional", "Despesas_Operacionais")) {
#  dt[Mes >= from & Mes <= til, paste(campo, "Cum", sep="_") := cumsum(get(campo)), by = .(Centro_Lucro, Ano)]

  return (dt[Mes >= from & Mes <= til, lapply(.SD, cumsum), by = .(Centro_Lucro, Ano), .SDcols = campos])
  
#  if campo is NULL { # Fazer para todos os numericos
#    tokeep <- which(sapply(head(dt,1),is.numeric))
#    x <- dt[ , tokeep, with=FALSE]
#    for (i in grep("_", names(x), value=TRUE)) {
      #print(i)
#    }
#  }
}



bar[, setNames( sapply(.SD, foo), c("mn", "sd")), by = y, .SDcols="x"]
bar[, as.list(unlist(lapply(.SD, foo))), by = y, .SDcols = c("x", "z")]

GERA_MARGEM[Mes >= 2 & Mes <= 4] <- getCumSum(GERA_MARGEM, 2,4,"Resultado_Operacional")[,.(Resultado_Operacional)]



#teste <- GERA_MARGEM[Mes >= 2 & Mes <=4, .(Resultado_OperacionalCum2 = cumsum(Resultado_Operacional)), by = .(Centro_Lucro, Ano)]

ke5z_dre[Centro_Lucro == "[PBACPBARLUBNOR00] LUBNOR", .(Receita_Bruta, Receita_BrutaCum)]

# EXPORTACAO ----------------------------------------------------------------------------------------------------------------------------------

GERA_MARGEM_KE5Z_EXPORTACAO <- read_excel("input/GERA_MARGEM_KE5Z_DRE.xlsx", sheet = "GERA_MARGEM_KE5Z_EXPORTACAO")
ke5z_exportacao <- data.table(GERA_MARGEM_KE5Z_EXPORTACAO)

ke5z_exportacao[is.na(Saldo), Saldo := 0]
ke5z_exportacao[is.na(Quantidade), Quantidade := 0]

setnames(ke5z_exportacao,
         c("Exercício/período","Centro de lucro", "Unidade de medida","Quantidade","Saldo"), 
         c("Exercicio_Periodo", "Centro_Lucro", "Unidade_Medida","Volume","Vendas_Liquidas"))

ke5z_exportacao[Exercicio_Periodo == "JAN 2017", Exercicio_Periodo := "01/01/2017"]
ke5z_exportacao[Exercicio_Periodo == "FEV 2017", Exercicio_Periodo := "01/02/2017"]
ke5z_exportacao[Exercicio_Periodo == "MAR 2017", Exercicio_Periodo := "01/03/2017"]
ke5z_exportacao[Exercicio_Periodo == "ABR 2017", Exercicio_Periodo := "01/04/2017"]
ke5z_exportacao[Exercicio_Periodo == "MAI 2017", Exercicio_Periodo := "01/05/2017"]

ke5z_exportacao[, Exercicio_Periodo := as.Date(Exercicio_Periodo, "%d/%m/%Y")]

setkey(ke5z_exportacao, Centro_Lucro, Exercicio_Periodo)

# Conversao de Unidades

# INTERSEGMENTO -------------------------------------------------------------------------------------------------------------------------


# FATURAMENTO ---------------------------------------------------------------------------------------------------------------------------
GERA_MARGEM_FATLIQ <- read_excel("input/GERA_MARGEM_KE5Z_DRE.xlsx", sheet = "GERA_MARGEM_FAT_LIQ", 
                                 col_types = c("text", "text", "text", "text", "numeric", "text", "text", "numeric", "numeric", "numeric", "numeric", "numeric"))
fat_liq <- data.table(GERA_MARGEM_FATLIQ)
setnames(fat_liq,
         c("Exercício/período","Centro de lucro", "Grupamento Marketing 2", "X__1", "Nº cliente", "X__2", "UM básica", "Volume Real", "Vendas líquidas","Abat. e Descontos", "Encargos Financeiros", "Receita Unitária"), 
         c("Exercicio_Periodo", "Centro_Lucro", "Produto", "Produto_Nome", "Cliente", "Cliente_Nome", "Unidade_Medida", "Volume", "Vendas_Liquidas","Abat_Descontos", "Encargos_Financeiros", "Receita_Unitaria"))

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