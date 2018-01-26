install.packages('data.table')
library(data.table)
library(readxl)

nomes <- function(x) {
  return(paste(x, c("Chave","Texto"), sep=" "))
}

DT.processPeriodo <- function(dt) {
  dt[Periodo == "JAN 2017", Periodo := "201701"]
  dt[Periodo == "FEV 2017", Periodo := "201702"]
  dt[Periodo == "MAR 2017", Periodo := "201703"]
  dt[Periodo == "ABR 2017", Periodo := "201704"]
  dt[Periodo == "MAI 2017", Periodo := "201705"]
  dt[Periodo == "JUN 2017", Periodo := "201706"]
  dt[Periodo == "JUL 2017", Periodo := "201707"]
  dt[Periodo == "AGO 2017", Periodo := "201708"]
  dt[Periodo == "SET 2017", Periodo := "201709"]
  dt[Periodo == "OUT 2017", Periodo := "201710"]
  dt[Periodo == "NOV 2017", Periodo := "201711"]
  dt[Periodo == "DEZ 2017", Periodo := "201712"]
  dt[Periodo == "JAN 2018", Periodo := "201801"]
  dt[Periodo == "FEV 2018", Periodo := "201802"]
  dt[Periodo == "MAR 2018", Periodo := "201803"]
  
#  dt[, Periodo := as.Date(Periodo, "%d/%m/%Y")]
}

DT.splitBWKeyText <- function(dt, campos = c("Periodo"), campos_return, drop_source ) {
  for (c in campos) {
    if (is.null(campos_return) | is.null(campos_return[[c]])) {
      dt[, c(nomes(c)) := lapply(.SD, splitBWKeyText)[[c]], .SDcols = c(c)]
    } else {
      if (length(campos_return[[c]]) == 2)
        dt[, c(campos_return[[c]]) := lapply(.SD, splitBWKeyText)[[c]], .SDcols = c(c)]
      else
        dt[, c(campos_return[[c]]) := lapply(.SD, splitBWKeyText)[[c]][1], .SDcols = c(c)]
    }
    if (is.null(drop_source[[c]]) == FALSE)
      if(drop_source[[c]] == TRUE)
        dt[, (c) := NULL]
  }
}

DT.removeCompoundBW <- function(dt, campos = c("Periodo Chave"), patterns = c("K4")) {
  i <- 1
  for (c in campos) {
    dt[, c(c) := lapply(.SD, removeCompoundBW, patterns[i]), .SDcols = c(c)]
    i <- i + 1
  }
}

splitBWKeyText <- function(x, new_columns = c("Chave", "Texto")) {
  pattern = "^\\["
  return(lapply(tstrsplit(gsub(pattern,"", x), "] ", names = new_columns), trimws))
}

removeCompoundBW <- function(x, compound_pattern = "PB", ifNA = -1) {
  pattern = paste("^", compound_pattern, sep = "")
  result <- trimws(gsub(pattern,"", x))
  return(ifelse(result == "", ifNA, result))
}

isNAorZero <- function(x) is.na(x) | x == 0

DT.getAllNAorZero <- function(dt) {
  return (apply(dt[ , getNumFields(dt), with=FALSE], MARGIN = 2, FUN=isNAorZero))
}

DT.removeRowsAllZero <- function(dt, campos = c("all_numeric")) {
  if((is.null(campos)) | (campos[1] == "all_numeric"))
    campos = getNumFields(dt)
  return (dt[dt[,rowSums(.SD, na.rm = TRUE) != 0, .SDcols = campos]])
}

fillNA <- function(x) {
  temp <- x
  temp[is.na(temp)] <- 0
  return (temp)
}

DT.fillNA <- function(dt, campos = c("all_numeric")) {
  if((is.null(campos)) | (campos[1] == "all_numeric"))
    campos = getNumFields(dt)
  dt[, c(campos) := lapply(.SD, fillNA), .SDcols = campos]
}

getNumFields <- function(dt) {
  tokeepNum <- sapply(head(dt, 1), is.numeric)
  tokeepInt <- sapply(head(dt, 1), is.integer)
  tokeep <- names(dt[1:5 , which(tokeepNum & !tokeepInt), with=FALSE])
  return (tokeep)
}

getCumSum <- function(dt, from = 1, til = 12, campos = c("all_numeric"), inplace = TRUE, camposby = c("Centro de Lucro", "Ano"), new_suffix = "Acumulado", only_grouped_fields = FALSE) {
  return(getGroupedCalculation(dt, calculation = cumsum, from, til, campos, inplace, camposby, new_suffix, only_grouped_fields))
}

getGroupedCalculation <- function(dt, calculation, from = 1, til = 12, campos = c("all_numeric"), inplace = TRUE, camposby = c("Centro de Lucro", "Periodo"), new_suffix = "", only_grouped_fields = FALSE) {
  if((is.null(campos)) | (campos[1] == "all_numeric"))
    campos = getNumFields(dt)
  
  if((is.null(new_suffix)) | (new_suffix == ""))
    new_fields = campos
  else
    new_fields = paste(campos, new_suffix, sep=" - ")
  
  if (inplace) {
    dt[Mes >= from & Mes <= til, c(new_fields) := lapply(.SD, calculation), by = camposby, .SDcols = campos]
    dt[, c(new_fields) := lapply(.SD, fillNA), .SDcols = c(new_fields)]
  } else {
    dt_temp <- dt[Mes >= from & Mes <= til]
    if(only_grouped_fields) {
      dt_temp <- dt_temp[ , lapply(.SD, calculation), by = camposby, .SDcols = campos]
      setnames(dt_temp, c(campos), c(new_fields))
    }
    else {
      dt_temp[, c(new_fields) := lapply(.SD, calculation), by = camposby, .SDcols = campos]
    }
    return(dt_temp)
  }
}

DT.convertFatorEscalonamento <- function(dt, campos = c("all_numeric")) {
  if((is.null(campos)) | (campos[1] == "all_numeric"))
    campos = getNumFields(dt)
  dt[, c(campos) := lapply(.SD, convertFatorEscalonamento), .SDcols = c(campos)]
}

convertFatorEscalonamento <- function(x, fator = 1000000) {
  return(x/fator)
}

DT.processNames <- function(dt) {
  return(processName(names(dt)))
}

processName <- function(n) {
  pattern = "\\í"
#  return(toupper(gsub(pattern,"i", n)))
  return(gsub(pattern,"i", n))
}

calculaMargem <- function(dt, fator = 1000000) {
  dt[, `Margem Operacional Unitária` := fillNA(`Margem Operacional` * fator/`Volume Real`)]
  dt[, `Margem Operacional Unitária - Acumulado` := fillNA(`Margem Operacional - Acumulado` * fator/`Volume Real - Acumulado`)]
}

getPerspective <- function(dt, from = 1, til = 12, calculation = sum, camposby = c("Centro de Lucro", "Centro de Lucro Nome", "Ano", "Mes", "Periodo"), campos = c("all_numeric"), visao = "DEFAULT") {
  if((is.null(campos)) | (campos[1] == "all_numeric"))
    campos = setdiff(getNumFields(dt), c("Margem Operacional Unitária", "Margem Operacional Unitária - Acumulado"))
  dt_temp <- getGroupedCalculation(dt, from = from, til = til, calculation = calculation, camposby = camposby, inplace = FALSE, only_grouped_fields = TRUE, campos = campos)
  calculaMargem(dt_temp)
  dt_temp[,VISAO := visao]
  return(dt_temp)
}

GERA_MARGEM <- read_excel("D:/R Stuff/Gera Margem/input/GERA_MARGEM.xlsx", sheet = "DRE")
GERA_MARGEM <- data.table(GERA_MARGEM)

names(GERA_MARGEM) <- DT.processNames(GERA_MARGEM)

DT.processPeriodo(GERA_MARGEM)

#GERA_MARGEM[, Ano := as.integer(format(GERA_MARGEM$Periodo,"%Y"))]
#GERA_MARGEM[, Mes := as.integer(format(GERA_MARGEM$Periodo,"%m"))]

GERA_MARGEM[, Ano := as.integer(substr(GERA_MARGEM$Periodo, 1, 4))]
GERA_MARGEM[, Mes := as.integer(substr(GERA_MARGEM$Periodo, 5, 6))]

setkey(GERA_MARGEM, "Centro de Lucro", Periodo)

DT.fillNA(GERA_MARGEM)
DT.convertFatorEscalonamento(GERA_MARGEM)
#GERA_MARGEM <- DT.removeRowsAllZero(GERA_MARGEM, campos = getNumFields(GERA_MARGEM))

getCumSum(GERA_MARGEM, from = 2, til = 4)
dt <- getCumSum(GERA_MARGEM, from = 2, til = 4, inplace = FALSE)

#######################################################################################################################################

FAT_LIQ <- read_excel("D:/R Stuff/Gera Margem/input/PBB_PBDPAM001_GERA_FATLIQBR.xls", sheet = "FATLIQ")
FAT_LIQ <- data.table(FAT_LIQ)

names(FAT_LIQ) <- DT.processNames(FAT_LIQ)

DT.fillNA(FAT_LIQ)
DT.convertFatorEscalonamento(FAT_LIQ, campos = setdiff(getNumFields(FAT_LIQ), c("Percentual Volume" ,"Margem Operacional Unitária", "Volume Real", "DRE: Total: Volume")))
FAT_LIQ <- DT.removeRowsAllZero(FAT_LIQ, campos = c("Volume Real", "Vendas liquidas", "Abat. e Descontos", "Encargos Financeiros", "Receita Liquida"))

DT.splitBWKeyText(FAT_LIQ, campos = c("Exercicio/periodo","Centro de lucro","Grupamento Marketing 2","CNPJ Básico"), 
                  campos_return = list("Exercicio/periodo" = c("Periodo"), "Centro de lucro" = c("Centro de Lucro","Centro de Lucro Nome"), "Grupamento Marketing 2" = c("Produto","Produto Nome"), "CNPJ Básico" = c("Cliente","Cliente Nome")),
                  drop_source = list("CNPJ Básico" = TRUE, "Grupamento Marketing 2" = TRUE, "Centro de lucro" = TRUE, "Exercicio/periodo" = TRUE)
                  )
DT.removeCompoundBW(FAT_LIQ, campos = c("Periodo", "Centro de Lucro", "Produto", "Cliente"), patterns = c("K4", "PBACPB", "PB", "PB"))

FAT_LIQ[, Periodo := paste(substr(Periodo, 1, 4), substr(Periodo, 6, 7), sep = "")]

#FAT_LIQ[, Periodo := as.Date(Periodo)]
#FAT_LIQ[, Ano := as.integer(format(FAT_LIQ$Periodo,"%Y"))]
#FAT_LIQ[, Mes := as.integer(format(FAT_LIQ$Periodo,"%m"))]

FAT_LIQ[, Ano := as.integer(substr(FAT_LIQ$Periodo, 1, 4))]
FAT_LIQ[, Mes := as.integer(substr(FAT_LIQ$Periodo, 5, 6))]

setkey(FAT_LIQ, "Centro de Lucro", Periodo, Produto, Cliente)

FAT_LIQ <- getCumSum(FAT_LIQ, from = 1, til = 12, inplace = FALSE, 
                                           camposby = c("Centro de Lucro", "Produto", "Cliente", "Ano"), 
                                           campos = setdiff(getNumFields(FAT_LIQ), c("Percentual Volume" , "Margem Operacional Unitária", "DRE: CPV", "DRE: Despesas Operacionais", "DRE: Total: Volume")), 
                                           only_grouped_fields = FALSE)

FAT_LIQ <- FAT_LIQ[, -c("Percentual Volume" , "Margem Operacional Unitária", "DRE: CPV", "DRE: Despesas Operacionais", "DRE: Total: Volume")]
calculaMargem(FAT_LIQ)
FAT_LIQ[,VISAO := "FATLIQ"]

fatliq.clucro.cliente <- getPerspective(FAT_LIQ, from = 1, til = 12, camposby = c("Centro de Lucro", "Centro de Lucro Nome", "Cliente", "Cliente Nome", "Ano", "Mes", "Periodo"), visao = "CLUCRO_CLIENTE")
fatliq.produto.cliente <- getPerspective(FAT_LIQ, from = 1, til = 12, camposby = c("Produto", "Produto Nome", "Cliente", "Cliente Nome", "Ano", "Mes", "Periodo"), visao = "PRODUTO_CLIENTE")
fatliq.cliente <- getPerspective(FAT_LIQ, from = 1, til = 12, camposby = c("Cliente", "Cliente Nome", "Ano", "Mes", "Periodo"), visao = "CLIENTE")
fatliq.clucro.produto <- getPerspective(FAT_LIQ, from = 1, til = 12, camposby = c("Centro de Lucro", "Centro de Lucro Nome", "Produto", "Produto Nome", "Ano", "Mes", "Periodo"), visao = "CLUCRO_PRODUTO")
fatliq.clucro <- getPerspective(FAT_LIQ, from = 1, til = 12, camposby = c("Centro de Lucro", "Centro de Lucro Nome", "Ano", "Mes", "Periodo"), visao = "CLUCRO")
fatliq.produto <- getPerspective(FAT_LIQ, from = 1, til = 12, camposby = c("Produto", "Produto Nome", "Ano", "Mes", "Periodo"), visao = "PRODUTO")

#fatliq.clucro.produto.cliente: 1
#fatliq.clucro.cliente: 2
#fatliq.produto.cliente: 3
#fatliq.cliente: 4
#fatliq.clucro.produto: 5
#fatliq.clucro: 6
#fatliq.produto: 7

tudo <- rbindlist(list(FAT_LIQ, fatliq.clucro.cliente, fatliq.produto.cliente, fatliq.cliente, fatliq.clucro.produto, fatliq.clucro, fatliq.produto),
                  use.names=TRUE, fill = TRUE, idcol="ID")
