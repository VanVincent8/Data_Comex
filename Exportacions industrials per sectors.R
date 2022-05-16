if (!require(httr)) install.packages("httr"); library(httr)
if (!require(tidyverse)) install.packages("tidyverse"); library(tidyverse)
if (!require(rvest)) install.packages("rvest"); library(rvest)
if (!require(Selenium)) install.packages("Selenium"); library(RSelenium)
if (!require(xlsx)) install.packages("xlsx"); library(xlsx)

rm(list = ls())
setwd("K:/QUOTA/OMD/ANALISI/0_USUARIS/VICENT/EXPORTACIONS/DataComex") 

######################################################################################################################
# Funcions

descarrega_dades_mensuals_taula_completa <- function(year) {
  
  path <- str_c("#table-data > tbody > tr:nth-child(2) > td:nth-child(", year, ") > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > a")
  
  remDr$switchToFrame(NULL)
  remDr$switchToFrame(0)
  remDr$findElement(using = "css selector", value = path)$clickElement()
  
  remDr$switchToFrame(NULL)
  taula <- remDr$findElement(using = "id", value = "marcoInforme")$getElementAttribute("src")[[1]]
  
  remDr$navigate(taula)
  html <- remDr$getPageSource()[[1]]
  
  remDr$goBack()
  remDr$goBack()
  
  return(html)
}

accedir_semimanufactures <- function() {
  
  remDr$switchToFrame(NULL)
  remDr$switchToFrame(0)
  remDr$findElement(using = "css selector", value = "#table-data > tbody > tr:nth-child(7) > td:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > a")$clickElement()
  
}

descarrega_dades_mensuals_semimanufactures <- function(year) {
  
  path <- str_c("#table-data > tbody > tr:nth-child(2) > td:nth-child(", year, ") > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > a")
  
  remDr$switchToFrame(NULL)
  remDr$switchToFrame(0)
  remDr$findElement(using = "css selector", value = path)$clickElement()
  
  remDr$switchToFrame(NULL)
  taula <- remDr$findElement(using = "id", value = "marcoInforme")$getElementAttribute("src")[[1]]
  
  remDr$navigate(taula)
  html <- remDr$getPageSource()[[1]]
  
  remDr$goBack()
  remDr$goBack()
  
  return(html)
}

generador_taula_mensual <- function(html, mesos) {
  
  t = 4 + 2 * (mesos-1)
  n = 6 + mesos
  
  months <- read_html(html) %>%
    html_nodes("table") %>%
    .[[8]] %>%
    html_table() %>%
    select(X2) %>%
    slice(4:t) %>%
    t() %>%
    discard(~all(. ==""))
  
  table <- read_html(html) %>%
    html_nodes("table") %>%
    .[[8]] %>%
    html_table() %>%
    select(X6:str_c("X",n)) %>%
    drop_na() %>%
    slice(-c(1,2))
  
  colnames(table) <- c("Categoria", months)
  
  
  if (dim(table)[1] %in% c(5,10) ) {
    table <- table %>% slice(-1)
  }
  
  table
}

switching_frames <- function(n) {
  if (n == 0) {
    
    remDr$switchToFrame(NULL)
    remDr$switchToFrame(n)
    
  } else {
    
    remDr$switchToFrame(n)
  }
}

######################################################################################################################
# Pàgina principal del Data Comex
url <- "https://comercio.serviciosmin.gob.es/Datacomex/"

# A vegades pot no funcionar. En aquest cas canviar el nombre del port, ex: passar de 5556 a 5557
rD <- rsDriver(browser="firefox", port=5558L, verbose=F)
remDr <- rD[["client"]]

# Naveguem a la url
remDr$navigate(url)

# Premem en el botó d'entrar a Datos Comex España
remDr$findElement(using = "css selector", value = "ul.row > li:nth-child(1) > ul > li:nth-child(1) > a")$clickElement()

# Seleccionem els criteris
switching_frames(0)

# Exportacions
remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(3) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(2) ")$clickElement()

# Productes
remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(4) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()

switching_frames(3)

remDr$findElement(using = "css selector", value = "#check1")$clickElement()

# Territorial
switching_frames(0)

remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(6) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(2) ")$clickElement()

remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(6) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()

switching_frames(3)

remDr$findElement(using = "css selector", value = "#check10")$clickElement()

# Dates
switching_frames(0)

remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(7) > td > table > tbody > tr > td:nth-child(3) > i")$clickElement()

switching_frames(3)

remDr$findElement(using = "css selector", value = "#Table1 > tbody > tr:nth-child(5) > td:nth-child(2) > p > a")$clickElement()

for (i in 23:29){  # en cas d'afegir-se un any nou (ex: 2023) s'haurà d'incrementar el rang de 29 --> 30
  year <- str_c("#check", i)
  remDr$findElement(using = "css selector", value = year)$clickElement()
}

# Unitats de mesura
switching_frames(0)

remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(10) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(4) ")$clickElement()

# Veim l'informe
switching_frames(0)

remDr$findElement(using = "id", value = "enviarDatos")$clickElement()

# Taula completa
switching_frames(0)
remDr$findElement(using = "css selector", value = "div.a63 > a")$clickElement()


######################################################################################################################
# Creació de la taula final i l'excel

mesos <- 2
any <- 2022

table <- generador_taula_mensual(descarrega_dades_mensuals_taula_completa(any - 2012), mesos)

accedir_semimanufactures()
taula_semimanufactures <- generador_taula_mensual(descarrega_dades_mensuals_semimanufactures(any - 2012), mesos) %>% slice(3)
remDr$goBack()


remDr$quit() # Tanques sessió

taula_final <- rbind(table, taula_semimanufactures) %>% type_convert(locale = locale(decimal_mark = ",", grouping_mark = "."))

nom_arxiu <- "Exportacions industrials per sectors.xlsx"
write.xlsx(taula_final, nom_arxiu)


  


