if (!require(httr)) install.packages("httr"); library(httr)
if (!require(tidyverse)) install.packages("tidyverse"); library(tidyverse)
if (!require(rvest)) install.packages("rvest"); library(rvest)
if (!require(RSelenium)) install.packages("RSelenium"); library(RSelenium)
if (!require(xlsx)) install.packages("xlsx"); library(xlsx)

rm(list = ls())
setwd(dirname(rstudioapi::getSourceEditorContext()$path)) 

######################################################################################################################
# Funcions

descarrega_dades_mensuals_taula_completa <- function() {
  
  remDr$switchToFrame(NULL)
  remDr$switchToFrame(0)
  remDr$findElement(using = "css selector", value = "table.xxdata > tbody > tr:nth-child(2) > td:nth-child(4) > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > a")$clickElement()
  
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
  
  colnames(table) <- c("Àrea geogràfica", months)
  
  
  if (dim(table)[1] == 16 ) {
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
# Quin any i mesos es cerquen?
year <- as.numeric(readline("Quin és l'últim any disponible a la web? (Introduir l'any en format yyyy): "))
months <- as.numeric(readline("Número de mesos que es volen extreure: "))

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

# Exportacions
switching_frames(0)
remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(3) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(2) ")$clickElement()

# Territorial
switching_frames(0)

remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(6) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(2) ")$clickElement()

remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(6) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()

switching_frames(3)

remDr$findElement(using = "css selector", value = "#check10")$clickElement()

# País
switching_frames(0)

remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(5) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(1) ")$clickElement()

remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(5) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()

switching_frames(3)

#### <<< Total món >>>
remDr$findElement(using = "css selector", value = "#check2")$clickElement()

#### <<< Amèrica >>>
for (i in 6:9){  
  check_box <- str_c("#check", i)
  remDr$findElement(using = "css selector", value = year)$clickElement()
}

#### <<< Àsia >>>
remDr$findElement(using = "css selector", value = "#check11")$clickElement()

remDr$findElement(using = "css selector", value = "#raiz > table > tbody > tr:nth-child(12) > td > i:nth-child(2)")$clickElement()

remDr$findElement(using = "css selector", value = "#check56")$clickElement()
remDr$findElement(using = "css selector", value = "#check59")$clickElement()

#### <<< Europa >>>
remDr$findElement(using = "css selector", value = "#check67")$clickElement()

remDr$findElement(using = "css selector", value = "#raiz > table > tbody > tr:nth-child(78) > td > i:nth-child(2)")$clickElement()

remDr$findElement(using = "css selector", value = "#check78")$clickElement()

remDr$findElement(using = "css selector", value = "#check298")$clickElement()

remDr$findElement(using = "css selector", value = "#raiz > table > tbody > tr:nth-child(248) > td > i:nth-child(2)")$clickElement()

remDr$findElement(using = "css selector", value = "#check299")$clickElement()
remDr$findElement(using = "css selector", value = "#check302")$clickElement()
remDr$findElement(using = "css selector", value = "#check303")$clickElement()
remDr$findElement(using = "css selector", value = "#check307")$clickElement()

# Dates
switching_frames(0)

remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(7) > td > table > tbody > tr > td:nth-child(3) > i")$clickElement()

switching_frames(3)

remDr$findElement(using = "css selector", value = "#Table1 > tbody > tr:nth-child(5) > td:nth-child(2) > p > a")$clickElement()

for (i in (year-1995+1):(year-1995+2)){  
  check_box <- str_c("#check", i)
  remDr$findElement(using = "css selector", value = check_box)$clickElement()
}

# Unitats de mesura
switching_frames(0)

remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(10) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(4) ")$clickElement()

# Informe
switching_frames(0)
remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(14) > td > table > tbody > tr >td:nth-child(2) > select > option:nth-child(4) ")$clickElement()

# Veim l'informe
switching_frames(0)

remDr$findElement(using = "id", value = "enviarDatos")$clickElement()

######################################################################################################################
# Creació de la taula final i l'excel

taula <- generador_taula_mensual(descarrega_dades_mensuals_taula_completa(), months) %>% 
  type_convert(locale = locale(decimal_mark = ",", grouping_mark = "."))

remDr$quit() # Tancar sessió

nom_arxiu <- "Exportacions industrials per àrees geogràfiques.xlsx"
write.xlsx(taula, nom_arxiu)




