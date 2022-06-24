#########################################################
#                                                       #
#        SCRIPT TO DOWNLOAD INDUSTRIAL EXPORTS          #
#              AND IMPORTS OF BARCELONA                 #
#                                                       #
#########################################################

###-----------------------------------------------------------------------------
# Libraries

library(httr)
library(tidyverse)
library(rvest)
library(RSelenium)
library(xlsx)

setwd(dirname(rstudioapi::getSourceEditorContext()$path)) 

###-----------------------------------------------------------------------------
# Functions

download_monthly_data_full_table <- function() {
  
  remDr$switchToFrame(NULL)
  remDr$switchToFrame(0)
  remDr$findElement(using = "css selector", value = "table.xxdata > tbody > tr:nth-child(2) > td:nth-child(3) > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > a")$clickElement()
  
  remDr$switchToFrame(NULL)
  taula <- remDr$findElement(using = "id", value = "marcoInforme")$getElementAttribute("src")[[1]]
  
  remDr$navigate(taula)
  html <- remDr$getPageSource()[[1]]
  
  remDr$goBack()
  remDr$goBack()
  
  return(html)
}

generator_monthly_table <- function(html, mesos, variable) {
  
  t = 4 + 2 * (mesos-1)
  n = 6 + mesos*2 
  
  months <- read_html(html) %>%
    html_nodes("table") %>%
    .[[8]] %>%
    html_table() %>%
    select(X2) %>%
    slice(4:t) %>%
    t() %>%
    discard(~all(. ==""))
  
  if (variable == 0) {
    
    table <- read_html(html) %>%
      html_nodes("table") %>%
      .[[8]] %>%
      html_table() %>%
      select(X6:str_c("X",n)) %>%
      drop_na() %>%
      slice(-c(1,2)) %>%
      select(1, seq(2, mesos*2 , by=2))
    
  } else {
    table <- read_html(html) %>%
      html_nodes("table") %>%
      .[[8]] %>%
      html_table() %>%
      select(X6:str_c("X",n)) %>%
      drop_na() %>%
      slice(-c(1,2)) %>%
      select(seq(1, mesos*2 + 2, by=2))
  }
  
  
  colnames(table) <- c("Categoria", months)
  
  
  if (dim(table)[1] == 2 ) {
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

criteria_selection <- function() {
  # Navigate to the URL
  remDr$navigate(url)
  
  # Switch the button "Entrar a Datos Comex España"
  remDr$findElement(using = "css selector", value = "ul.row > li:nth-child(1) > ul > li:nth-child(1) > a")$clickElement()
  
  # Products
  Sys.sleep(1.5)
  switching_frames(0)
  remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(4) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()
  
  switching_frames(3)
  
  Sys.sleep(1.5)
  remDr$findElement(using = "css selector", value = "#check1")$clickElement()
  
  # Territory
  Sys.sleep(1.5)
  switching_frames(0)
  remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(6) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(2) ")$clickElement()
  
  remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(6) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()
  
  Sys.sleep(1.5)
  switching_frames(3)
  
  remDr$findElement(using = "css selector", value = "#check10")$clickElement()
  
  # Dates
  Sys.sleep(1.5)
  switching_frames(0)
  
  remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(7) > td > table > tbody > tr > td:nth-child(3) > i")$clickElement()
  
  switching_frames(3)
  
  Sys.sleep(1.5)
  remDr$findElement(using = "css selector", value = "#Table1 > tbody > tr:nth-child(5) > td:nth-child(2) > p > a")$clickElement()
  
  check_box <- year - 1995 + 2  ## Compute where will be the check box
  check_box_pos <- str_c("#check", check_box)
  
  Sys.sleep(1.5)
  remDr$findElement(using = "css selector", value = check_box_pos)$clickElement()
  
  # Units of measure
  Sys.sleep(1.5)
  switching_frames(0)
  
  remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(10) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(4) ")$clickElement()
  
  # See the report
  Sys.sleep(1.5)
  switching_frames(0)
  
  remDr$findElement(using = "id", value = "enviarDatos")$clickElement()
  
}

######################################################################################################################
# What year and what months do you wnat to search?
year <- as.numeric(readline("\nQuin és l'any que es vol extreure? (Introduir l'any en format yyyy): "))
months <- as.numeric(readline("\nNúmero de mesos que es volen extreure: "))

# Url of Data Comex
url <- "https://comercio.serviciosmin.gob.es/Datacomex/"

# Sometimes may not work. In that case, you must change the port parameter, eg.: 5556 to 5557
rD <- rsDriver(browser="firefox", port=5552L, verbose=F)
remDr <- rD[["client"]]

seleccio_criteris()

######################################################################################################################
# Creació de la taula final i l'excel
table_exp <- generador_taula_mensual(descarrega_dades_mensuals_taula_completa(), months, 0) %>%  # 0 per exportacions
  type_convert(locale = locale(decimal_mark = ",", grouping_mark = "."))                                  # 1 per importacions

table_imp <- generador_taula_mensual(descarrega_dades_mensuals_taula_completa(), months, 1) %>%  
  type_convert(locale = locale(decimal_mark = ",", grouping_mark = ".")) 

taula_final <- rbind(table_exp, table_imp) 
taula_final[[1]] <- c("Exportacions", "Importacions")

taula_final <- taula_final %>% gather("Data","Valor", 2:(months+1), factor_key = T) %>% 
  spread(Categoria,Valor) %>% mutate(Saldo = Exportacions - Importacions)


remDr$quit() # Quit session

nom_arxiu <- "Saldo Exportacions_importacions industrials Barcelona.xlsx"
write.xlsx(taula_final, nom_arxiu)




