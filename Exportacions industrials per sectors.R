#########################################################
#                                                       #
#        SCRIPT TO DOWNLOAD INDUSTRIAL EXPORTS          #
#          OF BARCELONA ACCORDING TO SECTOR             #
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
  remDr$findElement(using = "css selector", value = "div.a32 > a")$clickElement()
  
  remDr$switchToFrame(NULL)
  taula <- remDr$findElement(using = "id", value = "marcoInforme")$getElementAttribute("src")[[1]]
  
  remDr$navigate(taula)
  html <- remDr$getPageSource()[[1]]
  
  remDr$goBack()
  remDr$goBack()
  
  return(html)
}

access_semimanufactures <- function() {
  
  remDr$switchToFrame(NULL)
  remDr$switchToFrame(0)
  remDr$findElement(using = "css selector", value = "table.xxdata > tbody > tr:nth-child(7) > td:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > a")$clickElement()
  
}

download_monthly_data_semimanufactures <- function() {
  
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

generator_monthly_table <- function(html, mesos) {
  
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

criteria_selection <- function() {
  # Navigate to the URL
  remDr$navigate(url)
  
  # Switch the button "Entrar a Datos Comex España"
  remDr$findElement(using = "css selector", value = "ul.row > li:nth-child(1) > ul > li:nth-child(1) > a")$clickElement()
  
  # Select parameters
  switching_frames(0)
  
  # Exports
  Sys.sleep(1.5)
  remDr$findElement(using = "css selector", value = "td.panelClass2 > table > tbody > tr:nth-child(3) > td > div > table > tbody > tr >td:nth-child(2) > select > option:nth-child(2) ")$clickElement()
  
  # Products
  Sys.sleep(1.5)
  remDr$findElement(using = "css selector", value = "#FormInforme > table > tbody > tr > td >table >tbody > tr:nth-child(4) > td > div > table > tbody > tr > td:nth-child(3) > i")$clickElement()
  
  switching_frames(3)
  
  Sys.sleep(1.5)
  remDr$findElement(using = "css selector", value = "#check1")$clickElement()
  
  # Territory
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
  
  check_box <- year - 1995 + 2  ## Calculem en quina posició es troba la casella
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
  
  # Full table
  Sys.sleep(1.5)
  switching_frames(0)
  remDr$findElement(using = "css selector", value = "div.a63 > a")$clickElement()
}

######################################################################################################################
# What year and what months do you wnat to search?
year <- as.numeric(readline("Quin és l'últim any disponible a la web? (Introduir l'any en format yyyy): "))
months <- as.numeric(readline("Número de mesos que es volen extreure: "))

# Url of Data Comex
url <- "https://comercio.serviciosmin.gob.es/Datacomex/"

# Sometimes may not work. In that case, you must change the port parameter, eg.: 5556 to 5557
rD <- rsDriver(browser="firefox", port=5558L, verbose=F)
remDr <- rD[["client"]]

criteria_selection()

######################################################################################################################
# Creació de la taula final i l'excel

table <- generator_monthly_table(download_monthly_data_full_table(), months)

access_semimanufactures()
taula_semimanufactures <- generator_monthly_table(download_monthly_data_semimanufactures(), months) %>% slice(3)
remDr$goBack()

remDr$quit() # Quit session

taula_final <- rbind(table, taula_semimanufactures) %>% 
  type_convert(locale = locale(decimal_mark = ",", grouping_mark = "."))

nom_arxiu <- "Exportacions industrials per sectors.xlsx"
write.xlsx(taula_final, nom_arxiu)


  


