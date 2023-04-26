# Script voor het aanmaken van excel data voor de rapportage
# Versie 1.0 gebaseerd op versie 5.0 van JV

# Auteur(s)
# Frouwke Veenstra - fveenstra@ggdghor.nl
# Auteurs basis JV
# Sander Vermeulen - s.vermeulen@vrln.nl
# Sjanne van der Stappen - s.van.der.stappen@vrln.nl

# 0. Voorbereiding --------------------------------------------------------

# Leegmaken environment
rm(list=ls())

# 1. Laden benodigde packages --------------------------------------------

# Packages moeten eenmalig worden geinstalleerd.
# Verwijder de # aan het begin van de volgende regel om de code te runnen en de benodigde packages te installeren
# install.packages(c('tidyverse', 'readxl', 'openxlsx', 'haven', 'labelled', 'fastDummies'))

# Laden van packages
library(tidyverse) # tidyverse package voor databewerking 
library(readxl) # voor read_excel()
library(openxlsx) # voor write.xlsx()
library(haven) # voor read_spss()
library(labelled) # voor to_character()
library(fastDummies) # voor dummy_cols()


# 2. Definieren minimum aantallen -----------------------------------------

# Minimum aantal respondenten per variabele
minimum_N <- 50

# Minimum aantal respondenten per cel/percentage
minimum_N1 <- 0

# 3. Definieren regionaam en -code ----------------------------------------

# Pas onderstaande naam en code aan naar de naam en code van je eigen regio.
regionaam <- 'GGD Limburg-Noord'
regiocode <- 23


# 4. Indicatorenoverzicht en data inladen ---------------------------------

# Verzet working directory naar de map  met het Indicatorenoverzicht, de conceptrapportage
# en dit script. De working directory is de map van waaruit R werkt. Hieronder zijn twee manieren 
# genoemd om dit te doen: via code of via het menu in R-studio.
# Via code: 
setwd("padnaam") # vul voor padnaam de padnaam van de map met data in.
# LET OP! Gebruik voor de pad naam de forward slash (/) en niet een backward slash (\) zoals Microsoft in de padnaam heeft staan.

# Via het menu in R-studio:
# 1. Klik in de balk bovenin op Session > Set Working Directory > Choose directory...
# 2. Selecteer in de pop-up die opent de map waarin de data staat en klik op Open
# In Console (beneden in het scherm) wordt nu een regel code toegevoegd. Kopieer deze 
# en plak hieronder. Wanneer je de volgende keer dit script runt, kun je meteen de code 
# hieronder runnen en hoef je de map niet opnieuw te selecteren.
setwd("C:") # Vul voor padnaam de padnaam van de map met het Indicatorenoverzicht in.

# Indicatorenoverzicht laden
ind.overzicht <- read_excel('Indicatoren overzicht.xlsx', sheet = 'indicatoren')

# Indicatorenoverzicht trends laden en regel aanmaken per indicator in de kolom 'niveau'
trends <- read_excel('Indicatoren overzicht.xlsx', sheet = 'indicatoren trends') %>%
  mutate(niveau = strsplit(niveau, ", ")) %>%
  unnest(niveau)

# Zet working directory naar de map met data.
# Via code: 
setwd("padnaam") # vul voor padnaam de padnaam van de map met data in.
# LET OP! Gebruik voor de pad naam de forward slash (/) en niet een backward slash (\) zoals Microsoft in de padnaam heeft staan.

# Inladen Totaalbestand:
# Als je in de vorige stap een working directory hebt gedefinieerd, hoef je in deze stap
# alleen de bestandsnaam op te geven (inclusief de bestandsextensie: .sav)
# Als je geen working directory hebt gedefinieerd, geef dan de volledige padnaam op.
# Naast het opgeven van de bestandsnaam (inclusief padnaam en bestandsextensie) is het ook mogelijk
# om met behulp van de functie file.choose() het bestand te selecteren in de verkenner.
# Werk je liever met de padnaam vervang dan file.choose() door 'padnaam/Data 2022.sav'

# Alle spss data wordt ingeladen met behulp van read_spss, waarbij het file argument wordt
# gedefinieerd aan de hand van de padnaam van het bestand. Daarnaast wordt met behulp van
# het argument col_select alleen de noodzakelijke indicatoren ingeladen. Dit wordt gebaseerd
# op het indicatorenoverzicht. De niveau variabelen 'nederland', 'regio' worden later in het
# script aangemaakt. Wil je zelf extra niveauvariabelen toevoegen dan moet je die toevoegen
# bij het col_select argument. Wil je bijvoorbeeld een variabele 'Subregio' toevoegen dan 
# kun je die hieronder achter 'Gemeentecode' zetten gevolgd door een komma (,).

# Data 2022 laden
data <- read_spss(file = file.choose(), 
                  col_select = c('MIREB201', 'Gemeentecode',
                                 ind.overzicht$indicator[ind.overzicht$dichotomiseren == 0],
                                 ind.overzicht$indicator[ind.overzicht$dichotomiseren == 1] %>% str_match('.+(?=_[0-9]+$)|.+') %>% as.vector() %>% unique(),
                                 ind.overzicht$uitsplitsing %>% .[is.na(.) == F] %>% str_split(', ') %>% unlist() %>% unique(),
                                 ind.overzicht$weegfactor %>% unlist() %>% unique()))

# Trenddata 2020 laden
# Werk je liever met de padnaam vervang dan file.choose() door 'padnaam/Trendbestand 2020 met indicatoren 2022.sav'
data2020 <- read_spss(file = file.choose(), 
                      col_select = c('MIREB201', 'Gemeentecode',
                                     trends$indicator[trends$dichotomiseren == 0 & !is.na(trends$weegfactor2020)] %>% unique(),
                                     trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2020)] %>% str_match('.+(?=_[0-9]+$)|.+') %>% as.vector() %>% unique(),
                                     trends$weegfactor2020 %>% unique() %>% setdiff(NA)))

# Trenddata 2016 laden
# Werk je liever met de padnaam vervang dan file.choose() door 'padnaam/Trendbestand 2016 met indicatoren 2022.sav'
data2016 <- read_spss(file = file.choose(),
                      col_select = c('MIREB201', 'Gemeentecode',
                                     trends$indicator[trends$dichotomiseren == 0 & !is.na(trends$weegfactor2016)] %>% unique(),
                                     trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2016)] %>% str_match('.+(?=_[0-9]+$)|.+') %>% as.vector() %>% unique(),
                                     trends$weegfactor2016 %>% unique() %>% setdiff(NA)))

# Trenddata 2012 laden
# Werk je liever met de padnaam vervang dan file.choose() door 'padnaam/Trendbestand 2012 met indicatoren 2022.sav'
data2012 <- read_spss(file = file.choose(),
                      col_select = c('MIREB201', 'Gemeentecode',
                                     trends$indicator[trends$dichotomiseren == 0 & !is.na(trends$weegfactor2012)] %>% unique(),
                                     trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2012)] %>% str_match('.+(?=_[0-9]+$)|.+') %>% as.vector() %>% unique(),
                                     trends$weegfactor2012 %>% unique() %>% setdiff(NA)))

# Werk je vanuit een trendbestand met data voor alle jaren dan kun je de data uit je totaalbestand filteren.
# Pas hierbij de indicator jaar_variabele aan naar de kolom die het jaartal bevat.
# data2020 <- data %>% filter(jaar_variabele == 2020)
# data2016 <- data %>% filter(jaar_variabele == 2016)
# data2012 <- data %>% filter(jaar_variabele == 2012)
# data <- data %>% filter(jaar_variabele == 2022)

# Inladen responsbestand
responsbestand <- read_spss(file.choose()) # <- read_spss('Padnaam/Datarespons.sav')

# 5. Databewerkingen uitvoeren --------------------------------------------
data <-  data %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))


# 6. Dichotomiseren -------------------------------------------------------

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(ind.overzicht$dichotomiseren == 1)){
  data <- dummy_cols(data, 
                     select_columns = unique(str_extract(ind.overzicht$indicator[ind.overzicht$dichotomiseren == 1], '.+(?=_[0-9]+$)')),
                     ignore_na = T)
}

# 7. Hercoderen -----------------------------------------------------------

# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data <- data %>%
  mutate_at(c(data %>%
                select(ind.overzicht$indicator) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                ind.overzicht$indicator[.]), list(~recode(., `8`= 0)))


# 8. Tabel aanmaken -------------------------------------------------------

# Het aanmaken van een tabel met elke opgegeven combinatie van indicator, uitsplitsing en niveau.
# Deze tabel vormt de input op basis waarvan de gemiddelden worden berekend.
input <- ind.overzicht %>%
  mutate(uitsplitsing = ifelse(is.na(uitsplitsing), 'totaal', paste('totaal,', uitsplitsing))) %>%
  mutate(uitsplitsing = strsplit(uitsplitsing, ", ")) %>%
  unnest(uitsplitsing) %>%
  mutate(niveau = strsplit(niveau, ", ")) %>%
  unnest(niveau) %>%
  select(-opmerkingen)


# 9. Gemiddeldes berekenen -----------------------------------------------

# Het aanmaken van een functie waarmee gemiddelden kunnen worden berekend
compute_mean <- function(data, indicator, uitsplitsing, niveau, weegfactor, weighted = T){
  
  groepering <- setdiff(c(uitsplitsing, niveau), NA)
  
  data %>%
    select(all_of(c(indicator, groepering, weegfactor))) %>%
    group_by(across(all_of(groepering))) %>%
    summarise_at(indicator, list(mean = if(weighted == F) ~mean(., na.rm = T) else( ~weighted.mean(x = ., w = .data[[weegfactor]], na.rm = T)), # berekenen gewogen gemiddelde
                                 N = ~sum(!is.na(.)), # berekenen van de N
                                 N0 = ~length(which(.==0)), # berekenen van de N0 
                                 N1 = ~length(which(.==1)))) %>% # berekenen van de N1
    drop_na(any_of(groepering)) %>% # Verwijderen van NA values voor groepering (op basis van uitsplitsing en niveau)
    {if(is.na(uitsplitsing)) rename(., niveau = 1) else rename(., uitsplitsing = 1, niveau = 2)} %>%
    {if(is.na(uitsplitsing)) mutate(., uitsplitsing = 'totaal') else mutate(., uitsplitsing = to_character(uitsplitsing))} %>%
    mutate(mean = ifelse(N < minimum_N | (N0 < minimum_N1) | (N1 < minimum_N1), NA, mean),
           indicator = indicator) %>%
    select(uitsplitsing, niveau, mean) %>%
    pivot_wider(names_from = uitsplitsing,
                names_prefix = paste0(indicator, '_'),
                values_from = mean,
                names_sep = '_') %>%
    pivot_longer(cols = -niveau, values_drop_na = F)

}

# De compute_mean functie kan ook worden gebruikt om cijfers voor een enkele indicator te berekenen, bijvoorbeeld:
# compute_mean(data = data,
#              indicator = 'KLGGA208',
#              uitsplitsing = 'AGGSA202',
#              niveau = 'nederland',
#              weegfactor = 'ewCBSGGD')


# Cijfers voor alle combinaties van indicatoren, uitsplitsingen en niveaus uit het indicatoren overzicht berekenen
cijfers <- input %>%
  select(indicator, uitsplitsing, niveau, weegfactor) %>% # selecteren van relevante kolommen
  pmap(compute_mean, data = data) %>% # compute_mean functie toepassen op het input object
  bind_rows() # output combineren tot een dataframe

# Namen van uitsplitsingen verkorten en afronden op 6 decimalen
output <- cijfers %>%
  mutate(name = str_replace(name, '[\\.][a-z]$', ''), # hernoemen van uitsplitsingen
         name = str_replace(name, 'Man', 'm'),
         name = str_replace(name, 'Vrouw', 'v'),
         name = str_replace(name, '18-34 jaar', '1834'),
         name = str_replace(name, '35-49 jaar', '3549'),
         name = str_replace(name, '50-64 jaar', '5064'),
         name = str_replace(name, '65-74 jaar', '6574'),
         name = str_replace(name, '75 en ouder', '75+'),
         name = str_replace(name, '18-64 jaar', '1864'),
         name = str_replace(name, '65 jaar en ouder', '65+'),
         name = str_replace(name, 'Laag \\(LO, MAVO, LBO\\)', 'Laag'), # \\ is nodig om aan te geven dat '(' en ')' als tekst moet worden geevalueerd
         name = str_replace(name, 'Midden \\(HAVO, VWO, MBO\\)', 'Midden'),
         name = str_replace(name, 'Hoog \\(HBO, WO\\)', 'Hoog'),
         name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) # output afronden op 6 decimalen (om de output leesbaarder te maken, maar geen dubbele afrondingsfouten te introduceren)


# 10. Trendcijfers --------------------------------------------------------

# Trendcijfers 2022
trends2022 <- trends %>%
  filter(!is.na(trends$weegfactor2022)) %>%
  select(indicator, niveau, weegfactor = weegfactor2022) %>%
  pmap(compute_mean, data = data, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2022')))

# Trendcijfers 2020
# Aanmaken van niveauvariabelen
data2020 <-  data2020 %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(trends$dichotomiseren == 1 & !is.na(trends$weegfactor2020))){
  data2020 <- dummy_cols(data2020,
                         select_columns = unique(str_extract(trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2020)], '.+(?=_[0-9]+$)')),
                         ignore_na = T)
}


# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data2020 <- data2020 %>%
  mutate_at(c(data2020 %>%
                select(unique(trends$indicator[!is.na(trends$weegfactor2020)])) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                unique(trends$indicator[!is.na(trends$weegfactor2020)])[.]), 
            list(~recode(., `8`= 0)))

# Berekenen trendcijfers
trends2020 <- trends %>%
  filter(!is.na(trends$weegfactor2020)) %>%
  select(indicator, niveau, weegfactor = weegfactor2020) %>%
  pmap(compute_mean, data = data2020, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2020')))

# Trendcijfers 2016
# Aanmaken van niveauvariabelen
data2016 <-  data2016 %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(trends$dichotomiseren == 1 & !is.na(trends$weegfactor2016))){
  data2016 <- dummy_cols(data2016,
                         select_columns = unique(str_extract(trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2016)], '.+(?=_[0-9]+$)')),
                         ignore_na = T)
}


# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data2016 <- data2016 %>%
  mutate_at(c(data2016 %>%
                select(unique(trends$indicator[!is.na(trends$weegfactor2016)])) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                unique(trends$indicator[!is.na(trends$weegfactor2016)])[.]), 
            list(~recode(., `8`= 0)))

# Berekenen trendcijfers
trends2016 <- trends %>%
  filter(!is.na(trends$weegfactor2016)) %>%
  select(indicator, niveau, weegfactor = weegfactor2016) %>%
  pmap(compute_mean, data = data2016, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2016')))

# Trendcijfers 2012
# Aanmaken van niveauvariabelen
data2012 <-  data2012 %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(trends$dichotomiseren == 1 & !is.na(trends$weegfactor2012))){
  data2012 <- dummy_cols(data2012,
                         select_columns = unique(str_extract(trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2012)], '.+(?=_[0-9]+$)')),
                         ignore_na = T)
}

# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data2012 <- data2012 %>%
  mutate_at(c(data2012 %>%
                select(unique(trends$indicator[!is.na(trends$weegfactor2012)])) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                unique(trends$indicator[!is.na(trends$weegfactor2012)])[.]), 
            list(~recode(., `8`= 0)))

# Berekenen trendcijfers
trends2012 <- trends %>%
  filter(!is.na(trends$weegfactor2012)) %>%
  select(indicator, niveau, weegfactor = weegfactor2012) %>%
  pmap(compute_mean, data = data2012, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2012')))


# Samenvoegen van alle trenddata 
trends_totaal <- trends2022 %>%
  left_join(trends2020, by = 'niveau') %>%
  left_join(trends2016, by = 'niveau')%>% 
  left_join(trends2012, by = 'niveau')%>% 
  select(niveau, sort(colnames(.))) %>% # sorteren van kolommen op naam
  mutate_at(vars(-niveau), round, 6) # gemiddelden afronden


# 11. Respons data bewerken -----------------------------------------------

# Aanmaken niveau variabelen
responsdata <-  responsbestand %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         GEMEENTECODE = ifelse(MIREB201 == regiocode, to_character(GEMEENTECODE), NA))

# Behouden van de benodigde variabelen
respons <- responsdata %>% 
  select(Respons_nettodich, Respons_netto, GEMEENTECODE, regio, totaal, nederland)

# Per gewenst niveau een respons tabel maken
responstabelGM <- respons %>%
  group_by(GEMEENTECODE)%>%
  summarise(Respons_perc = mean(Respons_nettodich, na.rm = T), respons_aantal = length(which(Respons_nettodich==1))) %>%
  rename(niveaurp=GEMEENTECODE)
responstabelRG <- respons %>%
  group_by(regio)%>%
  summarise(Respons_perc = mean(Respons_nettodich, na.rm = T), respons_aantal = length(which(Respons_nettodich==1))) %>%
  rename(niveaurp=regio) 
responstabelNL <- respons %>%
  group_by(nederland)%>%
  summarise(Respons_perc = mean(Respons_nettodich, na.rm = T), respons_aantal = length(which(Respons_nettodich==1))) %>%
  rename(niveaurp=nederland)

# Combineren van de responstabellen tot een
responstabeltotaal <- responstabelNL %>%
  rbind(responstabelRG)%>%
  rbind(responstabelGM) %>%
  drop_na()


# 12. Data in excelbestand zetten -----------------------------------------

# Data exporteren naar excelbestand met rapportage
# Als er geen working directory is gedefinieerd of bestanden bevinden zich buiten de working 
# directory dan moet de volledige padnaam worden opgegeven en niet alleen de bestandsnaam
wb <- loadWorkbook('Concept Rapportage VO.xlsx')
writeData(wb, sheet=1, output)
writeData(wb, sheet=2, trends_totaal)
writeData(wb, sheet=3, responstabeltotaal)
saveWorkbook(wb, 'Rapportage VO.xlsx', overwrite = TRUE)

