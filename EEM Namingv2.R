# -----------------------------------------------------------------------------------------------
# Title: EEM Naming.R
# Purpose: Categorizes EEM lists from a small list of GBA Lab Projects
# Author: Toby Bradshaw with use of Microsoft Copilot
# Date: 20 August 2025
# -----------------------------------------------------------------------------------------------


library(readxl)
library(openxlsx)
library(NLP)
library(openNLP)
library(stringr)

# Load original spreadsheet -- replace file path with your file path
originalList <- read_excel("N:\\TBradshaw\\R\\ECM Naming\\ECM New Hospital Matches.xlsx")

# Define output path -- replace file path with your file path
outputPath1 <- "N:\\TBradshaw\\R\\ECM Naming\\Test2.xlsx"
outputPath2 <- "N:\\TBradshaw\\R\\ECM Naming\\Test2.csv"

tagPOS <- function(text) {
  s <- as.String(text)
  annotations <- NLP::annotate(s, list(sentence_annotator, word_token_annotator))
  annotations <- NLP::annotate(s, pos_tag_annotator, annotations)
  
  word_annots <- subset(annotations, type == "word")
  
  if (length(word_annots) == 0) {
    return(data.frame(token = character(0), tag = character(0), stringsAsFactors = FALSE))
  }
  
  tokens <- sapply(word_annots, function(w) substr(s, w$start, w$end))
  tags <- sapply(word_annots$features, `[[`, "POS")
  
  return(data.frame(token = tokens, tag = tags, stringsAsFactors = FALSE))
}

# Define capitalization and punctuation cleaning functions
capitalize_words <- function(text) {
  words <- unlist(strsplit(text, "\\s+"))
  words <- gsub("(^[[:alpha:]])", "\\U\\1", words, perl = TRUE)
  paste(words, collapse = " ")
}

remove_hanging_punctuation <- function(text) {
  open_count <- stringr::str_count(text, "\\(")
  close_count <- stringr::str_count(text, "\\)")
  if (open_count != close_count) {
    text <- gsub("[()]", "", text)
  }
  return(text)
}

# Define original word vectors for action categories
install <- c("install","use","add","insulate","implement","provide","seal","select","create","apply","make","add","added")
replace <- c("replace","replaced","replacing","convert","replacement")
retrofit <- c("retrofit","retrofits","upgrade","improve","change","modify","convert","conversion")
adjust <- c("adjust","adjustments","reduce","reduction","set","turn off","minimize","increase","lower","optimize",
            "supply","avoid","correct","vary","refine","modification","modifications","modify","revise")
remove <- c("remove","removes","eliminate","separate","removal")
repair <- c("repair","repairs","clean","check","maintain","refurbish","fix")
weatherize <- c("weatherize","caulk","seal")
perform <- c("perform","commissioning","study","retrocommissioning","rcx","coordinate","optimization","reset")

sentence_annotator <- Maxent_Sent_Token_Annotator()
word_token_annotator <- Maxent_Word_Token_Annotator()
pos_tag_annotator <- Maxent_POS_Tag_Annotator()

# Store original first keyword for output
actionLabels <- list(
  Install = install[1], Replace = replace[1], Retrofit = retrofit[1], Adjust = adjust[1],
  Remove = remove[1], Repair = repair[1], Weatherize = weatherize[1], Perform = perform[1]
)

# Store lowercase versions for matching
actionCategories <- list(
  Install = tolower(install), Replace = tolower(replace), Retrofit = tolower(retrofit), Adjust = tolower(adjust),
  Remove = tolower(remove), Repair = tolower(repair), Weatherize =tolower(weatherize), Perform = tolower(perform)
)

# Define original word vectors for system categories
lighting <- c("Lighting","light","lumen","occupancy","fluorescent","lamp","lamps","LED","delamp","wall pack","wall packs")
airDistribution <- c("Air Distribution","ventilation","airflow","airflows","ahu","duct","hvac","exhaust","vfd","vfds",
                     "ahu's","ahus","VAV","variable air volume","VS","fan","fans","supply","air handling unit","air handling units",
                     "economizer","variable air volume","air","RTU")
coolingSystem <- c("Cooling System","chiller","chillers","chw","cw","chilled","cooling","cool","thermostat","thermostats","CHWP")
heatingSystem <- c("Heating System","heat","heater","heaters","heating","boiler","hw","thermostat","thermostats","reheat","hot water",
                   "hot water system","steam")
controlSystem <- c("Controls", "BAS","automation","user based","user based controls","setpoint","control",
                   "controls","sensor","sensors","schedule","scheduling","setback")
waterSystem <- c("Water System","sewer","condenser","pump","water","valve","Low-Flow","pumping","variable volume","filter")
energySupply <- c("Energy Supply","battery","generate","generation","generator","solar","wind","nuclear")
Envelope <- c("Envelope","window","windows","door","doors","caulk","building")
Services <- c("Services", "study","retrocommissioning","commissioning","rcx","cx")

# Store original first keyword for output
systemLabels <- list(
  Lighting = lighting[1], AirDistribution = airDistribution[1], CoolingSystem = coolingSystem[1], 
  HeatingSystem = heatingSystem[1], controlSystem = controlSystem[1], 
  WaterSystem = waterSystem[1], EnergySupply = energySupply[1], Envelope = Envelope[1], Services = Services[1]
)

# Store lowercase versions for matching
systemCategories <- list(
  Lighting = tolower(lighting), AirDistribution = tolower(airDistribution), CoolingSystem = tolower(coolingSystem),
  HeatingSystem = tolower(heatingSystem), controlSystem = tolower(controlSystem), WaterSystem = tolower(waterSystem), EnergySupply = tolower(energySupply),
  Envelope = tolower(Envelope), Services = tolower(Services)
)

# Identify the column index for "description" -- change for your description column
desc_index <- 10

# Initialize list to store processed rows
tagged_rows <- list()

# Loop through each row
for (i in 1:nrow(originalList)) {
  row <- originalList[i, ]
  desc_raw <- as.character(row[[desc_index]])
  
  # Remove all dash-followed segments (from each dash to the next space or end of string)
  desc_clean <- gsub("\\([^)]*\\)", "", desc_raw)
  desc_clean <- gsub("-\\d[^ ]*", "", desc_clean)  
  desc_clean <- gsub("'[^ ]*", "", desc_clean)
  desc_clean <- gsub("â€“[^ ]*", "", desc_clean) # removes weird thing â€“
  desc_clean <- gsub("â[^ ]*", "2", desc_clean) # removes weird thing â
  desc_clean <- gsub("&[^ ]*", "", desc_clean) # removes &
  desc_clean <- gsub("%[^ ]*", "", desc_clean) # removes %
  desc_clean <- gsub("/", " ", desc_clean)
  desc_clean <- gsub("\\b\\d+\\s+(through|and)\\s+\\d+\\b", "", desc_clean, ignore.case = TRUE) # Remove patterns like "number through number" or "number and number"
  desc_clean <- gsub("(?<=^|\\s)RCx(?=\\s|$)", "Retrocommissioning", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Retro-commissioning(?=\\s|$)", "Retrocommissioning", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Cx(?=\\s|$)", "Commissioning", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Air Handling Unit(?=\\s|$)", "AHU", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Air Handling Units(?=\\s|$)", "AHU", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Air-Handling Unit(?=\\s|$)", "AHU", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Air-Handling Units(?=\\s|$)", "AHU", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)Variable Air Volume(?=\\s|$)", "VAV", desc_clean, perl = TRUE)
  
  # cleans abbreviations
  desc_clean <- gsub("(?<=^|\\s)S(?=\\s|$)", "AHU", desc_clean, perl = TRUE)      # Replace "S" with "AHU"
  desc_clean <- gsub("(?<=^|\\s)SF(?=\\s|$)", "AHU", desc_clean, perl = TRUE)     # Replace "SF" with "AHU"
  desc_clean <- gsub("(?<=^|\\s)BS(?=\\s|$)", "AHU", desc_clean, perl = TRUE)     # Replace "BS" with "AHU"
  desc_clean <- gsub("(?<=^|\\s)VS(?=\\s|$)", "AHU", desc_clean, perl = TRUE)     # Replace "VS" with "AHU"
  desc_clean <- gsub("(?<=^|\\s)CH(?=\\s|$)", "Chiller", desc_clean, perl = TRUE) # Replace "CH" with "Chiller"
  desc_clean <- gsub("(?<=^|\\s)2-Speed(?=\\s|$)", "Two Speed", desc_clean, perl = TRUE)
  desc_clean <- gsub("(?<=^|\\s)\\w(?=\\s|$)", "", desc_clean, perl = TRUE)   # Remove single-character tokens
  desc_clean <- gsub("\\b\\d+\\b", "", desc_clean, perl = TRUE)  # remove single number tokens
  
  # removes location specifics
  locationList <- c("area", "lab", "laboratory", "office", "elevator", "lobby", "vending machine", "ER", "cath", "stairwell",
                    "stair","halogen","parking garage","kitchen","cafeteria","surgery","hall","ballroom","garage","laundry")
  pattern <- paste0("\\b(", paste(locationList, collapse = "|"), ")\\b")
  desc_clean <- gsub(pattern, "", desc_clean, ignore.case = TRUE)
  
  # Tokenize cleaned description
  desc_text <- if (!is.na(desc_clean)) tolower(unlist(strsplit(desc_clean, "\\s+"))) else character(0)
  
  # Match action category
  matched_action <- ""
  for (category in names(actionCategories)) {
    matched_tokens <- intersect(desc_text, actionCategories[[category]])
    if (length(matched_tokens) > 0) {
      matched_action <- actionLabels[[category]]
      if (matched_action != "perform"){
        desc_text <- desc_text[!desc_text %in% matched_tokens]
      }
      # Remove all matched action tokens from desc_text
      break
    }
  }
  
  if (matched_action == "") matched_action <- "Install!"
  
  # Match system categories (allow multiple systems)
  matched_systems <- c()
  
  for (category in names(systemCategories)) {
    matched_tokens <- intersect(desc_text, systemCategories[[category]])
    if (length(matched_tokens) > 0) {
      matched_systems <- unique(c(matched_systems, systemLabels[[category]]))
    }
  }
  
  # If no matches found, set to "/"
  if (length(matched_systems) == 0) {
    matched_systems <- "/"
  }
  
  # Combine system tags
  combined_systems <- if (length(matched_systems) > 1) {
    paste(matched_systems, collapse = "/")
  } else if (length(matched_systems) == 1) {
    matched_systems[1]
  } else {
    ""
  }
  
  # Initialize POS data early
  pos_data <- data.frame(token = character(0), tag = character(0), stringsAsFactors = FALSE)
  
  # POS tagging
  cleaned_desc <- paste(desc_text, collapse = " ")
  if (nchar(trimws(cleaned_desc)) > 0) {
    pos_data <- tagPOS(cleaned_desc)
  }
  
  # Repeated trimming based on prepositions, colons, and dashes
  repeat {
    if (nrow(pos_data) == 0) break
    
    # Define all marker tokens
    marker_tokens <- c(",",";", ":", "-", "–", "to", "in", "on", "at", "by", "with", "from", "of", "for")
    
    # Identify positions of any marker tokens
    marker_indices <- which(tolower(pos_data$token) %in% marker_tokens)
    
    if (length(marker_indices) == 0) break  # No more markers to trim
    
    trimmed <- FALSE
    
    for (cutoff in marker_indices) {
      left_tokens <- if (cutoff > 1) pos_data[1:(cutoff - 1), ] else pos_data[0, ]
      right_tokens <- if (cutoff < nrow(pos_data)) pos_data[(cutoff + 1):nrow(pos_data), ] else pos_data[0, ]
      
      contains_system <- function(tokens) {
        tokens_lower <- tolower(tokens$token)
        any(unlist(lapply(systemCategories, function(words) any(tokens_lower %in% words))))
      }
      
      left_has_system <- contains_system(left_tokens)
      right_has_system <- contains_system(right_tokens)
      
      # Favor left if it contains system keywords
      if (left_has_system) {
        pos_data <- left_tokens
      } else if (right_has_system) {
        pos_data <- right_tokens
      } else {
        pos_data <- left_tokens  # Default to left if neither side has system keywords
      }
      
      trimmed <- TRUE
      break  # Trim only once per repeat cycle
    }
    
    if (!trimmed) break  # No valid trim occurred
  }
  
  # Remove adjectives
  if ("tag" %in% names(pos_data) && nrow(pos_data) > 0) {
    pos_data <- pos_data[pos_data$tag != "JJ", ]
  }
  
  # Final description
  truncated_text <- paste(pos_data$token, collapse = " ")
  matched_action2 <- gsub("!", "", matched_action)
  raw_description <- paste(matched_action2, truncated_text)
  cleaned_description <- remove_hanging_punctuation(raw_description)
  modified_description <- capitalize_words(cleaned_description)
  modified_description <- gsub("(?<=^|\\s)And(?=$)", "", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)And (?=$)", "", modified_description, perl = TRUE)
  
  # fixes acronym issues
  modified_description <- gsub("(?<=^|\\s)Led(?=\\s|$)", "LED", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Ahu(?=\\s|$)", "AHU", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Ahu-o(?=\\s|$)", "AHU", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Ahus(?=\\s|$)", "AHU", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Vfd(?=\\s|$)", "VFD", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Vfds(?=\\s|$)", "VFD", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Vav(?=\\s|$)", "VAV", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Hx(?=\\s|$)", "HX", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Chw(?=\\s|$)", "CHW", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Chwp(?=\\s|$)", "CHWP", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Hw(?=\\s|$)", "HW", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Fcu(?=\\s|$)", "FCU", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Prv(?=\\s|$)", "PRV", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Rtu(?=\\s|$)", "RTU", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Ac(?=\\s|$)", "AC", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Acs(?=\\s|$)", "ACS", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Oa(?=\\s|$)", "OA", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Cav(?=\\s|$)", "CAV", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Hvac(?=\\s|$)", "HVAC", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Co2(?=\\s|$)", "CO2", modified_description, perl = TRUE)
  modified_description <- gsub("(?<=^|\\s)Or(?=\\s|$)", "OR", modified_description, perl = TRUE)
  
  # Normalize spacing again after removal
  modified_description <- gsub("\\s+", " ", modified_description)
  modified_description <- trimws(modified_description)
  
  # fix AHU repeats
  modified_description <- sub("\\bahu\\b", "___KEEP_AHU___", modified_description, ignore.case = TRUE)
  modified_description <- gsub("\\bahu\\b", "", modified_description, ignore.case = TRUE)
  modified_description <- gsub("___KEEP_AHU___", "AHU", modified_description)
  
  # Reconstruct row
  row_before <- row[1:desc_index]
  row_after <- if (desc_index < length(row)) row[(desc_index + 1):length(row)] else NULL
  
  new_row <- cbind(row_before,
                   ActionTag = matched_action,
                   SystemTag = combined_systems,
                   ModifiedDescription = modified_description,
                   row_after)
  
  tagged_rows[[length(tagged_rows) + 1]] <- new_row
}

# Combine all rows and write to CSV
result <- do.call(rbind, tagged_rows)
write.xlsx(result, outputPath1, rowNames = FALSE)
write.csv(result, outputPath2, row.names = FALSE)

data <- read.csv(outputPath2, stringsAsFactors = FALSE)

# Loop through each column and replace "NA" strings with empty string
for (col in names(data)) {
  data[[col]][data[[col]] == "NA"] <- ""
}

# Write the cleaned data back to CSV
write.csv(data, outputPath2, row.names = FALSE)

cat("Excel file saved to:", outputPath1, "\n")
cat("CSV file saved to:", outputPath2, "\n")
