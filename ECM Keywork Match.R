## ----------------------------------------------------------------------------------------------
## Title: ECM Keyword Scraping.R
## Purpose: create a new file with only Keyword matches from ECM Master List Original,
##          and include rows that follow a matched row if they share the same Project #
## Author: Toby Bradshaw with use of Microsoft Copilot
## Date: 29 September 2025
## ----------------------------------------------------------------------------------------------

# Read in the original list -- replace file path with your file path
originalList <- read.csv("N:\\TBradshaw\\R\\ECM Naming\\ECM Master List Original.csv", stringsAsFactors = FALSE)

# Define the list of keywords directly in R -- change for what you want to sort for
target_words <- c("hospital", "clinic", "medical", "radiology", "pharmacy") ## c("lab","laboratory","labs","laboratories")## 

# Define output path as a variable -- replace file path with your file path
output_path <- "N:\\TBradshaw\\R\\ECM Naming\\ECM New Hospital Matches.csv"

# Always include the column descriptions
matching_rows <- originalList[1:2, ]

# Initialize ProjectNum
ProjectNum <- "AAAAA"

# Track matched indices to avoid duplicates
matched_indices <- c(1, 2)

# Loop through each row starting from the third
for (i in 3:nrow(originalList)) {
  row <- originalList[i, ]
  row_text <- tolower(as.character(row))
  current_project <- as.character(row[[3]])  # third column value
  
  # Check if any keyword matches
  if (any(sapply(target_words, function(word) any(grepl(word, row_text, fixed = FALSE))), na.rm = TRUE)) {
    matched_indices <- c(matched_indices, i)
    ProjectNum <- current_project
  } else if (current_project == ProjectNum) {
    matched_indices <- c(matched_indices, i)
  }
}

# Remove duplicates and sort
matched_indices <- sort(unique(matched_indices))

# Extract matched rows
result <- originalList[matched_indices, ]

# Write to CSV using the predefined output path
write.csv(result, output_path, row.names = FALSE)
cat("Matching rows saved to:", output_path, "\n")