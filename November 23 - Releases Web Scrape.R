### Packages -------------------------------------------------------------------

library(tidyverse)
library(rvest)
library(httr)
library(xml2)
library(stringi)
library(pander)
library(stringr)
library(rvest)
library(reshape)
library(openxlsx)


### Published ONS User Requests ------------------------------------------------

# provides the base URL for us to scrape
UR_base_url <- "https://www.ons.gov.uk/aboutus/whatwedo/statistics/requestingstatistics/alladhocs?page="

# create values for the pages we'll want to scrape
UR_pages <- 1:50

# put these together
UR_target_pages <- str_c(UR_base_url, UR_pages)

# read the html of the links provided above
UR_release_pages <- UR_target_pages %>% 
  map(read_html)

# Extract search result titles
UR_titles <- UR_release_pages %>% 
  map(html_elements, css = "h3.search-results__title") %>% 
  map(html_text) %>% 
  unlist() %>% 
  as.data.frame()

colnames(UR_titles)<- "Title"

UR_titles<-UR_titles %>% 
  mutate(Title = str_replace(Title, "\\s+", ""))

# Extract meta data
UR_meta_data <-  UR_release_pages %>% 
  map(html_elements, css = "p.search-results__meta") %>% 
  map(html_text) %>% 
  unlist() %>% 
  as.data.frame()

# Separate the column by '|'
UR_meta_data <- UR_meta_data %>%
  separate(col = 1, into = c("Title", "Date", "Reference"), sep = "\\|", extra = "drop", fill = "right") %>%
  mutate(across(everything(), str_trim))

# Drop the user requested column
UR_meta_data<-UR_meta_data %>% subset(select = c(2,3))

# Remove the 'Released on' portion of the string found in the second column
UR_meta_data <- UR_meta_data %>%
  mutate(Date = str_replace(Date, "Released on", "") %>% str_trim())

# Extract links from the pages
UR_release_links <- UR_release_pages %>% 
  map(html_nodes, css = "h3.search-results__title a") %>%
  map(html_attr, "href") %>%
  unlist() %>% 
  as.data.frame()

# add additional link information
UR_release_links<-paste("https://www.ons.gov.uk", UR_release_links$., sep = "")

# make into a df 
UR_release_links<-as.data.frame(UR_release_links)

# Function to extract information from a given URL with a delay
UR_extract_information <- function(url) {
  Sys.sleep(0.1)  # Introduce a delay of 0.1 seconds between requests to avoid rate limiting
  page <- read_html(url)
  
  # Extract 'Summary of request' section
  summary_request_heading <- page %>%
    html_nodes("h2:contains('Summary of request')") %>%
    html_text()
  
  # Extract text after 'Summary of request' until the next heading
  if (length(summary_request_heading) > 0) {
    summary_request <- page %>%
      html_nodes("h2:contains('Summary of request') + *") %>%
      html_text() %>%
      paste(collapse = "\n")
  } else {
    summary_request <- NA
  }
  
  return(data.frame(Summary_Request = summary_request, stringsAsFactors = FALSE))
}

# Loop through the target pages and extract information
UR_results_list <- lapply(UR_pages, function(page) {
  url <- paste0("https://www.ons.gov.uk/aboutus/whatwedo/statistics/requestingstatistics/alladhocs?page=", page)
  
  # Extract links from the search results pages
  search_page <- read_html(url)
  links <- search_page %>%
    html_nodes(".search-results__title a") %>%
    html_attr("href") %>%
    paste0("https://www.ons.gov.uk", .)
  
  # Loop through each link and extract information
  UR_link_results <- lapply(links, UR_extract_information)
  return(UR_link_results)
})

# Flatten the nested list
UR_descriptions <- bind_rows(lapply(unlist(UR_results_list, recursive = FALSE), as.data.frame))

colnames(UR_descriptions)<- "Summary of Request"

# binds together the release info and the link info
UR_release_final<-cbind(UR_titles, UR_meta_data, UR_release_links, UR_descriptions)

# For the confirmed releases, convert the Date column to a DateTime object
UR_release_final$Date <- dmy(UR_release_final$Date)


### Main Data Releases ---------------------------------------------------------

# provides the base URL for us to scrape
MR_base_url <- "https://www.gov.uk/search/research-and-statistics?content_store_document_type=upcoming_statistics&page="

# create values for the pages we'll want to scrape
MR_pages <- 1:50

# put these together
MR_target_pages <- str_c(MR_base_url, MR_pages)

# read the html of the links provided above
MR_release_pages <- MR_target_pages %>% 
  map(read_html)

# create a df using map for elements
MR_titles<-MR_release_pages %>% 
  map(html_elements, css = "li.gem-c-document-list__item") %>% 
  map(html_text) %>%
  unlist() %>% 
  as.data.frame()

# formatting our output
MR_release_final <- MR_titles %>%
  separate(col = 1,into = c("Blank1","Title","Description","Blank2","Doc Type","Blank3","Source","Blank4","Date","Blank5","Status","Blank6"), sep = "\n ") %>%
  select(-contains("Blank")) %>%
  mutate(across(.cols = everything(),.fn = trimws)) %>%
  mutate(`Doc Type` = str_replace(`Doc Type`,"Document type: ","")) %>%
  mutate(`Source` = str_replace(Source,"Organisation: ","")) %>%
  mutate(`Date` = str_replace(`Date`,"Release date: ","")) %>%
  mutate(State = str_replace(Status,"State: ","")) %>%
  mutate(State = str_to_title(State)) %>% 
  select(-`Status`)

# Extract links from the pages
MR_release_links <- MR_release_pages %>% 
  map(html_nodes, css = "li.gem-c-document-list__item a") %>%
  map(html_attr, "href") %>%
  unlist() %>% 
  as.data.frame()

# add additional link information
MR_release_links<-paste("https://www.gov.uk", MR_release_links$., sep = "")

# set this back to a data.frame
MR_release_links<-as.data.frame(MR_release_links)

# rename the column in the df so rest of code works
MR_release_links<-MR_release_links %>%
  `colnames<-`(c("Link"))

# binds together the release info and the link info
MR_release_final<-cbind(MR_release_final, MR_release_links)


# Identify rows with unconfirmed dates
MR_Unconfirmed_Releases <- MR_release_final %>%
  filter(!grepl("\\d{1,2} \\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\\b \\d{4} \\d{1,2}:\\d{2}(?:am|pm)\\b", Date, perl = TRUE))

# Create a new data frame with unconfirmed releases
MR_release_final <- MR_release_final %>%
  anti_join(MR_Unconfirmed_Releases)

# For the confirmed releases, convert the Date column to a DateTime object
MR_release_final$Date <- dmy_hm(MR_release_final$Date)

# Remove the time component
MR_release_final <- MR_release_final %>%
  mutate(Date = floor_date(Date, "day"))

### Outputs --------------------------------------------------------------------

# Write the output into one workbook for use in Tableau
setwd('M:\\Secure Folders\\Research\\Work\\Projects\\.Cross-team Workgroups\\WGRP_Data calendar\\Calender Explorer\\Testing')

# Create workbook with sheets
wb<-createWorkbook()
addWorksheet(wb, "User_Requested")
addWorksheet(wb, "Main_Data_Releases")
addWorksheet(wb, "Unconfrimed_Releases")

# Write data into sheets
writeData(wb, "User_Requested", UR_release_final)
writeData(wb, "Main_Data_Releases", MR_release_final)
writeData(wb, "Unconfrimed_Releases", MR_Unconfirmed_Releases)

# Create filename structure
currentdate <- Sys.Date()
filename <- paste(currentdate,"_Upcoming Releases and ONS User Requests - TEST",".xlsx", sep="")

# Write / Save workbook
saveWorkbook(wb, (file = filename), overwrite=TRUE)