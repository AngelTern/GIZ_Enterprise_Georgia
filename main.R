library(readxl)
library(dplyr)
library(rstudioapi)

# Ensure the rstudioapi package is available
if (requireNamespace("rstudioapi", quietly = TRUE)) {
  # Print a message indicating the script is running in RStudio
  print("Running in RStudio")
  
  # Get the path of the active script
  script_path <- rstudioapi::getActiveDocumentContext()$path
  print(paste("Script path:", script_path))
  
  if (!is.null(script_path) && script_path != "") {
    # Extract the directory from the script path
    script_dir <- dirname(script_path)
    print(paste("Script directory:", script_dir))
    
    # Set the working directory to the script's directory
    setwd(script_dir)
    
    # Print the new working directory
    print(paste("New working directory:", getwd()))
  } else {
    print("Script path is null or empty. Ensure the script is sourced from an active RStudio document.")
  }
} else {
  stop("This script requires RStudio to run.")
}

# Define the directory containing the Excel files
#directory <- file.path(script_dir, "data_main")

# List all Excel files in the directory ----------------------------- for testing
excel_files <- list.files(file.path(script_dir, "data_test"), pattern = "\\.xlsx$", full.names = TRUE)

# List all Excel files in the directory ----------------------------- for main
#excel_files <- list.files(file.path(script_dir, "data_main"), pattern = "\\.xlsx$", full.names = TRUE)

# Function to create valid R list element names from file names
make_valid_name <- function(file_path) {
  file_name <- basename(file_path) # Get the base name of the file
  file_name <- tools::file_path_sans_ext(file_name) # Remove the file extension
  file_name <- make.names(file_name) # Make it a valid R list element name
  return(file_name)
}

# Initialize an empty list to store the data frames
data_list <- list()

# Iterate over the list of Excel files and read each one
for (file in excel_files) {
  # Create a valid list element name from the file name
  var_name <- make_valid_name(file)
  
  # Read the Excel file and store it in the list with the name
  data_list[[var_name]] <- read_excel(file)
}

# Print the list names to verify (optional)
print(names(data_list))

# Function to get sorted unique values of each column, excluding specified columns
get_sorted_unique_values <- function(df, exclude_cols = c("ReportCode", "ReportId", "IdCode", "OrgNameInReport")) {
  unique_values_list <- list()
  for (col in colnames(df)) {
    # Skip the columns to be excluded
    if (col %in% exclude_cols) {
      next
    }
    # Get sorted unique values
    unique_values <- sort(unique(df[[col]]))
    unique_values_list[[col]] <- unique_values
  }
  return(unique_values_list)
}

# Initialize an empty list to store the groups of lists
grouped_unique_values <- list()

# Iterate over each data frame in the list and get sorted unique values for each column
for (name in names(data_list)) {
  grouped_unique_values[[name]] <- get_sorted_unique_values(data_list[[name]])
}

# Function to merge unique values across all data frames for each column name
merge_unique_values <- function(grouped_list) {
  merged_unique_values <- list()
  for (df_name in names(grouped_list)) {
    for (col_name in names(grouped_list[[df_name]])) {
      if (!is.null(merged_unique_values[[col_name]])) {
        merged_unique_values[[col_name]] <- sort(unique(c(merged_unique_values[[col_name]], grouped_list[[df_name]][[col_name]])))
      } else {
        merged_unique_values[[col_name]] <- grouped_list[[df_name]][[col_name]]
      }
    }
  }
  return(merged_unique_values)
}

consolidated_unique_values <- merge_unique_values(grouped_unique_values)

#Define formula to flatten and strip lineitems

flatten_lineitem <- function(input, column_name = NULL){
  if (is.data.frame(input)){
    
    if(is.null(column_name) || !(column_name %in% colnames(input))){
      stop("Provide a valid column name")
    }
    
    input[[column_name]] <- gsub("[^a-zA-Z]", "", input[[column_name]])
    input[[column_name]] <- tolower(input[[column_name]])
    
    return(input)
  } else if (is.list(input) || is.vector(input)){
    
    input <- gsub("[^a-zA-Z]", "", unlist(input))
    input <- tolower(input)
    
    return(input)
  } else {
    stop("Input must either be a dataframe, vector, or a list")
  }
}

# Filtering criteria

## Columns to keep

columns_to_keep_1st_revision <- c("ReportCode", "IdCode", "ReportYear", "FVYear", "CategoryMain", "FormName",
                                  "SheetName", "LineItemGEO", "LineItemENG", "Value", "GEL")

###TBD
###columns_to_keep_2nd_revision <-list()

# Needed variables lists

variables_financial_non_financial <- list('Cash and cash equivalents', 'Current Inventory', 'Non current inventory', 'Trade receivables',
                                          'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
                                          'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
                                          'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
                                          'Total equity', 'Total liabilities and equity', 'Cash advances made to other parties', 'Investment property',
                                          'Investments in subsidiaries', 'Goodwill', 'Other intangible assets', 'Finance lease payable', 'Unearned income',
                                          'Current borrowings', 'Non current borrowings', 'Received grants', 'Total current assets', 'Total current liabilities',
                                          'Share capital (in case of Limited Liability Company - "capital", in case of cooperative entity - "unit capital"'
                                          )



variables_financial_other <- list('Cash and cash equivalents', 'Inventories', 'Trade receivables',
                                  'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
                                  'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
                                  'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
                                  'Total equity', 'Total liabilities and equity')

variables_profit_loss <- list('Net Revenue', 'Cost of goods sold', 'Gross profit', 'Other operating income',
                              'Personnel expense', 'Rental expenses', 'Depreciation and amortisation',
                              'Other administrative and operating expenses', 'Operating income', 
                              'Impairment (loss)/reversal of financial assets', 'Inventories',
                              'Net gain (loss) from foreign exchange operations', 'Dividends received',
                              'Other net operating income/(expense)', 'Profit/(loss) before tax from continuing operations',
                              'Income tax', 'Profit/(loss)', 'Revaluation reserve of property, plant and equipment',
                              'Other (include Share of associates and joint ventures in revaluation reserve of property, plant and equipment and defined benefit obligation)',
                              'Total other comprehensive (loss) income', 'Total comprehensive income / (loss)')

variables_cash_flow <- list('Net cash from operating activities', 'Net cash used in investing activities',
                            'Net cash raised in financing activities', 'Net cash inflow for the year',
                            'Effect of exchange rate changes on cash and cash equivalents',
                            'Cash at the beginning of the year', 'Cash at the end of the year')


# Function to filter out rows based on a condition

# Define the primary processing function, for initial optimization and filtering 
process_df_primary <- function(df, columns_to_keep) {
  df %>%
    select(all_of(columns_to_keep)) %>%
    filter(CategoryMain != "III ჯგუფი") %>%
    mutate(Value = if ("Value" %in% colnames(df)) as.numeric(Value) else Value) %>%
    # Changes in FormName
    mutate(
      FormName = case_when(
        FormName == "არაფინანსური ინსტიტუტებისთვის" ~ "non-financial institutions",
        FormName == "გამარტივებული ფორმები მესამე კატეგორიის საწარმოებისთვის" ~ "Cat III forms",
        FormName == "მეოთხე კატეგორიის საწარმოთა ანგარიშგების ფორმები" ~ "Cat IV forms",
        TRUE ~ FormName
      ),
      SheetName = case_when(
        SheetName == "საქმიანობის შედეგები" ~ "profit-loss",
        SheetName == "ფინანსური მდგომარეობა" ~ "financial position",
        SheetName == "ფულადი სახსრების მოძრაობა" ~ "cash flow",
        TRUE ~ SheetName
      ),
      Value = case_when(
        GEL == ".000 ლარი" & !is.na(Value) ~ Value * 1000,
        TRUE ~ Value
      )
    ) %>%
    filter(FormName != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") %>%
    filter(!(FormName == "non-financial institutions" & !LineItemENG %in% variables_financial_non_financial)) %>%
    filter(!(FormName == "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_non_financial)) %>%
    filter(!(FormName != "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_other)) %>%
    filter(!(SheetName == "profit-loss" & !LineItemENG %in% variables_profit_loss)) %>%
    filter(!(SheetName == "cash flow" & !LineItemENG %in% variables_cash_flow)) %>%
    arrange(ReportCode)
}



primary_processed_list <- lapply(data_list, function(df) process_df_primary(df, columns_to_keep_1st_revision))
# Final dataframe preparation

find_unique_values <- function(df, cols) {
  unique_values <- lapply(df[cols], unique)
  names(unique_values) <- cols
  return(unique_values)
}

columns_to_filter_for_unique <- c('ReportYear', 'FVYear' ,'LineItemENG')

unique_values_for_columns <- lapply(primary_processed_list, function(df) find_unique_values(df, columns_to_filter_for_unique))

combined_df <- bind_rows(primary_processed_list)




list_of_dfs_group_split_by_report_code <- combined_df %>% group_split(ReportCode)


nested_list_of_dfs_group_split_by_lineitemeng <- lapply(list_of_dfs_group_split_by_report_code, function(df) {
  df %>% group_split(LineItemENG)
})

process_df_secondary <- function(df) {
  processed_year <- list()
  column_to_check <- 'ReportYear'
  instances_over_two <- 0
  
  for (i in 1:nrow(df)) {
    current_value <- df$FVYear[i]
    
    if (!(current_value %in% processed_year)) {
      processed_year <- append(processed_year, current_value)
      found_matches <- which(df$FVYear == current_value)
      print(found_matches)  # found_matches is a vector of indices
      if (length(found_matches) > 1) {
        if (length(found_matches) > 2) {
          instances_over_two <- instances_over_two + 1
        }
        
        # Ensure column_to_check is numeric before which.max
        column_data <- df[found_matches, column_to_check]
        value_data <- df[found_matches, "Value"]
        
        # If column_data is not numeric, convert it
        if (!is.numeric(column_data)) {
          column_data <- as.numeric(as.character(column_data))
        }
        
        # Sort column_data and value_data in decreasing order
        sorted_indices <- order(column_data, decreasing = TRUE)
        column_data <- column_data[sorted_indices]
        value_data <- value_data[sorted_indices]
        
        '# Print debug information
        print(typeof(column_data))
        print(length(column_data))
        for (k in seq_along(column_data)) {
          print(column_data[k])
          print(typeof(column_data[k]))
        }'
        
        # Find the first non-zero value in the sorted order
        non_zero_index <- which(value_data != 0)
        if (length(non_zero_index) > 0) {
          # Use the first non-zero value
          max_value_index <- found_matches[sorted_indices[non_zero_index[1]]]
        } else {
          # All values are zero, use the latest year value
          max_value_index <- found_matches[sorted_indices[1]]
        }
        
        found_matches <- found_matches[found_matches != max_value_index]
        
        # Debugging: Print the indices to be removed
        cat("Dropping rows with indices:", found_matches, "\n")
        
        df <- df[-found_matches, ]
      }
    }
  }
  
  # Debugging: Print the number of rows after processing
  cat("Number of rows after processing:", nrow(df), "\n")
  cat("Number of instances with more than two matches:", instances_over_two, "\n")
  
  return(df)
}


final_processed_list <- lapply(nested_list_of_dfs_group_split_by_lineitemeng, function(inner_list){
  lapply(inner_list, function(df) process_df_secondary(df))
})

# Flatten the nested list correctly
flattened_final_list <- do.call(c, final_processed_list)

final_df <- bind_rows(flattened_final_list)



#################

all_picked_variables = unique(c(variables_financial_non_financial, variables_financial_other, variables_profit_loss, variables_cash_flow))


all_variables_given <- read_excel("all_lineitemeng_variables.xlsx")
all_variables_given <- as.list(all_variables_given$Variables)

given_track <- c()
found_track <- c()

variables_to_track <- list(given_track = given_track, found_track = found_track)

for (i in all_picked_variables){
  if (!(i %in% all_variables_given)) {
    print(i)
    variables_to_track$given_track <- append (i, variables_to_track$given_track)
  }
}

all_found_lineitemeng <- consolidated_unique_values[["LineItemENG"]]

for (i in all_found_lineitemeng){
  if(!(i %in% all_found_lineitemeng)){
    print(i)
    variables_to_track$found_track <- append(i, variables_to_track$found_track)
  }
}


all_foundlineitemgeo <- consolidated_unique_values[["LineItemGEO"]]

'#To find what LineItemGEO correspond to all_picked_variables

filtered_df_1 <- combined_df %>% filter(LineItemENG %in% all_picked_variables)

filtered_grouped_df_1 <- filtered_df_1 %>% group_by(LineItemENG) %>% summarise(LineItemGEO = list(unique(LineItemGEO)))

filtered_result <- setNames(filtered_grouped_df_1$LineItemGEO, filtered_grouped_df_1$LineItemENG)

missing_variables <- setdiff(all_picked_variables, filtered_grouped_df_1$LineItemENG)'




##################

# Statistics

## Preparation


third_list <- list()
fourth_list <- list()

split_dfs <- split(final_df, final_df$ReportYear)

category_report_process <- function(df) {
  
  year <- df$ReportYear[0]
  
  name_for_third <- paste0("third_", year)
  third_list[[name_for_third]] <- list()
  
  name_for_fourth <- paste0("fourth_", year)
  fourth_list[[name_for_fourth]] <- list()
  
  for (i in 1:nrow(df)){
    report_code <- df$ReportCode[i]
    id_code <- df$IdCode[i]
    category <- df$CategoryMain[i]
    
    
    
    if (!(is.na(id_code)) && category == 'III' && !(id_code %in% third_list[[name_for_third]])){
      third_list[[name_for_third]] <- append(id_code, third_list[[name_for_third]])
    }
    ifelse (is.na(id_code) && category == 'IV' && !(report_code %in% fourth_list[[name_for_fourth]])){
      fourth_list[[name_for_fourth]] <- append(report_code, fourth_list[[name_for_fourth]])
    }
  }
}

category_report_over_time <- function(list_a) {
  
}









# Check if final_df has the same number of rows as combined_df
cat("Number of rows in combined_df:", nrow(combined_df), "\n")
cat("Number of rows in final_df:", nrow(final_df), "\n")


write.csv(combined_df, file = "final_data.csv", row.names = FALSE, fileEncoding = "UTF-8")


'# Benchmark split function
split_time <- system.time({
  list_of_dfs_split <- split(combined_df, combined_df$ReportCode)
})

# Benchmark dplyr::group_split function
group_split_time <- system.time({
  list_of_dfs_group_split <- combined_df %>% group_split(ReportCode)
})

# Benchmark custom loop
custom_loop_time <- system.time({
  unique_codes <- unique(combined_df$ReportCode)
  list_of_dfs_custom <- vector("list", length(unique_codes))
  names(list_of_dfs_custom) <- unique_codes
  for (code in unique_codes) {
    list_of_dfs_custom[[code]] <- combined_df[combined_df$ReportCode == code, ]
  }
})

# Print benchmark results
print(split_time)
print(group_split_time)
print(custom_loop_time)'




'process_df_secondary <- function(df) {
  df %>%
    select(-FormName, -SheetName, -GEL) %>%
    group_by(ReportCode, FVYear) 
} 



secondary_processed_list <- process_df_secondary(combined_df)


process_and_split <- function(df) {
  # Group by ReportCode and FVYear, and filter out groups with only one row
  grouped <- df %>%
    group_by(ReportCode, FVYear) %>%
    filter(n() > 1) %>%
    ungroup()
  
  # Initialize an empty list to store the columns
  result_list <- list()
  
  # Iterate over unique combinations of ReportCode and FVYear
  for (code in unique(grouped$ReportCode)) {
    for (year in unique(grouped$FVYear)) {
      subset_df <- grouped %>%
        filter(ReportCode == code, FVYear == year)
      
      if (nrow(subset_df) > 0) {
        instance_name <- paste(code, year, sep = "_")
        result_list[[paste(instance_name, "ReportYear", sep = "_")]] <- subset_df$ReportYear
        result_list[[paste(instance_name, "Value", sep = "_")]] <- subset_df$Value
      }
    }
  }
  
  # Convert the list to a dataframe
  result_df <- as.data.frame(result_list)
  
  return(result_df)
}

# Apply the function to the sample dataframe
result_df <- process_and_split(combined_df)'



'# Get the current working directory

working_dir <- getwd()

# Output the combined data frame to a CSV file in the working directory

output_file_path <- file.path(working_dir, "combined_data.csv")
write.csv(combined_df, output_file_path, row.names = FALSE)

# Print message to confirm completion and file location

cat("Combined data frame has been written to:", output_file_path, "\n")
'

'combined_variable_list <- c(variables_financial_non_financial, variables_financial_other, variables_profit_loss, variables_cash_flow)

for (i in list(unique(processed_list$X2022.Lineitems.Cat.III$LineItemENG))) {
  if (!(i %in% combined_list)) {
    cat(i, "is not in the combined list\n")
    print(processed_list$X2022.Lineitems.Cat.III[processed_list$X2022.Lineitems.Cat.III$LineItemENG == item, ])
  }
    
}'


# Count values of variables

#count_values <- function(df, values_to_check){
#  df %>%
#    filter(FormName )
#}

#value_count_non_financial <- 1


