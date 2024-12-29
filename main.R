# Load necessary libraries for data manipulation and visualization
library(readxl)      # For reading Excel files
library(dplyr)       # For data manipulation
library(rstudioapi)  # For interacting with RStudio
library(tidyr)       # For data tidying
library(ggplot2)     # For data visualization
library(writexl)     # For writing Excel files

# Ensure the script is running within RStudio and set the working directory to the script's directory
if (requireNamespace("rstudioapi", quietly = TRUE)) {
  print("Running in RStudio")
  
  # Obtain the path of the currently active script
  script_path <- rstudioapi::getActiveDocumentContext()$path
  print(paste("Script path:", script_path))
  
  # If the script path is valid, set the working directory accordingly
  if (!is.null(script_path) && script_path != "") {
    script_dir <- dirname(script_path)
    print(paste("Script directory:", script_dir))
    
    # Set the working directory to the script's directory
    setwd(script_dir)
    print(paste("New working directory:", getwd()))
  } else {
    print("Script path is null or empty. Ensure the script is sourced from an active RStudio document.")
  }
} else {
  stop("This script requires RStudio to run.")
}

# List all Excel files in the 'data_cat_4' directory within the script's directory
excel_files <- list.files(file.path(script_dir, "data_main"), pattern = "\\.xlsx$", full.names = TRUE)

# Function to create valid R list element names from file names
make_valid_name <- function(file_path) {
  # Extract the base name of the file (without directory path)
  file_name <- basename(file_path)
  # Remove the file extension from the file name
  file_name <- tools::file_path_sans_ext(file_name)
  # Convert the file name into a valid R object name
  file_name <- make.names(file_name)
  return(file_name)
}

# Initialize an empty list to store the data frames from Excel files
data_list <- list()

# Iterate over each Excel file and read its contents into the list
for (file in excel_files) {
  # Create a valid list element name from the file name
  var_name <- make_valid_name(file)
  # Read the Excel file and store the data frame in the list with the name
  data_list[[var_name]] <- read_excel(file)
}

#Verify column types

check_column_type_consistency <- function(data_list) {
  # Helper function to get column types for a single data frame
  get_column_types <- function(df) {
    sapply(df, class)
  }
  
  # Step 1: Get column types for each data frame
  column_types_list <- lapply(data_list, get_column_types)
  
  # Step 2: Determine the standard column type for each column
  all_columns <- unique(unlist(lapply(column_types_list, names)))
  standard_types <- list()
  
  for (col in all_columns) {
    # Collect types for the column across all data frames where it exists
    column_types <- unlist(lapply(column_types_list, function(types) {
      if (col %in% names(types)) types[col] else NA
    }))
    
    # Remove NA values and determine the most common type
    column_types <- column_types[!is.na(column_types)]
    if (length(column_types) > 0) {
      standard_types[[col]] <- names(sort(table(column_types), decreasing = TRUE))[1]
    }
  }
  
  # Step 3: Check for deviations from the standard type
  deviations <- list()
  
  for (df_name in names(data_list)) {
    df <- data_list[[df_name]]
    df_types <- get_column_types(df)
    
    for (col in names(df_types)) {
      if (col %in% names(standard_types)) {
        if (df_types[col] != standard_types[[col]]) {
          # Record the deviation
          deviations[[df_name]] <- c(
            deviations[[df_name]],
            paste0("Column '", col, "' is of type '", df_types[col],
                   "' but should be '", standard_types[[col]], "'.")
          )
        }
      }
    }
  }
  
  # Step 4: Print results
  if (length(deviations) > 0) {
    cat("Deviations in column types:\n")
    for (df_name in names(deviations)) {
      cat("\nDataFrame:", df_name, "\n")
      cat(paste(deviations[[df_name]], collapse = "\n"), "\n")
    }
  } else {
    cat("All columns have consistent types across data frames.\n")
  }
  
  # Return the standard types for reference
  return(standard_types)
}

# Example usage
standard_types <- check_column_type_consistency(data_list)

#Adjust deviating columns
adjust_column_type_deviations <- function(data_list, standard_types) {
  # Function to coerce a column to a specific type
  coerce_column_type <- function(column, target_type) {
    if (target_type == "character") {
      return(as.character(column))
    } else if (target_type == "numeric") {
      return(as.numeric(column))
    } else if (target_type == "integer") {
      return(as.integer(column))
    } else if (target_type == "logical") {
      return(as.logical(column))
    } else if (target_type == "factor") {
      return(as.factor(column))
    } else {
      stop(paste("Unsupported target type:", target_type))
    }
  }
  
  # Adjust each data frame in the list
  adjusted_data_list <- lapply(data_list, function(df) {
    for (col in names(standard_types)) {
      if (col %in% colnames(df)) {
        current_type <- class(df[[col]])
        target_type <- standard_types[[col]]
        if (current_type != target_type) {
          cat(paste("Adjusting column", col, "from", current_type, "to", target_type, "\n"))
          df[[col]] <- coerce_column_type(df[[col]], target_type)
        }
      }
    }
    return(df)
  })
  
  return(adjusted_data_list)
}

adjusted_data_list <- adjust_column_type_deviations(data_list, standard_types)


# Define the columns to keep during the initial filtering of data
columns_to_keep_1st_revision <- c(
  "ReportCode", "IdCode", "ReportYear", "FVYear", "CategoryMain", "FormName",
  "SheetName", "LineItemGEO", "LineItemENG", "Value", "GEL", "LineItem"
)

# Define lists of necessary variables for different financial statements for CAT III
####
#variables_financial_non_financial <- list(
#  'Cash and cash equivalents', 'Current Inventory', 'Non current inventory', 'Trade receivables',
#  'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
#  'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
#  'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
#  'Total equity', 'Total liabilities and equity', 'Cash advances made to other parties', 'Investment property',
#  'Investments in subsidiaries', 'Goodwill', 'Other intangible assets', 'Finance lease payable', 'Unearned income',
#  'Current borrowings', 'Non current borrowings', 'Received grants', 'Total current assets', 'Total current liabilities',
#  'Share capital'
#)

variables_financial_non_financial <- list(
  "Cash and cash equivalents", "Current Inventory", "Non current inventory", "Trade receivables", 
  "Biological assets", "Other current assets", "Other non current assets", 
  "Property, plant and equipment", "Total assets", "Trade payables", 
  "Provisions for liabilities and charges", "Total liabilities", 
  "Share premium", "Treasury shares", 
  "Retained earnings / (Accumulated deficit)", "Other reserves", 
  "Total equity", "Total liabilities and equity"
)

#variables_financial_other <- list(
#  'Cash and cash equivalents', 'Inventories', 'Trade receivables',
#  'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
#  'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
#  'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
#  'Total equity', 'Total liabilities and equity', 'Cash advances made to other parties', 'Investment property',
#  'Investments in subsidiaries', 'Goodwill', 'Other intangible assets', 'Finance lease payable', 'Unearned income',
#  'Current borrowings', 'Non current borrowings', 'Received grants', 'Total current assets', 'Total current liabilities',
#  'Share capital'
#)

variables_financial_other <- list(
  "Cash and cash equivalents", "Inventories", "Trade receivables", 
  "Biological assets", "Other current assets", "Other non current assets", 
  "Property, plant and equipment", "Total assets", "Trade payables", 
  "Provisions for liabilities and charges", "Total liabilities", 
  "Share premium", "Treasury shares", 
  "Retained earnings / (Accumulated deficit)", "Other reserves", 
  "Total equity", "Total liabilities and equity"
)

"variables_profit_loss <- list(
  'Net Revenue', 'Cost of goods sold', 'Gross profit', 'Other operating income',
  'Personnel expense', 'Rental expenses', 'Depreciation and amortisation',
  'Other administrative and operating expenses', 'Operating income', 
  'Impairment (loss)/reversal of financial assets', 'Net gain (loss) from foreign exchange operations', 'Dividends received',
  'Other net operating income/(expense)', 'Profit/(loss) before tax from continuing operations',
  'Income tax', 'Profit/(loss)', 'Revaluation reserve of property, plant and equipment',
  'Other (include Share of associates and joint ventures in revaluation reserve of property, plant and equipment and defined benefit obligation)',
  'Total other comprehensive (loss) income', 'Total comprehensive income / (loss)'
)"

variables_profit_loss <- list(
  "Net Revenue", "Cost of goods sold", "Gross profit", 
  "Other operating income", "Personnel expense", "Rental expenses", 
  "Depreciation and amortisation", "Other administrative and operating expenses", 
  "Operating income", "Impairment (loss)/reversal of financial assets", 
  "Inventories", "Net gain (loss) from foreign exchange operations", 
  "Dividends received", "Other net operating income/(expense)", 
  "Profit/(loss) before tax from continuing operations", "Income tax", 
  "Profit/(loss)", 
  "Revaluation reserve of property, plant and equipment", 
  "Other (include Share of associates and joint ventures in revaluation reserve of property, plant and equipment and defined benefit obligation)", 
  "Total other comprehensive (loss) income", 
  "Total comprehensive income / (loss)"
)

variables_cash_flow <- list(
  'Net cash from operating activities', 'Net cash used in investing activities',
  'Net cash raised in financing activities', 'Net cash inflow for the year',
  'Effect of exchange rate changes on cash and cash equivalents',
  'Cash at the beginning of the year', 'Cash at the end of the year'
)


# Initialize lists to store adjusted data frames and data frames without 'LineItemENG'
data_list_adjusted <- list()
data_list_no_eng <- list()

# Function to correct and clean data frames that contain 'LineItemENG' column
# Adjusted to accept 'is_IV' parameter
# Function to correct and clean data frames that contain 'LineItemENG' column
correct_lineitems <- function(df, is_IV) {
  df <- df %>%
    # Exclude group III
    filter(CategoryMain != "III ჯგუფი") %>%
    #filter(if (!is_IV) FormName != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)" else TRUE) %>%
    filter(FormName != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") %>%
    # Standardize 'FormName' and 'SheetName'
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
      )
    ) %>%
    # Correct inconsistencies in 'LineItemENG' and 'LineItemGEO'
    mutate(
      LineItemENG = case_when(
        LineItemENG == "Retained earnings (Accumulated deficit)" ~ "Retained earnings / (Accumulated deficit)",
        LineItemENG == "Impairment loss/reversal of  financial assets" ~ "Impairment (loss)/reversal of financial assets",
        LineItemENG == "Total comprehensive income" ~ "Total comprehensive income / (loss)",
        LineItemENG == "Total comprehensive income(loss)" ~ "Total comprehensive income / (loss)",
        LineItemENG == "Prepayments" ~ "Cash advances made to other parties",
        LineItemENG == "Cash advances to other parties" ~ "Cash advances made to other parties",
        LineItemENG == 'Share capital (in case of Limited Liability Company - "capital", in case of cooperative entity - "unit capital"' ~ "Share capital",
        LineItemENG == "- inventories" ~ "Inventories",
        TRUE ~ LineItemENG
      ),
      LineItemGEO = case_when(
        LineItemGEO == "ამონაგები" ~ "ნეტო ამონაგები",
        LineItemGEO == "სხვა პირებზე ავანსებად და სესხებად გაცემული ფულადი სახსრები" ~ "სხვა მხარეებზე ავანსად გაცემული ფულადი სახსრები",
        LineItemGEO == "სხვა მხარეებზე ავანსებად გაცემული ფულადი სახსრები" ~ "სხვა მხარეებზე ავანსად გაცემული ფულადი სახსრები",
        LineItemGEO == "გაუფასურების (ხარჯი) / აღდგენა ფინანსურ აქტივებზე" ~ "ფინანსური აქტივების გაუფასურების (ხარჯი) / აღდგენა",
        LineItemGEO == "საემისიო კაპიტალი" ~ "საემისიო შემოსავალი",
        LineItemGEO == "მოგება / (ზარალი)" ~ "საანგარიშგებო პერიოდის მოგება / ზარალი",
        LineItemGEO == "მარაგების გაუფასურების (ხარჯი) / აღდგენა" ~ "მარაგები",
        LineItemGEO == "ნეტო ფულადი სახსრები საოპერაციო საქმიანობიდან" ~ "საოპერაციო საქმიანობიდან წარმოქმნილი (საოპერაციო საქმიანობაში გამოყენებული) წმინდა ფულადი ნაკადები",
        LineItemGEO == "საინვესტიციო საქმიანობაში გამოყენებული ნეტო ფულადი სახსრები" ~ "საინვესტიციო საქმიანობიდან წარმოქმნილი (საინვესტიციო საქმიანობაში გამოყენებული) წმინდა ფულადი ნაკადები",
        LineItemGEO == "ნეტო ფულადი სახსრები საფინანსო საქმიანობიდან" ~ "საფინანსო საქმიანობიდან წარმოქმნილი (საფინანსო საქმიანობაში გამოყენებული) წმინდა ფულადი ნაკადები",
        LineItemGEO == "ნეტო ფულადი სახსრების შემოსვლა ან (გასვლა) წლის განმავლობაში" ~ "ფულადი სახსრების შემოსვლა (გასვლა)",
        LineItemGEO == "სავალუტო კურსის ეფექტი" ~ "ვალუტის  კურსის ცვლილების გავლენა ფულად სახსრებსა და მათ ეკვივალენტებზე",
        LineItemGEO == "ფულადი სახსრები წლის დასაწყისისათვის" ~ "ფულადი სახსრები და მათი ეკვივალენტები საანგარიშგებო პერიოდის დასაწყისში",
        LineItemGEO == "ფულადი სახსრები წლის ბოლოს" ~ "ფულადი სახსრები და მათი ეკვივალენტები საანგარიშგებო პერიოდის ბოლოს",
        TRUE ~ LineItemGEO
      )
    ) %>%
    # Convert 'Value' to numeric
    mutate(Value = as.numeric(Value)) %>%
    # Adjust 'Value' based on 'GEL' column
    mutate(
      Value = case_when(
        GEL == ".000 ლარი" & !is.na(Value) ~ Value * 1000,
        TRUE ~ Value
      )
    )
  
  # Apply filtering based on 'SheetName' and appropriate variables
  if (is_IV) {
    # For Category IV, filter based only on 'SheetName' and relevant variables
    df <- df %>%
      filter(
        !(FormName == "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_non_financial)
      ) %>%
      filter(
        !(FormName != "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_other)
      ) %>%
      filter(
        !(SheetName == "profit-loss" & !LineItemENG %in% variables_profit_loss)
      ) %>%
      filter(
        !(SheetName == "cash flow" & !LineItemENG %in% variables_cash_flow)
      )
  } else {
    # For other categories, include 'FormName' and 'variables_financial_non_financial' in filters
    df <- df %>%
      filter(
        !(FormName == "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_non_financial)
      ) %>%
      filter(
        !(FormName != "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_other)
      ) %>%
      filter(
        !(SheetName == "profit-loss" & !LineItemENG %in% variables_profit_loss)
      ) %>%
      filter(
        !(SheetName == "cash flow" & !LineItemENG %in% variables_cash_flow)
      )
  }
  
  df <- df %>% arrange(ReportCode)
  
  return(df)
}


# Function to correct and clean data frames that do not contain 'LineItemENG' column
correct_geo_dfs <- function(df, is_IV) {
  df <- df %>%
    # Exclude group III
    filter(
      if ("Category" %in% colnames(df)) Category != "III ჯგუფი"
      else if ("CategoryMain" %in% colnames(df)) CategoryMain != "III ჯგუფი"
      else TRUE
    ) %>%
    # Exclude financial institutions if not Category IV
    filter(FormName != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") %>%
    # Standardize 'FormName' and 'SheetName'
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
      )
    ) %>%
    # Correct inconsistencies in 'LineItem' (Georgian names)
    mutate(
      LineItem = case_when(
        LineItem == "ამონაგები" ~ "ნეტო ამონაგები",
        LineItem == "სხვა პირებზე ავანსებად და სესხებად გაცემული ფულადი სახსრები" ~ "სხვა მხარეებზე ავანსად გაცემული ფულადი სახსრები",
        LineItem == "სხვა მხარეებზე ავანსად გაცემული ფულადი სახსრები" ~ "სხვა მხარეებზე ავანსად გაცემული ფულადი სახსრები",
        LineItem == "გაუფასურების (ხარჯი) / აღდგენა ფინანსურ აქტივებზე" ~ "ფინანსური აქტივების გაუფასურების (ხარჯი) / აღდგენა",
        LineItem == "საემისიო კაპიტალი" ~ "საემისიო შემოსავალი",
        LineItem == "მოგება / (ზარალი)" ~ "საანგარიშგებო პერიოდის მოგება / ზარალი",
        TRUE ~ LineItem
      )
    ) %>%
    # Convert 'Value' to numeric
    mutate(Value = as.numeric(Value)) %>%
    # Adjust 'Value' based on 'GEL' column
    mutate(
      Value = case_when(
        GEL == ".000 ლარი" & !is.na(Value) ~ Value * 1000,
        TRUE ~ Value
      )
    )
  
  
  return(df)
}


# Iterate over each data frame in 'data_list' and process accordingly
for (i in seq_along(data_list)) {
  # Access the current data frame
  df <- data_list[[i]] 
  
  # Get the name of the current data frame
  df_name <- names(data_list)[i]
  
  # Determine if the data frame name contains "IV" (case-insensitive)
  is_IV <- grepl("IV", df_name, ignore.case = TRUE)
  
  # Check if 'LineItemENG' exists in the current data frame
  if (!"LineItemENG" %in% colnames(df)) {
    # Correct the data frame and save it into 'data_list_no_eng', passing 'is_IV'
    df_geo_corrected <- correct_geo_dfs(df, is_IV)
    data_list_no_eng[[df_name]] <- df_geo_corrected
  } else {
    # Remove incorrect entries where 'LineItemENG' is 'Inventories' in 'profit-loss' sheet
    df <- df %>%
      filter(!(SheetName == "profit-loss" & LineItemENG == "Inventories"))
    
    # Correct the line items using the 'correct_lineitems' function with 'is_IV' parameter
    df_corrected <- correct_lineitems(df, is_IV)
    
    # Save the corrected data frame to 'data_list_adjusted'
    data_list_adjusted[[df_name]] <- df_corrected
    
    # Perform variable checking and print statements
    if (!is_IV) {
      # For non-Category IV data frames
      
      # Check and print variables for financial sections
      # Filter for non-financial institutions in 'financial position' sheet
      df_filtered_financial_non_financial <- df_corrected %>%
        filter(FormName == "non-financial institutions" & SheetName == "financial position")
      
      # Check if all required variables are present
      if (all(variables_financial_non_financial %in% df_filtered_financial_non_financial$LineItemENG)) {
        print(paste("All variables found in financial_non_financial for", df_name, ":", TRUE))
      } else {
        missing_vars <- setdiff(variables_financial_non_financial, df_filtered_financial_non_financial$LineItemENG)
        print(paste("Variables not found in financial_non_financial for", df_name, ":", 
                    paste(missing_vars, collapse = ", ")))
      }
      
      # Repeat the process for 'financial_other' variables
      df_filtered_financial_other <- df_corrected %>%
        filter(FormName != "non-financial institutions" & SheetName == "financial position")
      
      if (all(variables_financial_other %in% df_filtered_financial_other$LineItemENG)) {
        print(paste("All variables found in financial_other for", df_name, ":", TRUE))
      } else {
        missing_vars <- setdiff(variables_financial_other, df_filtered_financial_other$LineItemENG)
        print(paste("Variables not found in financial_other for", df_name, ":", 
                    paste(missing_vars, collapse = ", ")))
      }
    } else {
      # For Category IV data frames
      
      # Check variables only based on 'SheetName' and appropriate variables lists
      
      # Financial Position sheet
      df_filtered_financial_position <- df_corrected %>%
        filter(SheetName == "financial position")
      
      if (all(variables_financial_other %in% df_filtered_financial_position$LineItemENG)) {
        print(paste("All variables found in financial_position for", df_name, ":", TRUE))
      } else {
        missing_vars <- setdiff(variables_financial_other, df_filtered_financial_position$LineItemENG)
        print(paste("Variables not found in financial_position for", df_name, ":", 
                    paste(missing_vars, collapse = ", ")))
      }
    }
    
    # Check 'profit_loss' variables
    df_filtered_profit_loss <- df_corrected %>%
      filter(SheetName == "profit-loss")
    
    if (all(variables_profit_loss %in% df_filtered_profit_loss$LineItemENG)) {
      print(paste("All variables found in profit_loss for", df_name, ":", TRUE))
    } else {
      missing_vars <- setdiff(variables_profit_loss, df_filtered_profit_loss$LineItemENG)
      print(paste("Variables not found in profit_loss for", df_name, ":", 
                  paste(missing_vars, collapse = ", ")))
    }
    
    # Check 'cash_flow' variables
    df_filtered_cash_flow <- df_corrected %>%
      filter(SheetName == "cash flow")
    
    if (all(variables_cash_flow %in% df_filtered_cash_flow$LineItemENG)) {
      print(paste("All variables found in cash_flow for", df_name, ":", TRUE))
    } else {
      missing_vars <- setdiff(variables_cash_flow, df_filtered_cash_flow$LineItemENG)
      print(paste("Variables not found in cash_flow for", df_name, ":", 
                  paste(missing_vars, collapse = ", ")))
    }
  }
}

# Combine all variables into a single list for further processing
combined_variable_list <- union(
  union(variables_financial_non_financial, variables_financial_other),
  union(variables_profit_loss, variables_cash_flow)
)

# Function to return a list of 'LineItemENG' and corresponding unique 'LineItemGEO' values within a data frame
check_unique_geo_values <- function(df, variables_list) {
  result_list <- list()  # Initialize an empty list to store the found results
  missing_variables <- c()  # Initialize a vector to store missing variables
  
  # Iterate over each variable in the variables list
  for (var in variables_list) {
    if (var %in% df$LineItemENG) {
      # Filter rows where 'LineItemENG' matches the variable
      filtered_df <- df %>% filter(LineItemENG == var)
      # Get the unique 'LineItemGEO' values for this 'LineItemENG'
      unique_geo_values <- unique(filtered_df$LineItemGEO)
      # Store the result as a named list (LineItemENG and its corresponding LineItemGEO values)
      result_list[[var]] <- unique_geo_values
    } else {
      # If the variable is not found, add it to the missing variables list
      missing_variables <- c(missing_variables, var)
    }
  }
  
  # Return both found results and missing variables
  return(list(found = result_list, missing = missing_variables))
}

# Initialize an empty list to store the results for all data frames
all_results <- list()

# Initialize a list to collect unique FormName values for each data frame
form_names_by_df <- list()

# Iterate over each data frame in 'data_list_adjusted' and store the result
for (df_name in names(data_list_adjusted)) {
  if ("LineItemENG" %in% colnames(data_list_adjusted[[df_name]])) {
    df <- data_list_adjusted[[df_name]]
    
    # Call the function and store the result in a named list
    result <- check_unique_geo_values(df, combined_variable_list)
    
    # Collect unique FormName values for the current data frame
    form_names <- unique(df$FormName)
    form_names_by_df[[df_name]] <- form_names  # Store FormName values for each data frame
    
    # Store the result for each data frame in a named element of 'all_results'
    all_results[[df_name]] <- result
    
    # Output the found and missing variables for this data frame
    cat(paste("\nFor DataFrame", df_name, ":\n"))
    
    # Print found variables and their corresponding 'LineItemGEO'
    if (length(result$found) > 0) {
      cat("Found variables and their LineItemGEO values:\n")
      print(result$found)
      
      # Identify and print English variables with more than one Georgian pair
      for (eng_var in names(result$found)) {
        georgian_pairs <- result$found[[eng_var]]
        if (length(georgian_pairs) > 1) {
          cat(paste("English variable with multiple Georgian pairs:", eng_var, "\n"))
          cat("Georgian pairs:\n")
          print(georgian_pairs)
        }
      }
    }
    
    # Print missing variables
    if (length(result$missing) > 0) {
      cat("\nMissing variables from LineItemENG:\n")
      print(result$missing)
    } else {
      cat("All variables found in LineItemENG.\n")
    }
  }
}

# Add the FormName values for each data frame as a sublist to 'all_results'
all_results$FormNames_By_DataFrame <- form_names_by_df

# Function to check the consistency of 'LineItemENG'-'LineItemGEO' pairs across data frames
check_consistency_across_dataframes <- function(all_results) {
  # Get the list of data frame names
  df_names <- names(all_results)
  
  # Remove 'FormNames_By_DataFrame' from the list
  df_names <- df_names[df_names != "FormNames_By_DataFrame"]
  
  # Initialize the reference mapping using the first data frame's results
  reference_pairs <- all_results[[df_names[1]]]$found
  inconsistent_pairs <- list()  # To store inconsistencies found
  
  # Iterate through each subsequent data frame's results and compare
  for (i in 2:length(df_names)) {
    current_pairs <- all_results[[df_names[i]]]$found
    
    # Compare each 'LineItemENG' in the reference with the current data frame
    for (line_item in names(reference_pairs)) {
      if (line_item %in% names(current_pairs)) {
        # Check if the 'LineItemGEO' values are the same
        if (!identical(reference_pairs[[line_item]], current_pairs[[line_item]])) {
          # Record the inconsistency
          inconsistent_pairs[[paste("DataFrame", df_names[i], "LineItem:", line_item)]] <- list(
            reference = reference_pairs[[line_item]],
            current = current_pairs[[line_item]]
          )
        }
      }
    }
  }
  
  # Print inconsistent pairs if found
  if (length(inconsistent_pairs) > 0) {
    cat("\nInconsistent LineItemENG-LineItemGEO pairs found across data frames:\n")
    print(inconsistent_pairs)
  } else {
    cat("\nAll LineItemENG-LineItemGEO pairs are consistent across data frames.\n")
  }
}

# After processing all results, check for consistency
check_consistency_across_dataframes(all_results)




# Create a lookup table from the first data frame's results
lookup_table <- data.frame(
  English = names(all_results$`DataFrame 1`$found),      # English 'LineItemENG' names
  Georgian = unlist(all_results$`DataFrame 1`$found)     # Corresponding Georgian 'LineItemGEO' values
)

lookup_table <- data.frame(
  English = names(all_results[[1]]$found),
  Georgian = unlist(all_results[[1]]$found)
)

# Reset row names to default
rownames(lookup_table) <- NULL

#Check that Georgian values do not have more than one english pair
check_georgian_lead_consistency <- function(data_list_adjusted, lookup_table) {
  # Initialize a list to store inconsistencies
  inconsistencies <- list()
  
  # Iterate through each Georgian variable in the lookup table
  for (geo_variable in unique(lookup_table$Georgian)) {
    # Track English variables corresponding to the Georgian variable across data frames
    english_mapping <- list()
    
    # Iterate through each data frame in the list
    for (df_name in names(data_list_adjusted)) {
      # Extract the data frame
      df <- data_list_adjusted[[df_name]]
      
      # Check if the Georgian variable (LineItemGEO) exists in the data frame
      if (geo_variable %in% df$LineItemGEO) {
        # Find corresponding English variables (LineItemENG)
        eng_values <- unique(df$LineItemENG[df$LineItemGEO == geo_variable])
        english_mapping[[df_name]] <- eng_values
      }
    }
    
    # Flatten and check uniqueness of English variables
    all_eng_values <- unlist(english_mapping)
    unique_eng_values <- unique(all_eng_values)
    
    # If the Georgian variable maps to more than one unique English variable, record it
    if (length(unique_eng_values) > 1) {
      inconsistencies[[geo_variable]] <- list(
        DataFrames = names(english_mapping),
        EnglishVariables = english_mapping
      )
    }
  }
  
  # Print inconsistencies if found
  if (length(inconsistencies) > 0) {
    cat("\nInconsistent Georgian variables found:\n")
    for (geo_var in names(inconsistencies)) {
      cat("\nGeorgian Variable:", geo_var, "\n")
      print(inconsistencies[[geo_var]])
    }
  } else {
    cat("\nAll Georgian variables have consistent English mappings across data frames.\n")
  }
}

check_georgian_lead_consistency(data_list_adjusted, lookup_table)


# Function to make data frames uniform by adding 'LineItemENG' based on 'LineItemGEO' and lookup table
make_identical <- function(df, lookup_table) {
  # Rename columns for consistency
  df <- df %>%
    rename(
      LineItemGEO = any_of("LineItem"),
      CategoryMain = any_of("Category")
    )
  
  # Filter out values in 'LineItemGEO' that are not in the lookup table before the join
  df <- df %>%
    filter(LineItemGEO %in% lookup_table$Georgian)
  
  # Check for missing Georgian variables
  missing_georgian_vars <- setdiff(lookup_table$Georgian, df$LineItemGEO)
  
  if (length(missing_georgian_vars) > 0) {
    cat("Warning: The following Georgian variables were not found in the data frame:\n")
    print(missing_georgian_vars)
  } else {
    cat("All Georgian variables from the lookup table are found in the data frame.\n")
  }
  
  # If 'LineItemENG' does not exist, add it by joining with the lookup table
  if (!"LineItemENG" %in% colnames(df)) {
    df <- df %>%
      left_join(lookup_table, by = c("LineItemGEO" = "Georgian")) %>%
      rename(LineItemENG = English)
  }
  
  return(df)
}

# Apply the function to data frames in 'data_list_no_eng'
data_list_no_eng_adjusted <- lapply(data_list_no_eng, make_identical, lookup_table = lookup_table)

# Combine the two lists into one
combined_data_list <- c(data_list_adjusted, data_list_no_eng_adjusted)

# Ensure unique names for data frames in the combined list
names(combined_data_list) <- make.unique(c(names(data_list_adjusted), names(data_list_no_eng_adjusted)))

# Convert 'Value' to numeric in each data frame
combined_data_list <- lapply(combined_data_list, function(df) {
  df <- df %>%
    mutate(Value = as.numeric(Value))  # Convert 'Value' to numeric
  return(df)
})

# Function to make column types consistent across data frames
standardize_column_types <- function(data_list) {
  # Standardize 'ReportId' as character (or any column causing issues)
  data_list <- lapply(data_list, function(df) {
    if ("ReportId" %in% colnames(df)) {
      df <- df %>%
        mutate(ReportId = as.character(ReportId))  # Convert 'ReportId' to character
    }
    return(df)
  })
  return(data_list)
}

# Apply the function to standardize column types
combined_data_list <- standardize_column_types(combined_data_list)

'inconsistent_dfs <- lapply(combined_data_list, function(df) {
  if ("IdCode" %in% colnames(df) && class(df$IdCode) != "character") {
    return(names(df))
  }
})
print(inconsistent_dfs)'



# Combine the list of data frames into one data frame
combined_df <- bind_rows(combined_data_list)


# Function to group and split data frames by 'ReportCode' and then by 'LineItemENG'
nested_list_of_dfs_group_split <- function(df) {
  # Step 1: Group and split by 'ReportCode'
  list_of_dfs_group_split_by_report_code <- df %>% group_split(ReportCode)
  
  # Step 2: For each data frame from the 'ReportCode' split, group and split by 'LineItemENG'
  nested_list_of_dfs_group_split_by_lineitemeng <- lapply(list_of_dfs_group_split_by_report_code, function(df_grouped) {
    df_grouped %>% group_split(LineItemENG)
  })
  
  # Return the nested list of data frames grouped by 'LineItemENG' within 'ReportCode'
  return(nested_list_of_dfs_group_split_by_lineitemeng)
}

# Apply the grouping function to the combined data frame
nested_list_of_dfs_group_split <- nested_list_of_dfs_group_split(combined_df)  

# Function to process each grouped data frame and handle multiple entries
process_df_secondary <- function(df) {
  processed_year <- list()
  column_to_check <- 'ReportYear'
  instances_over_two <- 0
  
  cat("Processing data frame with", nrow(df), "rows\n")
  
  for (i in 1:nrow(df)) {
    current_value <- df$FVYear[i]
    
    if (!(current_value %in% processed_year)) {
      processed_year <- append(processed_year, current_value)
      found_matches <- which(df$FVYear == current_value)
      cat("Found matches for FVYear =", current_value, ":", paste(found_matches, collapse = ", "), "\n")
      
      if (length(found_matches) > 1) {
        cat("Multiple rows for FVYear", current_value, "\n")
        
        if (length(found_matches) > 2) {
          instances_over_two <- instances_over_two + 1
        }
        
        # Extract the column data and check if it's numeric
        column_data <- df[[column_to_check]][found_matches]
        value_data <- df[["Value"]][found_matches]
        
        # Convert column data to numeric, handling any non-numeric values
        column_data_numeric <- suppressWarnings(as.numeric(as.character(column_data)))
        if (any(is.na(column_data_numeric))) {
          cat("Warning: Non-numeric values found in ReportYear. Treating as NA:\n", paste(column_data, collapse = ", "), "\n")
          column_data_numeric <- ifelse(is.na(column_data_numeric), -Inf, column_data_numeric) # Handle NAs by assigning a very low value
        }
        
        # Sort column data and value data in decreasing order
        sorted_indices <- order(column_data_numeric, decreasing = TRUE)
        column_data_numeric <- column_data_numeric[sorted_indices]
        value_data <- value_data[sorted_indices]
        
        cat("Column data after sorting:", paste(column_data_numeric, collapse = ", "), "\n")
        cat("Value data after sorting:", paste(value_data, collapse = ", "), "\n")
        
        # Find the first non-zero value in the sorted order
        non_zero_index <- which(value_data != 0)
        if (length(non_zero_index) > 0) {
          # Use the first non-zero value
          max_value_index <- found_matches[sorted_indices[non_zero_index[1]]]
          cat("Non-zero value found at index:", max_value_index, "\n")
        } else {
          # All values are zero, use the latest year value
          max_value_index <- found_matches[sorted_indices[1]]
          cat("All zero values, keeping latest year at index:", max_value_index, "\n")
        }
        
        # Remove rows that are not the 'max_value_index'
        rows_to_remove <- found_matches[found_matches != max_value_index]
        
        if (length(rows_to_remove) > 0) {
          # Debugging: Print the indices to be removed
          cat("Dropping rows with indices:", paste(rows_to_remove, collapse = ", "), "\n")
          df <- df[-rows_to_remove, ]
        }
      }
    }
  }
  
  # Debugging: Print the number of rows after processing
  cat("Number of rows after processing:", nrow(df), "\n")
  cat("Number of instances with more than two matches:", instances_over_two, "\n")
  
  return(df)
}

# Apply the processing function to each grouped data frame
final_processed_list <- lapply(nested_list_of_dfs_group_split, function(inner_list){
  lapply(inner_list, function(df) process_df_secondary(df))
})

# Flatten the nested list into a single list of data frames
flattened_final_processed_list <- unlist(final_processed_list, recursive = FALSE)

# Function to transform the data to wide format
transform_to_wide_format <- function(df) {
  # Ensure 'Value' is numeric
  df <- df %>%
    mutate(Value = as.numeric(Value))
  
  # Group by 'IdCode', 'FVYear', and 'LineItemENG', and sum 'Value'
  df_grouped <- df %>%
    group_by(IdCode, FVYear, LineItemENG) %>%
    summarise(Value = sum(Value, na.rm = TRUE), .groups = 'drop')
  
  # Transform the data frame from long to wide format using 'pivot_wider'
  df_wide <- df_grouped %>%
    pivot_wider(names_from = LineItemENG, values_from = Value)
  
  # Reorder columns to have 'IdCode' and 'FVYear' at the front
  df_wide <- df_wide %>%
    select(IdCode, FVYear, everything())
  
  return(df_wide)
}

# Combine the processed data frames into a single data frame
combined_df_processed <- bind_rows(flattened_final_processed_list)

# Transform the combined data frame to wide format
final_wide_df <- transform_to_wide_format(combined_df_processed)

# Read beneficiaries data from an Excel file
benefitiaries_path <- "benefitiaries_data.xlsx"
beneficiaries_df <- read_excel(benefitiaries_path)

# Rename columns in the beneficiaries data frame for consistency
beneficiaries_df <- beneficiaries_df %>%
  rename(
    IdCode = `ს/კ`,           # Registration code
    ReportCode = `რეპორტ კოდი`,  # Report code
    Program = პროგრამა          # Program
  ) %>%
  # Convert 'Program' to numeric codes
  mutate(
    Program = case_when(
      Program == "ინდუსტრიული" ~ 1,
      Program == "უნივერსალური" ~ 2,
      Program == "საკრედიტო-საგარანტიო" ~ 3,
      Program == "ორივე პროგრამით სარგებლობა" ~ 4,
      TRUE ~ NA_real_ 
    )
  )

# Ensure 'IdCode' columns are of the same type in both data frames
beneficiaries_df$IdCode <- as.character(beneficiaries_df$IdCode)
final_wide_df$IdCode <- as.character(final_wide_df$IdCode)

# Identify 'IdCode's in 'beneficiaries_df' that are present in 'final_wide_df'
matched_idcodes <- beneficiaries_df %>%
  filter(IdCode %in% final_wide_df$IdCode)

# Identify 'IdCode's in 'beneficiaries_df' that are not present in 'final_wide_df'
unmatched_idcodes <- beneficiaries_df %>%
  filter(!IdCode %in% final_wide_df$IdCode)

# Combine the 'Program' values for each 'IdCode' by collapsing them into a single string
beneficiaries_collapsed <- matched_idcodes %>%
  group_by(IdCode) %>%
  summarise(ProgramBeneficiary = paste(unique(Program), collapse = ","), .groups = "drop")

# Perform the left join with 'final_wide_df'
final_wide_df_with_program <- final_wide_df %>%
  left_join(beneficiaries_collapsed, by = "IdCode")

# Move 'ProgramBeneficiary' to the front
final_wide_df_with_program <- final_wide_df_with_program %>%
  select(ProgramBeneficiary, everything())

# Output the number of matched and unmatched 'IdCode's
cat("Number of IdCodes in beneficiaries_df:", nrow(beneficiaries_df), "\n")
cat("Number of IdCodes matched in final_wide_df:", nrow(matched_idcodes), "\n")
cat("Number of IdCodes not matched:", nrow(unmatched_idcodes), "\n")

# Optional: View the unmatched 'IdCode's
print("Unmatched IdCodes:")
print(unmatched_idcodes)

# Function to create new variables based on two existing columns and an operator
create_new_variables <- function(df, column1, column2, new_column_name, operator){
  if (all(c(column1, column2) %in% colnames(df))) {
    df <- df %>%
      mutate(
        !!new_column_name := case_when(
          operator == "+" ~ ifelse(is.na(.data[[column1]]) | is.na(.data[[column2]]), 
                                   NA, 
                                   .data[[column1]] + .data[[column2]]),
          operator == "-" ~ ifelse(is.na(.data[[column1]]) | is.na(.data[[column2]]), 
                                   NA, 
                                   .data[[column1]] - .data[[column2]]),
          operator == "*" ~ ifelse(is.na(.data[[column1]]) | is.na(.data[[column2]]), 
                                   NA, 
                                   .data[[column1]] * .data[[column2]]),
          operator == "/" ~ ifelse(is.na(.data[[column1]]) | is.na(.data[[column2]]) | .data[[column2]] == 0, 
                                   NA, 
                                   .data[[column1]] / .data[[column2]]),
          TRUE ~ NA_real_ 
        )
      )
  } else {
    cat("One or both of the specified columns do not exist in the data frame.\n")
  }
  
  return(df)
}

# Create new variables using the 'create_new_variables' function
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Profit/(loss)", "Net Revenue", "Margin", "/")
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Total liabilities", "Total assets", "Liabilities to Assets", "/")
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Current borrowings", "Non current borrowings", "Total Borrowings", "+")
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Total Borrowings", "Total assets", "Borrowings to Assets", "/")
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Cash and cash equivalents", "Total assets", "Cash to Assets", "/")
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Operating income", "Total assets", "Operating income to Assets", "/")
final_wide_df_with_program <- create_new_variables(final_wide_df_with_program, "Total liabilities", "Operating income", "Liabilities to Operating income", "/")

# Function to calculate statistics by year for specified columns
calculate_statistics_by_year <- function(df, columns) {
  summary_list <- list()
  
  for (column_name in columns) {
    if (column_name %in% colnames(df)) {
      summary_stats <- df %>%
        group_by(FVYear) %>%
        summarise(
          Mean = mean(.data[[column_name]], na.rm = TRUE),
          Median = median(.data[[column_name]], na.rm = TRUE),
          Percentile_1 = quantile(.data[[column_name]], 0.01, na.rm = TRUE),
          Percentile_5 = quantile(.data[[column_name]], 0.05, na.rm = TRUE),
          Percentile_10 = quantile(.data[[column_name]], 0.1, na.rm = TRUE),
          Percentile_25 = quantile(.data[[column_name]], 0.25, na.rm = TRUE),
          Percentile_50 = quantile(.data[[column_name]], 0.5, na.rm = TRUE),
          Percentile_75 = quantile(.data[[column_name]], 0.75, na.rm = TRUE),
          Percentile_90 = quantile(.data[[column_name]], 0.9, na.rm = TRUE),
          Percentile_95 = quantile(.data[[column_name]], 0.95, na.rm = TRUE),
          Percentile_99 = quantile(.data[[column_name]], 0.99, na.rm = TRUE)
        ) %>%
        mutate(Column = column_name)  # Add a column to identify the column name
      
      cat("\nSummary statistics for", column_name, "grouped by FVYear:\n")
      print(summary_stats)
      
      summary_list[[column_name]] <- summary_stats
    } else {
      cat(paste("Column", column_name, "does not exist in the data frame.\n"))
    }
  }
  
  combined_summary_df <- bind_rows(summary_list)
  
  return(combined_summary_df)
}

# Function to drop extreme percentiles for a specified column
drop_percentiles_for_column <- function(df, column_name) {
  # Calculate the 1st and 99th percentiles for the specified column 
  lower_bound <- quantile(df[[column_name]], 0.01, na.rm = TRUE)
  upper_bound <- quantile(df[[column_name]], 0.99, na.rm = TRUE)
  
  # Filter the data frame based on these percentiles for the specific column
  df_filtered <- df %>%
    filter(df[[column_name]] >= lower_bound & df[[column_name]] <= upper_bound)
  
  return(df_filtered)
}

# Process 'Margin' variable
df_wide_margin <- final_wide_df_with_program %>%
  filter(Margin > -1 & Margin < 1)

df_wide_margin_beneficiaries <- df_wide_margin %>%
  filter(!is.na(ProgramBeneficiary))

df_wide_margin_non_beneficiaries <- df_wide_margin %>%
  filter(is.na(ProgramBeneficiary))

summary_margin_beneficiaries <- calculate_statistics_by_year(df_wide_margin_beneficiaries, "Margin")
summary_margin_non_beneficiaries <- calculate_statistics_by_year(df_wide_margin_non_beneficiaries, "Margin")

# Process 'Liabilities to Operating income' variable
df_wide_liabilities_to_operating <- drop_percentiles_for_column(final_wide_df_with_program, "Liabilities to Operating income")

df_wide_liabilities_to_operating_beneficiaries <- df_wide_liabilities_to_operating %>%
  filter(!is.na(ProgramBeneficiary))

df_wide_liabilities_to_operating_non_beneficiaries <- df_wide_liabilities_to_operating %>%
  filter(is.na(ProgramBeneficiary))

summary_liabilities_to_operating_beneficiaries <- calculate_statistics_by_year(df_wide_liabilities_to_operating_beneficiaries, "Liabilities to Operating income")
summary_liabilities_to_operating_non_beneficiaries <- calculate_statistics_by_year(df_wide_liabilities_to_operating_non_beneficiaries, "Liabilities to Operating income")

# Process 'Operating income to Assets' variable
df_wide_operating_to_assets <- drop_percentiles_for_column(final_wide_df_with_program, "Operating income to Assets")

df_wide_operating_to_assets_beneficiaries <- df_wide_operating_to_assets %>%
  filter(!is.na(ProgramBeneficiary))

df_wide_operating_to_assets_non_beneficiaries <- df_wide_operating_to_assets %>%
  filter(is.na(ProgramBeneficiary))

summary_operating_to_assets_beneficiaries <- calculate_statistics_by_year(df_wide_operating_to_assets_beneficiaries, "Operating income to Assets")
summary_operating_to_assets_non_beneficiaries <- calculate_statistics_by_year(df_wide_operating_to_assets_non_beneficiaries, "Operating income to Assets")

# Split main data frame into beneficiaries and non-beneficiaries
final_wide_df_with_program_beneficiaries <- final_wide_df_with_program %>%
  filter(!is.na(ProgramBeneficiary))

final_wide_df_with_program_non_beneficiaries <- final_wide_df_with_program %>%
  filter(is.na(ProgramBeneficiary))

# Columns that did not need processing
unprocessed_columns <- c("Liabilities to Assets", "Borrowings to Assets", "Cash to Assets")

# Calculate summary statistics for unprocessed columns
summary_for_unprocessed_beneficiaries <- calculate_statistics_by_year(final_wide_df_with_program_beneficiaries, unprocessed_columns)
summary_for_unprocessed_non_beneficiaries <- calculate_statistics_by_year(final_wide_df_with_program_non_beneficiaries, unprocessed_columns)

# Combine summary statistics
summary_beneficiaries <- bind_rows(
  summary_margin_beneficiaries,
  summary_liabilities_to_operating_beneficiaries,
  summary_operating_to_assets_beneficiaries,
  summary_for_unprocessed_beneficiaries
)

summary_non_beneficiaries <- bind_rows(
  summary_margin_non_beneficiaries,
  summary_liabilities_to_operating_non_beneficiaries,
  summary_operating_to_assets_non_beneficiaries,
  summary_for_unprocessed_non_beneficiaries
)

# Unique beneficiaries
unique_idcodes_non_beneficiaries <- final_wide_df_with_program %>%
  filter(is.na(ProgramBeneficiary)) %>%  # Filter rows where ProgramBeneficiary is NA
  select(IdCode) %>%                     # Select the IdCode column
  distinct() 

unique_idcodes_beneficiaries <- final_wide_df_with_program %>%
  filter(!is.na(ProgramBeneficiary)) %>%  # Filter rows where ProgramBeneficiary is not NA
  select(IdCode) %>%                     # Select the IdCode column
  distinct() 

cat("Number of beneficiaries :", length(unique_idcodes_beneficiaries$IdCode), "\n")
cat("Number of non beneficiaries :", length(unique_idcodes_non_beneficiaries$IdCode), "\n")

# Write the summary statistics and final data to Excel and CSV files
# Ensure the 'final' directory exists
if (!dir.exists("final")) {
  dir.create("final")
}

write_xlsx(unique_idcodes_beneficiaries, "final/unique_beneficiaries.xlsx")
write_xlsx(unique_idcodes_non_beneficiaries, "final/unique_non_beneficiaries.xlsx")


write_xlsx(summary_beneficiaries, "final/summary_beneficiaries.xlsx")
write.csv(summary_beneficiaries, "final/summary_beneficiaries.csv", row.names = FALSE)

write_xlsx(summary_non_beneficiaries, "final/summary_non_beneficiaries.xlsx")
write.csv(summary_non_beneficiaries, "final/summary_non_beneficiaries.csv", row.names = FALSE)

write_xlsx(final_wide_df_with_program, "final/final_data.xlsx")
write.csv(final_wide_df_with_program, "final/final_data.csv", row.names = FALSE)
