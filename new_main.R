library(readxl)
library(dplyr)
library(rstudioapi)
#library(jsonlite)
library(tidyr)

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

excel_files <- list.files(file.path(script_dir, "data_test"), pattern = "\\.xlsx$", full.names = TRUE)

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

# Filtering criteria
columns_to_keep_1st_revision <- c("ReportCode", "IdCode", "ReportYear", "FVYear", "CategoryMain", "FormName",
                                  "SheetName", "LineItemGEO", "LineItemENG", "Value", "GEL", "LineItem")

# Define necessary variables lists
variables_financial_non_financial <- list('Cash and cash equivalents', 'Current Inventory', 'Non current inventory', 'Trade receivables',
                                          'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
                                          'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
                                          'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
                                          'Total equity', 'Total liabilities and equity', 'Cash advances made to other parties', 'Investment property',
                                          'Investments in subsidiaries', 'Goodwill', 'Other intangible assets', 'Finance lease payable', 'Unearned income',
                                          'Current borrowings', 'Non current borrowings', 'Received grants', 'Total current assets', 'Total current liabilities',
                                          'Share capital')

variables_financial_other <- list('Cash and cash equivalents', 'Inventories', 'Trade receivables',
                                  'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
                                  'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
                                  'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
                                  'Total equity', 'Total liabilities and equity', 'Cash advances made to other parties', 'Investment property',
                                  'Investments in subsidiaries', 'Goodwill', 'Other intangible assets', 'Finance lease payable', 'Unearned income',
                                  'Current borrowings', 'Non current borrowings', 'Received grants', 'Total current assets', 'Total current liabilities',
                                  'Share capital')

variables_profit_loss <- list('Net Revenue', 'Cost of goods sold', 'Gross profit', 'Other operating income',
                              'Personnel expense', 'Rental expenses', 'Depreciation and amortisation',
                              'Other administrative and operating expenses', 'Operating income', 
                              'Impairment (loss)/reversal of financial assets', 'Net gain (loss) from foreign exchange operations', 'Dividends received',
                              'Other net operating income/(expense)', 'Profit/(loss) before tax from continuing operations',
                              'Income tax', 'Profit/(loss)', 'Revaluation reserve of property, plant and equipment',
                              'Other (include Share of associates and joint ventures in revaluation reserve of property, plant and equipment and defined benefit obligation)',
                              'Total other comprehensive (loss) income', 'Total comprehensive income / (loss)')

variables_cash_flow <- list('Net cash from operating activities', 'Net cash used in investing activities',
                            'Net cash raised in financing activities', 'Net cash inflow for the year',
                            'Effect of exchange rate changes on cash and cash equivalents',
                            'Cash at the beginning of the year', 'Cash at the end of the year')

## Initialize a new list to store the corrected dataframes
data_list_adjusted <- list()

# Define the correct_lineitems function
correct_lineitems <- function(df) {
  df <- df %>%
    # Apply the initial filtering
    #select(all_of(columns_to_keep_1st_revision)) %>%
    filter(CategoryMain != "III ჯგუფი") %>%
    filter(FormName != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") %>%
    
    # Apply the corrections to FormName and SheetName
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
    
    # Apply the corrections to LineItemENG and LineItemGEO
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
        LineItemGEO == "სხვა პირებზე ავანსებად და სესხებად გაცემული ფულადი სახსრები" ~ "სხვა მხარეებზე ავანსებად გაცემული ფულადი სახსრები",
        TRUE ~ LineItemGEO
      )
    ) %>%
    
    # Convert 'Value' to numeric if it exists
    mutate(Value = if ("Value" %in% colnames(df)) as.numeric(Value) else Value) %>%
    
    # Adjust 'Value' for specific cases
    mutate(
      Value = case_when(
        GEL == ".000 ლარი" & !is.na(Value) ~ Value * 1000,
        TRUE ~ Value
      )
    ) %>%
    
    # Apply filtering based on FormName, SheetName, and LineItemENG
    filter(!(FormName == "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_non_financial)) %>%
    filter(!(FormName != "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_other)) %>%
    filter(!(SheetName == "profit-loss" & !LineItemENG %in% variables_profit_loss)) %>%
    filter(!(SheetName == "cash flow" & !LineItemENG %in% variables_cash_flow)) %>%
    
    # Arrange the dataframe by ReportCode
    arrange(ReportCode)
  
  return(df)
}

# Initialize two lists: one for adjusted dataframes, one for those without LineItemENG
data_list_adjusted <- list()
data_list_no_eng <- list()

# Iterate over each dataframe in the data_list
for (i in seq_along(data_list)) {
  
  # Access the current dataframe
  df <- data_list[[i]] 
  
  # Check if 'LineItemENG' exists in the current dataframe
  if (!"LineItemENG" %in% colnames(df)) {
    # Save the dataframe to 'data_list_no_eng' if it does not have 'LineItemENG'
    data_list_no_eng[[names(data_list)[i]]] <- df
  } else {
    # Step 1: Apply the specific filter only if LineItemENG exists
    df <- df %>%
      filter(!(SheetName == "profit-loss" & LineItemENG == "Inventories"))
    
    # Step 2: Correct the line items if 'LineItemENG' exists
    df_corrected <- correct_lineitems(df)
    
    # Step 3: Save the corrected dataframe to 'data_list_adjusted'
    data_list_adjusted[[names(data_list)[i]]] <- df_corrected
    
    # Step 4: Check and print variables for financial sections
    df_filtered_financial_non_financial <- df_corrected %>%
      filter(FormName == "non-financial institutions" & SheetName == "financial position")
    
    if (all(variables_financial_non_financial %in% df_filtered_financial_non_financial$LineItemENG) == TRUE) {
      print(paste("All found in financial_non_financial for", names(data_list)[i], ":", TRUE))
    } else {
      print(paste("Variables not found in financial_non_financial for", names(data_list)[i], ":", 
                  setdiff(variables_financial_non_financial, df_filtered_financial_non_financial$LineItemENG)))
    }
    
    # Step 5: Filter for financial_other and print result
    df_filtered_financial_other <- df_corrected %>%
      filter(FormName != "non-financial institutions" & SheetName == "financial position")
    
    if (all(variables_financial_other %in% df_filtered_financial_other$LineItemENG) == TRUE) {
      print(paste("All found in financial_other for", names(data_list)[i], ":", TRUE))
    } else {
      print(paste("Variables not found in financial_other for", names(data_list)[i], ":", 
                  setdiff(variables_financial_other, df_filtered_financial_other$LineItemENG)))
    }
    
    # Step 6: Filter for profit_loss and print result
    df_filtered_profit_loss <- df_corrected %>%
      filter(SheetName == "profit-loss")
    
    if (all(variables_profit_loss %in% df_filtered_profit_loss$LineItemENG) == TRUE) {
      print(paste("All found in profit_loss for", names(data_list)[i], ":", TRUE))
    } else {
      print(paste("Variables not found in profit_loss for", names(data_list)[i], ":", 
                  setdiff(variables_profit_loss, df_filtered_profit_loss$LineItemENG)))
    }
    
    # Step 7: Filter for cash_flow and print result
    df_filtered_cash_flow <- df_corrected %>%
      filter(SheetName == "cash flow")
    
    if (all(variables_cash_flow %in% df_filtered_cash_flow$LineItemENG) == TRUE) {
      print(paste("All found in cash_flow for", names(data_list)[i], ":", TRUE))
    } else {
      print(paste("Variables not found in cash_flow for", names(data_list)[i], ":", 
                  setdiff(variables_cash_flow, df_filtered_cash_flow$LineItemENG)))
    }
  }
}

combined_variable_list <- union(
  union(variables_financial_non_financial, variables_financial_other),
  union(variables_profit_loss, variables_cash_flow)
)

# Function to return a list of LineItemENG and corresponding unique LineItemGEO values
check_unique_geo_values <- function(df, variables_list) {
  result_list <- list()  # Initialize an empty list to store the found results
  missing_variables <- c()  # Initialize a vector to store missing variables
  
  for (var in variables_list) {
    if (var %in% df$LineItemENG) {
      # Filter rows where LineItemENG matches the variable
      filtered_df <- df %>% filter(LineItemENG == var)
      
      # Get the unique LineItemGEO values
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

# Initialize an empty list to store the results for all dataframes
all_results <- list()

# Iterate over each dataframe in data_list_adjusted and store the result
for (i in seq_along(data_list_adjusted)) {
  if ("LineItemENG" %in% colnames(data_list_adjusted[[i]])) {
    df <- data_list_adjusted[[i]]
    
    # Call the function and store the result in a named list
    result <- check_unique_geo_values(df, combined_variable_list)
    
    # Store the result for each dataframe in a named element of all_results
    all_results[[paste("DataFrame", i)]] <- result
    
    # Output the found and missing variables for this dataframe
    cat(paste("\nFor DataFrame", i, ":\n"))
    
    # Print found variables and their corresponding LineItemGEO
    if (length(result$found) > 0) {
      cat("Found variables and their LineItemGEO values:\n")
      print(result$found)
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

# Function to check the consistency of LineItemENG-LineItemGEO pairs across dataframes
check_consistency_across_dataframes <- function(all_results) {
  # Initialize the reference mapping using the first dataframe's results
  reference_pairs <- all_results[[1]]$found
  inconsistent_pairs <- list()  # To store inconsistencies found
  
  # Iterate through each subsequent dataframe's results and compare
  for (i in 2:length(all_results)) {
    current_pairs <- all_results[[i]]$found
    
    # Compare each LineItemENG in the reference with the current dataframe
    for (line_item in names(reference_pairs)) {
      if (line_item %in% names(current_pairs)) {
        # Check if the LineItemGEO values are the same
        if (!identical(reference_pairs[[line_item]], current_pairs[[line_item]])) {
          # Record the inconsistency
          inconsistent_pairs[[paste("DataFrame", i, "LineItem:", line_item)]] <- list(
            reference = reference_pairs[[line_item]],
            current = current_pairs[[line_item]]
          )
        }
      }
    }
  }
  
  # Print inconsistent pairs if found
  if (length(inconsistent_pairs) > 0) {
    cat("\nInconsistent LineItemENG-LineItemGEO pairs found across dataframes:\n")
    print(inconsistent_pairs)
  } else {
    cat("\nAll LineItemENG-LineItemGEO pairs are consistent across dataframes.\n")
  }
}

# After processing all results, check for consistency
check_consistency_across_dataframes(all_results)



#Load corresponding geo-eng lineitems

##Convert JSON list to dataframe making sure both English and Georgian names are put in columns (and not in column names)

lookup_table <- data.frame(
  English = names(all_results$`DataFrame 1`$found),      # English LineItemENG names
  Georgian = unlist(all_results$`DataFrame 1`$found)     # Corresponding Georgian LineItemGEO values
)

###Reset row names to null
rownames(lookup_table) <- NULL

#Make dataframes uniform

make_identical <- function(df, lookup_table) {
  
  # Rename columns
  df <- df %>%
    rename(
      LineItemGEO = any_of("LineItem"),
      CategoryMain = any_of("Category")
    )
  
  # Filter out values in 'LineItemGEO' that are not in the lookup table before the join
  df <- df %>%
    filter(LineItemGEO %in% lookup_table$Georgian)
  
  # Check if 'LineItemENG' already exists, and only add it if it doesn't
  if (!"LineItemENG" %in% colnames(df)) {
    df <- df %>%
      left_join(lookup_table, by = c("LineItemGEO" = "Georgian")) %>%
      rename(LineItemENG = English)
  }
  
  return(df)
}

# Apply the function to dataframes in data_list_no_eng
data_list_no_eng_adjusted <- lapply(data_list_no_eng, make_identical, lookup_table = lookup_table)

# Combine the two lists into one
combined_data_list <- c(data_list_adjusted, data_list_no_eng_adjusted)


names(combined_data_list) <- make.unique(c(names(data_list_adjusted), names(data_list_no_eng_adjusted)))

#####უნდა გავაერთიანო

combined_data_list <- lapply(combined_data_list, function(df) {
  df <- df %>%
    mutate(Value = as.numeric(Value))  # Convert Value to numeric
  return(df)
})

# Combine the list of dataframes into one dataframe
combined_df <- bind_rows(combined_data_list)

####gadavamowmot rom gaetianebulshi yvelaferi sworia


####

# Apply group_split on each dataframe in combined_data_list, first by ReportCode, then by LineItemENG


nested_list_of_dfs_group_split <- function(df) {
  
  # Step 1: Group and split by ReportCode
  list_of_dfs_group_split_by_report_code <- df %>% group_split(ReportCode)
  
  # Step 2: For each dataframe from the ReportCode split, group and split by LineItemENG
  nested_list_of_dfs_group_split_by_lineitemeng <- lapply(list_of_dfs_group_split_by_report_code, function(df_grouped) {
    df_grouped %>% group_split(LineItemENG)
  })
  
  # Return the nested list of dataframes grouped by LineItemENG within ReportCode
  return(nested_list_of_dfs_group_split_by_lineitemeng)
}

# Apply this function to a single dataframe, 'your_dataframe'
nested_list_of_dfs_group_split <- nested_list_of_dfs_group_split(combined_df)  

#
process_df_secondary <- function(df) {
  processed_year <- list()
  column_to_check <- 'ReportYear'
  instances_over_two <- 0
  
  cat("Processing dataframe with", nrow(df), "rows\n")
  
  for (i in 1:nrow(df)) {
    current_value <- df$FVYear[i]
    
    if (!(current_value %in% processed_year)) {
      processed_year <- append(processed_year, current_value)
      found_matches <- which(df$FVYear == current_value)
      cat("Found matches for FVYear =", current_value, ":", found_matches, "\n")
      
      if (length(found_matches) > 1) {
        cat("Multiple rows for FVYear", current_value, "\n")
        
        if (length(found_matches) > 2) {
          instances_over_two <- instances_over_two + 1
        }
        
        # Extract the column data and check if it's numeric
        column_data <- df[found_matches, column_to_check]
        value_data <- df[found_matches, "Value"]
        
        # Convert column_data to numeric, handling any non-numeric values
        column_data_numeric <- suppressWarnings(as.numeric(as.character(column_data)))
        if (any(is.na(column_data_numeric))) {
          cat("Warning: Non-numeric values found in ReportYear. Treating as NA:\n", column_data, "\n")
          column_data_numeric <- ifelse(is.na(column_data_numeric), -Inf, column_data_numeric) # Handle NAs by assigning a very low value
        }
        
        # Sort column_data and value_data in decreasing order
        sorted_indices <- order(column_data_numeric, decreasing = TRUE)
        column_data_numeric <- column_data_numeric[sorted_indices]
        value_data <- value_data[sorted_indices]
        
        cat("Column data after sorting:", column_data_numeric, "\n")
        cat("Value data after sorting:", value_data, "\n")
        
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
        
        # Remove rows that are not the max_value_index
        found_matches <- found_matches[found_matches != max_value_index]
        
        if (length(found_matches) > 0) {
          # Debugging: Print the indices to be removed
          cat("Dropping rows with indices:", found_matches, "\n")
          df <- df[-found_matches, ]
        }
      }
    }
  }
  
  # Debugging: Print the number of rows after processing
  cat("Number of rows after processing:", nrow(df), "\n")
  cat("Number of instances with more than two matches:", instances_over_two, "\n")
  
  return(df)
}


final_processed_list <- lapply(nested_list_of_dfs_group_split, function(inner_list){
  lapply(inner_list, function(df) process_df_secondary(df))
})

transform_to_wide_format <- function(df) {
  # Create the new column combining LineItemENG and the last 2 digits of FVYear
  df <- df %>%
    mutate(LineItem_FVYear = paste0(LineItemENG, "_", substr(FVYear, 3, 4))) %>%
    
    # Ensure Value is numeric
    mutate(Value = as.numeric(Value))
  
  # Transform the dataframe from long to wide format using pivot_wider
  df_wide <- df %>%
    select(ReportCode, LineItem_FVYear, Value) %>%
    pivot_wider(names_from = LineItem_FVYear, values_from = Value)
  
  return(df_wide)
}

# Apply the function to each dataframe in a list of dataframes
# combined_data_list is the list of dataframes
final_wide_data_list <- lapply(combined_data_list, transform_to_wide_format)



