library(readxl)
library(dplyr)
library(rstudioapi)
library(jsonlite)

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

#Load corresponding geo-eng lineitems

##Load JSON file as list
corresponding_lineitems_json <- fromJSON("lineitem_data/corresponding_lineitems.json")

##Convert JSON list to dataframe making sure both English and Georgian names are put in columns (and not in column names)

lookup_table <- data.frame(English = names(corresponding_lineitems_json), 
                 Georgian = unlist(corresponding_lineitems_json))

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
  
  # Check if 'LineItemENG' already exists, and only add it if it doesn't
  if (!"LineItemENG" %in% colnames(df)) {
    df <- df %>%
      left_join(lookup_table, by = c("LineItemGEO" = "Georgian")) %>%
      rename(LineItemENG = English)
  }
  
  return(df)
}



uniform_data_list <- lapply(data_list, make_identical, lookup_table = lookup_table)



# Filtering criteria

## Columns to keep

columns_to_keep_1st_revision <- c("ReportCode", "IdCode", "ReportYear", "FVYear", "CategoryMain", "FormName",
                                  "SheetName", "LineItemGEO", "LineItemENG", "Value", "GEL", "LineItem")

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

check_and_process_dfs <- function(dfs, columns_to_keep, variables_financial_non_financial, variables_financial_other, variables_profit_loss, variables_cash_flow) {
  
  # A list to store the unique LineItemENG-LineItemGEO mappings
  line_item_mappings <- list()
  
  # Iterate over each dataframe in the list
  for (i in seq_along(dfs)) {
    df <- dfs[[i]]
    
    # Step 1: Apply filtering based on the specified logic
    df <- df %>%
      select(all_of(columns_to_keep)) %>%
      filter(CategoryMain != "III ჯგუფი") %>%
      filter(FormName != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") %>%
      filter(LineITemGEO != "მარაგების გაუფასურების (ხარჯი) / აღდგენა")
      mutate(Value = if ("Value" %in% colnames(df)) as.numeric(Value) else Value) %>%
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
      filter(!(FormName == "non-financial institutions" & !LineItemENG %in% variables_financial_non_financial)) %>%
      filter(!(FormName == "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_non_financial)) %>%
      filter(!(FormName != "non-financial institutions" & SheetName == "financial position" & !LineItemENG %in% variables_financial_other)) %>%
      filter(!(SheetName == "profit-loss" & !LineItemENG %in% variables_profit_loss)) %>%
      filter(!(SheetName == "cash flow" & !LineItemENG %in% variables_cash_flow)) %>%
      arrange(ReportCode)
    
    # Step 2: Check if LineItemENG exists and store unique pairs
    if ("LineItemENG" %in% colnames(df)) {
      unique_pairs <- unique(df[, c("LineItemENG", "LineItemGEO")])
      line_item_mappings[[i]] <- unique_pairs
      message(paste("DataFrame", i, "has LineItemENG column."))
    } else {
      # If LineItemENG does not exist, create it based on LineItemGEO
      if ("LineItemGEO" %in% colnames(df)) {
        df$LineItemENG <- df$LineItemGEO  # Customize this as needed
        message(paste("DataFrame", i, "created LineItemENG based on LineItemGEO."))
      } else {
        message(paste("DataFrame", i, "does not have LineItemGEO, no action taken."))
      }
    }
    
    # Update the dataframe in the list
    dfs[[i]] <- df
  }
  
  # Step 3: Check for consistency in LineItemENG and LineItemGEO mappings across dataframes
  if (length(line_item_mappings) > 1) {
    reference_mapping <- line_item_mappings[[1]]
    
    for (i in seq(2, length(line_item_mappings))) {
      if (!identical(reference_mapping, line_item_mappings[[i]])) {
        warning(paste("Inconsistency found between DataFrame 1 and DataFrame", i))
      } else {
        message(paste("DataFrame", i, "has consistent LineItemENG and LineItemGEO values with DataFrame 1."))
      }
    }
  }
  
  return(dfs)
}

primary_processed_list <- lapply(uniform_data_list, check_and_process_dfs, columns_to_keep = columns_to_keep,
                                 variables_financial_non_financial = variables_financial_non_financial,
                                 variables_financial_other = variables_financial_other,
                                 variables_profit_loss = variables_profit_loss,
                                 variables_cash_flow)

