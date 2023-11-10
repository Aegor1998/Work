# Author: Aegor 1998  |  Created: 11/10/2023
# Purpose: To automate the generation of excel reports

# Load Raw Data Sets
# Set 1: Normal inventory
# Set 3: Damaged Computers
# Set 4: Service-Now Everett Inventory

# Creaty Empty Data Sets
# Table 1: In Stock Inventory. This will have 6 columns: Company, Serial, Type, Make, Model, Warrenty, Asset Function
inStockInventory <- data.frame(
  Company = c(),
  Serial = c(),
  Type = c(),
  Make = c(),
  Model = c(),
  Warrenty = c(), 
  assetFunction = c()
)
# Table 2: Missing Stock (Stock that Service now says we have but we did not find) This will have 6 columns: Company, Serial, Type, Make, Model, Warrenty, Asset Function
missingStock <- data.frame(
  Company = c(),
  Serial = c(),
  Type = c(),
  Make = c(),
  Model = c(),
  Warrenty = c(), 
  assetFunction = c()
)
# Table 3: Out of Warrenty. This will have 6 columns: Company, Serial, Type, Make, Model, Warrenty, Asset Function
outofWarrenty <- data.frame(
  Company = c(),
  Serial = c(),
  Type = c(),
  Make = c(),
  Model = c(),
  Warrenty = c(), 
  assetFunction = c()
)
# Table 4: Damaged Stock. This will have 6 columns: Company, Serial, Type, Make, Model, Warrenty, Asset Function
damagedStock <- data.frame(
  Company = c(),
  Serial = c(),
  Type = c(),
  Make = c(),
  Model = c(),
  Warrenty = c(), 
  assetFunction = c()
)
# Table 5: Missing and Out of Warrenty: This will have 6 columns: Company, Serial, Type, Make, Model, Warrenty, Asset Function
missingAndOutOfWarrenty <- data.frame(
  Company = c(),
  Serial = c(),
  Type = c(),
  Make = c(),
  Model = c(),
  Warrenty = c(), 
  assetFunction = c()
)

# Ingest Raw Data Sets into the Empty Data Sets
# Table 3: Out of warrenty will need to check and remove duplicates from all other Tables Except Missing and Out of Warrenty.
# That will be put in Table 5: Missing and Out of Warrenty.

# Export New Data Sets as 1 Excel File with mutiple sheets.
