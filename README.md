1. Account_clean_up.py
Purpose: Cleans and restructures CRM account data for consistency and standardization.

Inputs:

Excel file with columns like Account Name, Business Type, etc.

Features:

Auto-removes unnecessary columns

Reorders columns to a standardized format

Exports a cleaned Excel file for CRM upload

2. Application,Machine,series.py
Purpose: Matches IMM models to machine series and application keywords using uploaded regex and mapping files.

Inputs:

regex.xlsx — with regex patterns for model extraction

MATCH_APPLICATION.xlsx — mapping of keywords to application types

Input Excel with Product Description and Supplier

Features:

Extracts machine model and tonnage

Matches application types

Fully upload-based; no local paths required

3. Clean_Up_Shipper_and_Consignee.py
Purpose: Clusters and cleans up brand names for both buyers and suppliers using fuzzy matching.

Inputs:

Excel file with Buyer and Supplier columns

Features:

Cleans names (removing legal suffixes, punctuation)

Groups similar entities using brand extraction and fuzzy logic

Outputs a cleaned version of the file

4. IMM_Clean_Up_Shipper_and_Consignee.py
Purpose: Extracts IMM model names, and tonnage from import records using uploaded regex and keyword files.

Inputs:

regex.xlsx — regex patterns for machine series

Input Excel with Product Description and Supplier

Features:

Detects headers automatically

Extracts and inserts Model, Tonnage columns

Provides downloadable Excel output

5. Matching_Classification_Tool.py
Purpose: Combines CRM status matching and importer classification in one tool.

Inputs:

Excel file with company names, CRM status, and yearly import volumes, and CRM Database

Features:

Classifies companies into A/B/C/D/P/N based on import trends and CRM status

Automatically matches CRM information

One-click download of final output
