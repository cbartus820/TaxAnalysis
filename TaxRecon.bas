Attribute VB_Name = "TaxRecon"
Option Explicit
' If you are reading these notes, ethier I have moved on from the company, or been hit by a bus. Regardless, I have tried to leave a good trail with information that you will need to use this script etc.
' below you will find global and public variables that have been defined for this collection of scripts, although there are also some local variables defined within the various scripts below.
' I have tried to leave notes for you to understand the purpose and function of each variable and code. Good Luck if you run into trouble, if I am alive I am happy to help fix this.

'User Execution Options and Common Data between scripts - Userform
Global Credit_Report_Toggle As Long 'On the Userform that loads when you start a new tax recon, this variable holds the true or false value from the user's toggling of the import error report check box
Public StartDate As String 'Define Variable for Start date for sage data querry
Public EndDate As String 'Define Variable for End date for sage data querry
Public ReconState As String 'Define variable for State to be reconciled for sage Data Querry. This variable is populated at the start from the User Interface
Public i As Integer 'Defined looping variable used in many many places throughout the scriptings

'Sage Queries - Objects used within data dumps from ERP
Public Conn As New ADODB.Connection ' Variable for ADODB SQL Querry into Sage to download Sage Data
Public Recset As New ADODB.RecordSet ' Variable to define a recordset to store the data querried from Sage
Public SQL As String ' Variable to hold the SQL querry in string form
Public ConString As String 'Variable to hold DNS information for the ODCB database connection

'Pivottable Headers - Strings, not strictly needed
Public V_Sum_Tx_Sales As String 'String object to Defined Column Header for Taxable Sales
Public V_Sum_NTx_Sales As String 'String object to Defined Column Header for Non-Taxable Sales
Public V_Sum_Freight_Sales As String 'String object to Defined Column Header for Freight Sales
Public V_Sum_Tax As String 'String object to Defined Column Header for Tax
Public V_Sum_Exempt_Amt As String 'String object to Defined Column Header for Exempt Amount
Public V_Sum_Gross_Amt As String 'String object to Defined Column Header for Gross Amount
Public Vertex_NonTax_Compare As Worksheet

'General Variables
Public Taxbook As Workbook  'Workbook Variable to defined the reconciliation book to Vertex Source Data
Public SourceData As Workbook ' Variable to set the raw vertex transactional report download too

'Formula Strings - User in Error Report or used in formula to calculate variances between data and facilitate error report sheet
Public MktplcLookupTax As String ' String Variable to hold a formula string to lookup the tax amount on transactions coded as mktplcfac from the Mktplcfac Sage Tab
Public MktplcLookupFreight As String 'Variable to hold formula string for comparing mktplc data
Public FreightLookupNonTax As String 'String Containing Xlookup Function to Retreveie nontaxable Freight Sales from the Gross Freight Sales Pivot Table (Vertex)
Public FreightLookupExemptTax As String 'String Containing lookup function to Retreveie exempt Freight Sales from the Gross Freight Sales Pivot Table for a particular invoice no (Vertex)
Public VertexTaxSales As String 'String Containing a Lookup function to lookup vertex gross taxable sales for a particular invoice no
Public VertexNoNTaxSales As String 'String Containing a Lookup function to lookup vertex non-taxable sales for a particular invoice no
Public Vertex_Gross As String ' String to store cell reference location of vertex gross sales for summary sheet
Public Vertex_Taxable As String ' String to store cell reference location of vertex taxable sales for summary sheet
Public Vertex_NonTaxable As String ' String to store cell reference location of vertex non-taxable sales for summary sheet
Public Vertex_Exempt As String ' String to store cell reference location of vertex exempt sales for summary sheet
Public Vertex_Freight_Amt As String 'String to Lookup Freight Amount from Vertex Data
Public Vertex_Tax_Amt As String 'String to Lookup Tax Amount from Vertex Data by Invoice No
Public Vertex_Taxable_Freight As String 'String to Lookup taxable freight Amount from Vertex Data by Invoice No
Public Vertex_NonTaxable_Freight As String 'String to Lookup non-taxable freight Amount from Vertex Data by Invoice No
Public Vertex_Exempt_Freight As String 'String to Lookup exempt freight Amount from Vertex Data for summary Sheet
Public FreightLookupTax As String 'String to Lookup tax on freight from vertex freght sales pivot table
Public Mktplc_Taxable As String 'String to Lookup Taxable Sales from Marketplace Facilitator Pivot
Public Mktplc_NonTaxable As String 'String to Lookup nontaxable Sales from Marketplace Facilitator Pivot
Public Mktplc_Freight As String 'String to Lookup freight Sales from Marketplace Facilitator Pivot
Public Mktplc_Tax As String 'String to Lookup tax from Marketplace Facilitator Pivot
Public Sage_Taxable As String 'String to Lookup Sage Taxable Sales
Public Sage_NonTaxable As String ' String to lookup Sage nontaxable sales
Public Sage_Freight As String 'String to lookup Sage Freight
Public Sage_Tax As String 'String to Lookup Sage Tax

'Data Source Dimensions
Public V_Last_Row As Long 'variable to hold the numeric value of the last row in the vertex transactional detail report
Public V_Last_Col As Long 'variable to hold the numeric value of the last column in the vertex transactional detail report
Public Vertex_Tax_First_Row As Long 'First Row of Data in Vertex Gross Tax Pivot Table
Public Vertex_Tax_Last_Row As Long 'Last Row of Data in Vertex Gross Tax Pivot Table
Public Vertex_Freight_Sales_First_Row As Long 'First Row of Data in Vertex Gross Freight Sales Pivot Table
Public Vertex_Freight_Sales_Last_Row As Long 'Last Row of Data in Vertex Gross Freight Sales Pivot Table
Public S_First_Row As Long ' First Row of Sage AR Source Data
Public S_Last_Row As Long ' Law Row of Sage AR Source Data
Public S_Last_Col As Long ' Last Column of Sage AR Source Data
Public Sage_Sales_First_Row As Long ' First Row of Sage AR Data pivoted - excluding Mktplcfac
Public Sage_Sales_Last_Row As Long 'Last Row of Sage AR Data Pivoted - excluding Mktplcfac
Public Mktplc_LR As Long ' Last Row of Sage AR Data - Excluding non mktplcfac
Public Mktplc_HR As Long ' First Row of Sage AR Data - Excluding non mktplcfac
Public Mktplc_LC As Long ' Last Column of Sage AR Data - Excluding non mktplcfac
Public DropboxDir As String



'Sheet - "Credit Report"
Public creditreport As Worksheet ' Variable to define the sheet where the credit report would be stored if a user toggled import error report on the initial user form
Public openfilename As String 'Variable to hold the directory location of the credit report inputed by the user
Public Credit_Report_Inquiry As String ' Variable is superceded. Was part of replaced import credit report process. Did not delete variable out of fear of accidentially breaking something

'Sheet - Vertex - Sheet Vertex is the home of the Vertex Data for the reconciliation
Public Vertex As Worksheet ' Define Variable for Vertex Data Tab

'Sheet - Summary - This sheet is used as a high level information summary sheet upon completion of the execution of the reconciliation script
Public Summary As Worksheet 'Define Variable for Tax Data Summary Tab - This sheet will summarize the gross data, present gross variances, and other important information that may be relevant for filing taxes in various states

'Sheet - Sage AR Data
Public Sage As Worksheet 'Define Variable For Sage Data Tab - Where data from Sage SQL querry will be stored as a datasource

'Sheet - Sage Pivot
Public Sage_Pivot As Worksheet 'Define Variable for Sage Pivot tab - Sage Pivot tab will pivot all sage sales based on trimmed invoice no that are not marketplace sales
Public S_Pivot As Range
Public Sage_Data_Cache As PivotCache 'Variable to create a pivot cache, the database element for a pivot table, for Sage transactional data

Public Vertex_Freight As PivotTable 'pivot table was repurposed for tax
Public Vertex_Data As PivotTable 'PivotTable Object for Vertex Sourced Data

'Sheet - Zipcodes
Public Import_ZIPCODE_CSV As Workbook 'Workbook object to load zipcode database from
Public ZipCode_Data As Worksheet 'Worksheet Object to assign variable to destination for filtered zipcode information
Public ZipCode_Source As Worksheet 'Worksheet Object to assign variable to Source Sheet for Zipcode Database
Public V_InvoiceNo_PF As PivotField 'PivotTable Pivotfield Object for Trimmed Invoice Number in the Vertex Gross Sales Pivottable

'Variable by Sheets

'Sheet - Vertex Pivot - Vertex Pivot Sheets contains various pivots of vertex data to fleshout data for reconciliation.
'This includes a pivot of sales which is a pivot table of vertex fitlered for State Jurisdiction, and General Sales and Use Tax imposition type.

Public Vertex_Pivot As Worksheet ' Define Variable for Vertex Pivots tab - Worksheet Will contain multiple pivot tables with different filtered views of the Vertex Data pivoted used in various different variance checks
Public Vertex_Data_Cache As PivotCache 'Variable to create a pivot cache, the database element for a pivot table, for Vertex transactional data
Public V_Pivot As Range 'Range Object for Pivot Cache

    'Vertex Gross Sales Pivot on Sheet "Vertex Pivot"
    Public Vertex_Sales_Last_Row As Long 'First Row to define range for vertex gross sales pivottable (used in formula string objects)
    Public Vertex_Sales_First_Row As Long 'Last Row to define range for vertex gross sales pivottable (used in formula string objects)
    Public V_Exempt_Amt_PF As PivotField 'PivotTable Pivotfield Object for Exempt Sales in the Vertex Gross Sales Pivottable
    Public V_Taxable_Amt_PF As PivotField 'PivotTable Pivotfield Object for Taxable Sales in the Vertex Gross Sales Pivottable
    Public V_NonTaxable_Amt_PF As PivotField 'PivotTable Pivotfield Object for Nontaxable-Sales in the Vertex Gross Sales Pivottable
    Public V_Gross_Amt_PF As PivotField 'Pivottable Pivotfield Object for Vertex Gross Sales PivotTable
    Public V_Tax_Amt_PF As PivotField 'Pivottable Pivotfield Object for tax amount for Vertex Gross Sales PivotTable
    Public V_Imposition_Type As PivotField 'Pivottable Pivotfield Object Imposition Type for Vertex Gross Sales PivotTable
    Public V_Jurisdiction_Type As PivotField 'Pivottable Pivotfield Object Jurisdiction Type Type for Vertex Gross Sales PivotTable
    Public V_FlexCode05 As PivotField 'Pivottable Pivotfield Object Flexcode 05 for Vertex Gross Sales PivotTable
    Public V_Product_Class_Code As PivotField 'Pivottable Pivotfield Object Product Class Code for Vertex Gross Sales PivotTable
    Public V_TaxExempt As PivotField
    'Vertex_Gross_Tax Pivottable on Sheet "Vertex Pivot"
    Public Vertex_Tax As PivotTable ' Variable to define vertex pivot table for tax
    Public VT_Product_Class_Code As PivotField 'Pivottable Pivotfield Object for product class code for Vertex Gross Tax PivotTable
    Public VT_FlexCode05 As PivotField 'Pivottable Pivotfield Object for Flex Code 05 for Vertex Gross Tax PivotTable
    Public VT_Imposition_Type As PivotField 'Pivottable Pivotfield Object for Imposition Type for Vertex Gross Tax PivotTable
    Public VT_Jurisdiction_Type As PivotField 'Pivottable Pivotfield Object for Jurisdiction Type for Vertex Gross Tax PivotTable
    Public VT_Tax As PivotField 'Pivottable Pivotfield Object for Tax Amount for Vertex Gross Tax PivotTable
    Public VT_Tax_InvNo As PivotField 'Pivottable Pivotfield Object for Trimmed Invoice No for Vertex Gross Tax PivotTable
    'Vertex_Gross_Freight_Sales Pivottable on Sheet "Vertex Pivot"
    Public VFS_InvoiceNo As PivotField
    Public VFS_Tax_Amount As PivotField
    Public VFS_Product_Class_Code As PivotField 'Pivottable Pivotfield Object for Product Class Code for Vertex Freight Sales PivotTable
    Public VFS_Flexcode05 As PivotField 'Pivottable Pivotfield Object for Flex Code 05 for Vertex Freight Sales PivotTable
    Public VFS_Gross_Amt As PivotField 'Pivottable Pivotfield Object for Gross Amount for Vertex Freight Sales PivotTable
    Public VFS_Exempt_Amt As PivotField 'Pivottable Pivotfield Object for Exempt Amount for Vertex Freight Sales PivotTable
    Public VFS_NonTax_Amt As PivotField 'Pivottable Pivotfield Object for Non-taxable Amount for Vertex Freight Sales PivotTable
    Public VFS_Taxable_Amt As PivotField 'Pivottable Pivotfield Object for Taxable Amount for Vertex Freight Sales PivotTable
    Public VFS_Jurisdiction_Type As PivotField 'Pivottable Pivotfield Object for Jurisdiction Type for Vertex Freight Sales PivotTable
    Public VFS_ImpositionType As PivotField 'Pivottable Pivotfield Object for Jurisdiction Type for Vertex Freight Sales PivotTable
    
    Public MKTPLC_Imposition_Type As PivotField
    Public Loopitem4 As PivotItem
    Public Loopitem6 As PivotItem
    
    Public Vertex_Nontax_Exclude_Freight As PivotTable
    Public Vertex_Nontax_Exclude_Freight_InvoiceNo As PivotField
    Public Vertex_Nontax_Exclude_Freight_Product_Class_Code As PivotField
    Public Vertex_Nontax_Exclude_Freight_Amount As PivotField
    Public Vertex_Nontax_Exclude_Freight_Flexcode05 As PivotField
    Public Vertex_Nontax_Exclude_Freight_Gross_Amt As PivotField
    Public Vertex_Nontax_Exclude_Freight_Exempt_Amt As PivotField
    Public Vertex_Nontax_Exclude_Freight_NonTax_Amt As PivotField
    Public Vertex_Nontax_Exclude_Freight_Taxable_Amt As PivotField
    Public Vertex_Nontax_Exclude_Freight_Jurisdiction_Type As PivotField
    Public Vertex_Nontax_Exclude_Freight_ImpositionType As PivotField
    Public Vertex_NonTax_Exclude_Freight_Sheet As Worksheet
    Public Vertex_Nontax_Exclude_Freight_TotalExempt As PivotField
    
    Public Vertex_Gross_Sheet As Worksheet
    Public VGP_Imposition_Type As PivotField
    Public VGP_Jurisdiction_Type As PivotField
    Public VGP_Flexcode05 As PivotField
    Public VGP_Product_Class_Code As PivotField
    Public VGP_NonTax As PivotField
    Public VGP_Exempt_Amount As PivotField
    Public Vertex_Pivot_LR As Long
    
    Public Vertex_Gross_Pivot As PivotTable
    Public VGP_InvoiceNo_PF As PivotField
    Public VGP_Taxable_Sales_PF As PivotField
    Public VGP_Nontaxable_Sales_PF As PivotField
    Public VGP_Gross_Sales_PF As PivotField
    Public VGP_Freight_PF As PivotField
    Public VGP_Tax_Amt_PF As PivotField
    Public VGP_TaxSch_PF As PivotField
    Public Vertex_Gross_LC As Long
    Public LoopItem As PivotItem
    Public LoopItem_CO As PivotItem
    
    Public Vertex_NonTaxable_Pivot As PivotTable
    
    Public VNTP_InvoiceNo_PF As PivotField
    Public VNTP_Taxable_Sales_PF As PivotField
    Public VNTP_Nontaxable_Sales_PF As PivotField
    Public VNTP_Gross_Sales_PF As PivotField
    Public VNTP_Freight_PF As PivotField
    Public VNTP_Tax_Amt_PF As PivotField
    Public VNTP_TaxSch_PF As PivotField
    Public Vertex_Nontaxable_Pivot_LC As Long
    Public VNTP_Loopitem As PivotItem
    Public VNTP_LoopField As PivotField
    Public VNTP_Imposition_Type As PivotField
    Public VNTP_Jurisdiction_Type As PivotField
    Public VNTP_Flexcode05 As PivotField
    Public VNTP_Product_Class_Code As PivotField
    Public VNTP_NonTax As PivotField
    Public VNTP_Exempt_Amount As PivotField
    Public VNTP_Gross_Nontax_PF As PivotField
    Public VNTP_Product_Class As PivotField
    Public VNTP_PostalCode As PivotField
    Public VNTP_SitusCity As PivotField
    Public Vertex_NonTax_Compare_LC As Long
    Public Vertex_NonTax_Compare_LR As Long



'The purpose of the Mktplcfac pivot is to isolate the Mktplfac sales outside of regular sage sales, so that when preparing the error report, it will list first issues with non mktplcfacilitated sales, and then issues with mktplcfacilitated sales.
'The reason they are segregated is because at mid month check in, if all the marketplace facilitator sales are in one

'Sheet - MKTPLCFAC
Public Mktplc As Worksheet ' Define Variable for MKTPLCFAC tab, - Sage Pivot tab based on trimmed invoice number that is exclusively marketplace facilitator sales
Public X As Integer ' Integer Object used in a loop through the marketplace facilitator sales (SAGE DATA)
Public MktplcPiv As PivotTable ' Pivottable Object to pivot sage data for facilitator sales (SAGE DATA)
Public M_T_InvoiceNo_PF As PivotField 'Pivotfield object for Trimmed Invoice No for Mktplcfac Pivot (SAGE DATA)
Public M_Taxable_Sales_PF As PivotField 'Pivotfield object for Taxable Sales for Mktplcfac Pivot (SAGE DATA)
Public M_Non_Taxable_Sales_PF As PivotField 'Pivotfield object for Nontaxable for Mktplcfac Pivot (SAGE DATA)
Public M_Tax_Amt_PF As PivotField 'Pivotfield object for Tax Amt for Mktplcfac Pivot (SAGE DATA)
Public M_Freight_PF As PivotField 'Pivotfield object for Freight Amt for Mktplcfac Pivot (SAGE DATA)
Public M_TaxSch_PF As PivotField 'Pivotfield object for Tax Schedule for Mktplcfac Pivot (SAGE DATA)
Public M_Gross_Sales_PF As PivotField
Public Loopitem5 As PivotItem



'Colorado HomeRule Specific
Public VTC_Product_Class_Code As PivotField
Public VTC_FlexCode05 As PivotField
Public VTC_Imposition_Type As PivotField
Public VTC_Jurisdiction_Type As PivotField
Public VTC_Tax As PivotField
Public VTC_Tax_InvNo As PivotField

'Sheet -Sage Pivot
Public Sage_Data As PivotTable 'Define Variable for Sage Pivot Table
Public S_T_InvoiceNo_PF As PivotField
Public S_Taxable_Sales_PF As PivotField
Public S_Non_Taxable_Sales_PF As PivotField
Public S_Tax_Amt_PF As PivotField
Public S_Freight_PF As PivotField
Public S_TaxSch_PF As PivotField
Public S_Gross_Sales_PF As PivotField
Public Sage_Pivot_LR As Long
Public Sage_Pivot_LC As Long




Public S_Sum_Tx_Sales As String
Public S_Sum_NTx_Sales As String
Public S_Sum_Freight_Sales As String
Public S_Sum_Tax As String


'Sheet - Error Report
Public Counter As Integer ' Integer Object Used in Creation of Error Report for Variances



'Defunct/Future Use
Public Total_Return_Tax As Long
Public VF_InvoiceNo As PivotField
Public VF_Product_Class_Code As PivotField
Public VF_Tax_Amount As PivotField
Public VF_Flexcode05 As PivotField
Public VF_Imposition_Type As PivotField
Public Reconfile As Workbook 'Variable is superceded. Was part of replaced import credit report process. Did not delete variable out of fear of accidentially breaking something
Public Vertex_Freight_Sales As PivotTable 'Pivottable meant to contain exclusively freight sales. There is a bit of legacy coding involved with this. The method of determining freight was changed to a more reliable method. This variable is still used and needed in the new method
Public Vertex_Freight_Tax_First_Row As Long
Public Vertex_Freight_Tax_Last_Row As Long


Public FileName As String
Public SaveDir As String
Public Error_Report As Worksheet
Public Vertex_Inv_Tax As String
Public VT_TAXJUR As PivotField
Public Colorado_HR As Workbook
Public Homerule As Worksheet
Public CO_HomeRule_Conn As ADODB.Connection
Public Vertex_Tax_COHR As PivotTable
Public VTC_TAXJUR As PivotField
Public VTC_TAXCity As PivotField
Public b As Integer
Public p As Integer
Public Return_Data_Available As String
Global Raw_Vertex As Worksheet
Public Config As Worksheet
Global Config_LastRow As Long
Global Config_Date_LastRow As Long
Global State_List As Range
Global State_Cell As Range
Global Vertex_Extract As Workbook
Global precent As Single
Global Total_Discount_Amt As Long
Global Total_Tax_Amt As Long

Public Exemptions As Worksheet
Public Compare_Exclude_Freight As Worksheet

Public Sage_Gross As Worksheet
Public Sage_Gross_Pivot As PivotTable
Public SGP_InvoiceNo_PF As PivotField
Public SGP_Taxable_Sales_PF As PivotField
Public SGP_Nontaxable_Sales_PF As PivotField
Public SGP_Freight_PF As PivotField
Public SGP_Tax_Amt_PF As PivotField
Public SGP_TaxSch_PF As PivotField
Public SGP_Sage_Gross_Sales As PivotField
Public Sage_Gross_LC As Long

'Vertex Data Sheet
Public SitusMain As Range
Public DestinationMain As Range
Public DestinationCountry As Range
Public PostingDate As Range

'Vertex Fright Gross Sales Pivot
Public Lookup_Vertex_Freight_Exempt As Range
Public Lookup_Vertex_Freight_Taxable As Range
Public Lookup_Vertex_Freight_Sales_InvoiceNo As Range
Public Lookup_Vertex_Freight_Gross As Range

' Vertex Gross Sales Pivot
Public Lookup_Vertex_Gross_Gross As Range
Public Lookup_Vertex_Gross_Taxable As Range
Public Lookup_Vertex_Gross_InvoiceNo As Range
Public Lookup_Vertex_Gross_Exempt As Range

'Vertex Gross Sales Compare Pivot
Public Lookup_Vertex_GrossComp_Gross As Range
Public Lookup_Vertex_GrossComp_Taxable As Range
Public Lookup_Vertex_GrossComp_Exempt As Range
Public Lookup_Vertex_GrossComp_InvoiceNo As Range

'Vertex Nontaxable Gross Pivot
Public Lookup_Vertex_Exempt_Gross As Range
Public Lookup_Vertex_Exempt_Taxable As Range
Public Lookup_Vertex_Exempt_Exempt As Range
Public Lookup_Vertex_Exempt_InvoiceNo As Range


'Vertex Nontaxable Excluding Freight Pivot
Public Lookup_Vertex_NontaxNoFreight_Gross As Range
Public Lookup_Vertex_NontaxNoFreight_Taxable As Range
Public Lookup_Vertex_NontaxNoFreight_Exempt As Range
Public Lookup_Vertex_NontaxNoFreight_InvoiceNo As Range

' Sage Gross Sales Compare Pivot
Public Lookup_Sage_Gross As Range
Public Lookup_Sage_Nontaxable As Range
Public Lookup_Sage_Taxable As Range
Public Lookup_Sage_Freight As Range
Public Lookup_SageTax As Range
Public Lookup_Sage_InvoiceNo As Range

'Vertex Tax Pivot All taxes
Public Lookup_Vertex_Tax_Tax As Range
Public Lookup_Vertex_Tax_InvoiceNo As Range

'Sage Non-Marketplace Pivot
Public Lookup_Sage_NonMarketplace_Gross As Range
Public Lookup_Sage_NonMarketplace_Nontaxable As Range
Public Lookup_Sage_NonMarketplace_Taxable As Range
Public Lookup_Sage_NonMarketplace_Freight As Range
Public Lookup_Sage_NonMarketplace_Tax As Range
Public Lookup_Sage_NonMarketplace_InvoiceNo As Range


'Sage Marketplace Pivot
Public Lookup_Sage_Marketplace_Gross As Range
Public Lookup_Sage_Marketplace_Nontaxable As Range
Public Lookup_Sage_Marketplace_Taxable As Range
Public Lookup_Sage_Marketplace_Freight As Range
Public Lookup_Sage_Marketplace_Tax As Range
Public Lookup_Sage_Marketplace_InvoiceNo As Range


Public Sage_Gross_Amount As Double
Public Sage_Taxable_Amount As Double
Public Sage_NonTaxableAmount As Double
Public Sage_FreightAmt As Double
Public Vertex_Gross_Amount As Double
Public Vertex_Taxable_Amount As Double
Public Vertex_NonTaxable_Amount As Double
Public Vertex_Exempt_Amount As Double
Public Vertex_Tax_Amount As Double
Public LookupInvoiceNo As String

'Set Up Error Report Variables
Public HeaderRow As Range
Public HeaderCell As Range
Public MarketplaceStart As Long
Public SageCompareStart As Long
Public VertexCompareStart As Long
Public Error_Report_Last_Row As Long
Public MarketplaceEnd As Long
Public SageEnd As Long
Public VertexEnd As Long
Public MatchString As String
Public SageAR_LastRow As Long
Public Sage_InvoiceDate As Range
Public Sage_CustomerNo As Range
Public Sage_ShipTo1 As Range
Public Sage_ShipTo2 As Range
Public Sage_ShipTo3 As Range
Public Sage_ShipToCity As Range
Public Sage_ShipToState As Range
Public Sage_ShipToZipCode As Range
Public Sage_OrderManager As Range

'Metrics
Public StartTime As Double
Public TimeToComplete As Worksheet

'Setup Error Report

Public Trimmed_Invoice_Col As Long
Public Trimmed_Invoice_Col_Letter As String
Public Lookuplocation As String
Public LinkCell As Range
Public LinkAddress As String
Public VertexCutRange As Range
Public VertexPasteRange As Range

'Exemption Reasons

Sub RecordExecutionTime(TimeToComplete As Worksheet, row As Long, label As String, StartTime As Double, elapsedTime As Double)
    TimeToComplete.Cells(row, 1).Value = label
    TimeToComplete.Cells(row, 2).Value = elapsedTime
    TimeToComplete.Cells(row, 3).Value = elapsedTime - StartTime
End Sub

Sub Step_02_Main_Call_Recon_Subroutines()

    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Set Taxbook = Vertex_Extract


    StartTime = Timer
    Set TimeToComplete = Taxbook.Sheets.Add
    TimeToComplete.Name = "Script Metrics"
    TimeToComplete.Visible = False

'Create Blank Worksheets for tabs; Summary; Sage;SagePivot, Vertex Pivot, MKTPLCFAC, Error Report, Zipcodes, and credit report
'originally intended to create all sheets within the reconciliation file, as additional revisions and improvements and added features were implimented to the code, some subroutines still create their own worksheets

    Call Step_03_Main_CreateWorksheets
    RecordExecutionTime TimeToComplete, 2, "Fn_Create_Worksheets", StartTime, Timer
    StartTime = Timer

'Imports a Zipcode Database CSV file, originally intended to help with the populating of tax returns, the use of the zipcode database is unfortunately depreciated but the code has been retained in case it is revisited
    Call Step_04_Main_Import_Zipcode_Database
    RecordExecutionTime TimeToComplete, 3, "Import_Zipcode_Database", StartTime, Timer
    StartTime = Timer

'Determines SQL Statements for Data import and imports the data based on whether the user enabled the use quanda data checkbox on the userform and import company data (Step 1 Accumulated Vertex Data already)
    Call Step_05_Main_Import_Sage_Data
    RecordExecutionTime TimeToComplete, 4, "Import_Sage_Data", StartTime, Timer
    StartTime = Timer

'Extract from the Vertex Configuration Report a list of customers an non-expired exemptions for the reconciliation perameters
    Call Step_06_Main_ImportAndFilter_Customer_Exemptions
    RecordExecutionTime TimeToComplete, 5, "Import_Exemptions", StartTime, Timer
    StartTime = Timer

'Create a Primary Key "Trimmed Invoice No" between both datasets for comparison
    Call Step_07_Main_Convert_Invoice_Numbers
    RecordExecutionTime TimeToComplete, 6, "Convert_Invoice_Numbers", StartTime, Timer
    StartTime = Timer

'Define Ranges and Create Pivot Caches for Sage and Vertex Datasets
    Call Step_08_Main_Prepare_Data_For_Pivots
    RecordExecutionTime TimeToComplete, 7, "Create_Pivot_Range_References", StartTime, Timer
    StartTime = Timer
    
'Create Pivots of Various Views of Data

'Create a Pivot of Nontaxable Vertex Sales that exclude both Marketplace and Regular Coded Sales Freight Charges to create a comparable data set between Sage which includes frieght as its own column, and Vertex, which includes freight in Gross, Taxable, Exempt, and Nontaxable Sales.
    Call Step_09_Main_Vertex_Create_NonTaxable_Pivot_Exclude_Freight
    RecordExecutionTime TimeToComplete, 8, "Vertex_Create_NonTaxable_Pivot_Exclude_Freight", StartTime, Timer
    StartTime = Timer

'Create a pivot of vertex sales that includes only freight
    Call Step_10_Main_Create_Vertex_Freight_Sales_Pivot
    RecordExecutionTime TimeToComplete, 9, "Create_Vertex_Freight_Sales_Pivot", StartTime, Timer
    StartTime = Timer

'Create Pivot Tables for Vertex Gross Sales, and Nontaxable Sales
    Call Step_11_Main_Create_Vertex_Gross_Pivots
    RecordExecutionTime TimeToComplete, 10, "Create_Vertex_Gross_Pivot", StartTime, Timer
    StartTime = Timer
    
'Create Pivot Tables for Vertrex Tax
    Call Step_12_Create_Vertex_Tax_Pivot
    RecordExecutionTime TimeToComplete, 11, "Create_Vertex_Tax_Pivot", StartTime, Timer
    StartTime = Timer
'Create Pivot of Sage Data for Comparison to Vertex Data
    Call Step_13_FN_Create_Sage_Pivot
    RecordExecutionTime TimeToComplete, 12, "Create_Sage_Pivot", StartTime, Timer
    StartTime = Timer
    
'Create a Pivot of Sage Data for Marketplace Sales
    Call Step_14_Create_Mktplc_Sage_Pivot
    RecordExecutionTime TimeToComplete, 13, "Create_Mktplc_Sage_Pivot", StartTime, Timer
    StartTime = Timer

    Call Step_15_Main_Compare_Data
    RecordExecutionTime TimeToComplete, 14, "Compare_Data", StartTime, Timer
    StartTime = Timer

    Call Step_16_Main_Setup_Error_Report
    RecordExecutionTime TimeToComplete, 15, "Setup_Error_Report", StartTime, Timer
    StartTime = Timer

    Call Step_17_Main_Setup_Summary_Sheet
    RecordExecutionTime TimeToComplete, 16, "Setup_Summary_Sheet", StartTime, Timer
    StartTime = Timer

    Call W_FN_Input_Return_Data
    RecordExecutionTime TimeToComplete, 17, "Input_Return_Data", StartTime, Timer
    StartTime = Timer

    Call X_FN_Formatting
    RecordExecutionTime TimeToComplete, 18, "Formatting", StartTime, Timer
    StartTime = Timer

    Call Y_FN_Create_Exemption_Pivot
    RecordExecutionTime TimeToComplete, 19, "Create_Exemption_Pivot", StartTime, Timer
    StartTime = Timer

    Call Z_FN_Lookup_Exemption_Reason
    RecordExecutionTime TimeToComplete, 20, "Lookup_Exemption_Reason", StartTime, Timer
    StartTime = Timer

    Call ZA_FN_Format_Summary
    RecordExecutionTime TimeToComplete, 21, "Format_Summary", StartTime, Timer
    StartTime = Timer

    Call ZB_FN_Format_Headers
    RecordExecutionTime TimeToComplete, 22, "Format_Headers", StartTime, Timer
    StartTime = Timer

    Call ZD_Create_Index
    RecordExecutionTime TimeToComplete, 23, "Create_Index", StartTime, Timer
    StartTime = Timer

If Round(Summary.Range("E4").Value, 2) <> Round(Summary.Range("G4").Value, 2) _
Or Round(Summary.Range("E5").Value, 2) <> Round(Summary.Range("G5").Value, 2) _
Or Round(Summary.Range("E6").Value, 2) <> Round(Summary.Range("G6").Value, 2) _
Or Round(Summary.Range("E7").Value, 2) <> Round(Summary.Range("G7").Value, 2) Then
MsgBox ("Completed Task but there are variances in the datasets.")


Call FN_Filter_for_Errors
    RecordExecutionTime TimeToComplete, 25, "Filter_for_Errors", StartTime, Timer
    StartTime = Timer

Else
    If User_Interface.Close_Query.Value = True Then
        Else
        MsgBox ("Completed Task with no data set variances. Closing Reconciliation.")
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    End If
End If


    Call ZC_FN_Save_Recon
    RecordExecutionTime TimeToComplete, 26, "FN_Save_Recon", StartTime, Timer
    StartTime = Timer

    ' Finalize your main procedure
    Application.ScreenUpdating = True
End Sub
Sub Step_03_Main_CreateWorksheets()

Set Vertex = ActiveSheet
    On Error Resume Next
    Vertex.Name = "Vertex"
    Vertex.Activate
    Vertex.Range("A2").Select
    ActiveWindow.FreezePanes = True
    

'Create a Worksheet called Summary in order to store summarized information from the two data sets to display variances at a high-level
Set Summary = Taxbook.Sheets.Add(Before:=Vertex)
Summary.Name = "Summary"

'Create a Worksheet called Sage AR Data to hold and store Sage data for reconciliation
Set Sage = Taxbook.Sheets.Add(Before:=Vertex)
    Sage.Name = "Sage AR Data"
    
    Sage.Activate
    Sage.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
'Create a Sheet to store a Pivotable of the Sage AR Data to Summarize Invoice Data by Document Number (Adjustments will typically begin with a Letter and then the 6 digit invoice number excluding the 0, and a pivot table allows for a comparison of those documents as a single reference number)
Set Sage_Pivot = Taxbook.Sheets.Add
Sage_Pivot.Name = "Sage Pivot"
    Sage_Pivot.Activate
    Sage_Pivot.Range("A7").Select
    ActiveWindow.FreezePanes = True

' Create a Worksheet named Vertex Pivot to store various pivot tables with Vertex data in various views
Set Vertex_Pivot = Taxbook.Sheets.Add
Vertex_Pivot.Name = "Vertex Pivot"


'Create a Worksheet named MKTPLCFAC to compare Marketplace Sales between Sage and Vertex
Set Mktplc = Taxbook.Sheets.Add
ActiveSheet.Name = "MKTPLCFAC"
Mktplc.Activate
Mktplc.Range("A7").Select
ActiveWindow.FreezePanes = True

'Create a Worksheet named "Error Report" to contain various variances between Vertex and Sage Data on a per document number level
Set Error_Report = Taxbook.Sheets.Add
Error_Report.Name = "Error Report"

'Create a Worksheet named "Zipcodes" to contain a zipcode database used in various state specific reconciliations
Set ZipCode_Data = Taxbook.Sheets.Add
ZipCode_Data.Name = "Zipcodes"

' Check to see if the Script User opted to import the credit report, if so then open the report file stored in the dropbox directory, and import it into the excel workbook.
If User_Interface.CheckBox1 = True Then
'MsgBox ("Select the Credits and Adjustments Export file")
    openfilename = DropboxDir & "Multi-State Sales Tax\Vertex\Vertex - Credit Report\" & Format(EndDate, "YYYY-MM") & " FOT_Credits_And_Adjustments.csv"
    
    With Taxbook
        Sheets.Add After:=ActiveWorkbook.Sheets("Vertex")
        Set creditreport = ActiveSheet
        creditreport.Name = "Credit Report"
    End With
    
    With creditreport.QueryTables.Add(Connection:="Text;" & openfilename, Destination:=creditreport.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
End If

End Sub

Sub Y_FN_Create_Exemption_Pivot()
Dim Exemption_Pivot As Worksheet
Dim Exemption_Reason_Pivot_Cache As PivotCache
Dim Exemption_Range As Range
Dim Exemption_Pivots As PivotTable
Dim County As PivotField
Dim city As PivotField
Dim Exemption_Reason As PivotField
Dim Exemp_VGA As PivotField
Dim Exemp_VNTA As PivotField
Dim Exemp_VEA As PivotField
Dim Exemp_VGNA As PivotField
Dim Exemption_Pivot_Sheet As Worksheet
Set Exemption_Pivot = Taxbook.Sheets.Add(After:=Taxbook.Sheets("Summary"))
Exemption_Pivot.Name = "Exemption Pivot Data"
Vertex_NonTax_Compare.Range("A" & Vertex_NonTax_Compare.Cells.Find("Trimmed Invoice No", After:=Vertex_NonTax_Compare.Range("A1")).row & ":T" & WorksheetFunction.CountA(Vertex_NonTax_Compare.Range("A:A")) + 1).Copy
Exemption_Pivot.Range("A1").PasteSpecial (xlPasteAll)
Set Exemption_Range = Exemption_Pivot.UsedRange
Set Exemption_Pivot_Sheet = Taxbook.Sheets.Add(After:=Taxbook.Sheets("Summary"))
Exemption_Pivot_Sheet.Name = "Return Report"
Exemption_Pivot_Sheet.UsedRange.Style = "Comma"
Set Exemption_Reason_Pivot_Cache = Taxbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=Exemption_Range)
Set Exemption_Pivots = Exemption_Reason_Pivot_Cache.CreatePivotTable(TableDestination:=Exemption_Pivot_Sheet.Range("A1"), _
TableName:="Exemption_Pivots")
Exemption_Pivots.RowAxisLayout xlTabularRow
'Set County = Exemption_Pivots.PivotFields("County")
'Set City = Exemption_Pivots.PivotFields("City")
Set Exemp_VGA = Exemption_Pivots.PivotFields("Vertex_Gross_Amt")
Set Exemp_VNTA = Exemption_Pivots.PivotFields("Vertex_Gross Nontaxable")
Set Exemption_Reason = Exemption_Pivots.PivotFields("Exemption Reason")
Exemption_Reason.Orientation = xlColumnField
'County.Orientation = xlRowField
'County.Subtotals(1) = True
'County.Subtotals(1) = False
'City.Orientation = xlRowField
'City.Subtotals(1) = True
'City.Subtotals(1) = False
Exemption_Reason.Orientation = xlRowField
Exemption_Pivots.AddDataField Exemp_VGA, "Vertex Nontax or Exempt", xlSum
Exemption_Pivots.RepeatAllLabels xlRepeatLabels
Exemption_Pivots.NullString = "0"
Exemption_Pivot_Sheet.Columns.AutoFit
Exemption_Pivot_Sheet.UsedRange.Style = "Comma"
Summary.Activate
End Sub

Sub ZC_FN_Save_Recon()
Dim fso As New FileSystemObject
Dim Archive As New FileSystemObject
Dim file As Variant
Dim FileDir As Object
Dim LoopFile As Object
Dim ParseString As String

Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
fso.CreateFolder (DropboxDir & "Multi-State Sales Tax\Tax Jursidictions\" & ReconState & "\" & Format(EndDate, "YYYY"))
SaveDir = DropboxDir & "Multi-State Sales Tax\Tax Jursidictions\" & ReconState & "\" & Format(EndDate, "YYYY") & "\" & Format(EndDate, "yyyy-mm")
On Error Resume Next

fso.CreateFolder (SaveDir)
FileName = Format(EndDate, "YYYY-MM-DD") & " " & ReconState & " Sales Tax Reconciliation " & Format(Now, "yyyy-mm-dd h mm ss AM/PM")
Taxbook.SaveAs FileName:=SaveDir & "\" & FileName, FileFormat:=xlWorkbookDefault

Set Archive = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Archive.CreateFolder (SaveDir & "\Archive")



Dim OldReconFile As String
Dim ArchiveFile As String
Dim fs, f
Dim FileDateCreated
Dim CurrentRecon As Workbook
Dim NewReconDir As String


Set CurrentRecon = Taxbook
NewReconDir = CurrentRecon.FullName

OldReconFile = Dir("C:\Users\christopher.bartus\Fotronic Dropbox\Christopher Bartus\Multi-State Sales Tax\Tax Jursidictions\" & ReconState & "\" & Format(StartDate, "YYYY") & "\" & Format(StartDate, "YYYY-MM") & "\" & Format(EndDate, "YYYY-MM-DD") & " " & ReconState & " Sales Tax Reconciliation*")
ArchiveFile = "C:\Users\christopher.bartus\Fotronic Dropbox\Christopher Bartus\Multi-State Sales Tax\Tax Jursidictions\" & ReconState & "\" & Format(StartDate, "YYYY") & "\" & Format(StartDate, "YYYY-MM") & "\" & OldReconFile

If NewReconDir = ArchiveFile Then

Else

Set fs = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set f = fs.GetFile(ArchiveFile)
FileDateCreated = f.DateCreated
Name ArchiveFile As "C:\Users\christopher.bartus\Fotronic Dropbox\Christopher Bartus\Multi-State Sales Tax\Tax Jursidictions\" & ReconState & "\" & Format(StartDate, "YYYY") & "\" & Format(StartDate, "YYYY-MM") & "\Archive\" & OldReconFile


End If


End Sub
   

Public Sub Step_05_Main_Import_Sage_Data()

    ' Warn the user that querying the backend of sage will take some time
    Step_05_Secondary_InformUser

    ' Build the connection string and SQL query
    Step_05_Secondary_BuildQuery

    ' Execute the query and load data into the worksheet
    Step_05_Secondary_LoadData

End Sub

Private Sub Step_05_Secondary_InformUser()
    Dim Infobox1 As Object
    Dim AckTime As Integer
    Set Infobox1 = CreateObject("Wscript.Shell")
    AckTime = 3
    Select Case Infobox1.PopUp("Excel will become temporarily unresponsive while it retrieves data from Sage. Do not close Excel, please minimize it and work on something else until the script finishes. Click OK (this window closes automatically after 10 seconds.", AckTime, "User Notice", 0)
        Case 1, -1
    End Select
End Sub

Private Sub Step_05_Secondary_BuildQuery()
    If User_Interface.Quanda_Option = False Then
        Step_05_Secondary_Subfunction_BuildSageQuery
    ElseIf User_Interface.Quanda_Option = True Then
        Step_05_Secondary_Subfunction_BuildQuandaQuery
    End If
End Sub

Private Sub Step_05_Secondary_Subfunction_BuildSageQuery()
'Depending on whether user indicated using Quanda or Sage as Data Source format SQL querry to company data

    ConString = "DSN=SOTAMAS90; UID=cba; PWD=Huya7; Directory=\\fot00erp\Sage100_2020\MAS90; Prefix=\\fot00erp\Sage100_2020\MAS90\SY\, \\fot00erp\Sage100_2020\MAS90\==\; ViewDLL=\\fot00erp\Sage100_2020\MAS90\HOME; Company=fot; LogFile=\PVXODBC.LOG; CacheSize=4; DirtyReads=1; BurstMode=1; StripTrailingSpaces=1; SERVER=NotTheServer"

If InStr(ReconState, "ALL") > 0 Then
    SQL = "SELECT * FROM AR_InvoiceHistoryHeader WHERE (AR_InvoiceHistoryHeader.InvoiceDate >= {d '" & StartDate & "'}) AND (AR_InvoiceHistoryHeader.InvoiceDate <= {d '" & EndDate & "'}) AND (AR_InvoiceHistoryHeader.ShipToCountryCode ='US' OR AR_InvoiceHistoryHeader.ShipToCountryCode ='USA')"
Else

If ReconState = "FL" Or ReconState = "Fl" Then
SQL = "SELECT * FROM AR_InvoiceHistoryHeader WHERE (AR_InvoiceHistoryHeader.InvoiceDate >= {d '" & StartDate & "'}) AND (AR_InvoiceHistoryHeader.InvoiceDate <= {d '" & EndDate & "'}) AND (AR_InvoiceHistoryHeader.ShipToCountryCode ='US' OR AR_InvoiceHistoryHeader.ShipToCountryCode ='USA') AND AR_InvoiceHistoryHeader.ShipToState ='" & ReconState & "' AND AR_InvoiceHistoryHeader.TaxSchedule<>'MKTPLCFAC'"

Else
SQL = "SELECT * FROM AR_InvoiceHistoryHeader WHERE (AR_InvoiceHistoryHeader.InvoiceDate >= {d '" & StartDate & "'}) AND (AR_InvoiceHistoryHeader.InvoiceDate <= {d '" & EndDate & "'}) AND (AR_InvoiceHistoryHeader.ShipToCountryCode ='US' OR AR_InvoiceHistoryHeader.ShipToCountryCode ='USA') AND AR_InvoiceHistoryHeader.ShipToState ='" & ReconState & "'"
End If
End If
    ' ...
End Sub

Private Sub Step_05_Secondary_Subfunction_BuildQuandaQuery()
    ' Your code to build the Quanda query goes here
   
ConString = "DSN=Quanda;Description=Daily Replication of Sage data in SQL;Trusted_Connection=Yes;APP=Microsoft Office;DATABASE=Quanda;ApplicationIntent=READONLY;"

If InStr(ReconState, "ALL") > 0 Then
    SQL = "SELECT * FROM AR_InvoiceHistoryHeader WHERE (AR_InvoiceHistoryHeader.InvoiceDate BETWEEN '" & StartDate & "' AND '" & EndDate & "') AND (AR_InvoiceHistoryHeader.ShipToCountryCode ='US' OR AR_InvoiceHistoryHeader.ShipToCountryCode ='USA')"
Else

If ReconState = "FL" Or ReconState = "Fl" Then
SQL = "SELECT * FROM AR_InvoiceHistoryHeader WHERE (AR_InvoiceHistoryHeader.InvoiceDate BETWEEN '" & StartDate & "' AND '" & EndDate & "') AND (AR_InvoiceHistoryHeader.ShipToCountryCode ='US' OR AR_InvoiceHistoryHeader.ShipToCountryCode ='USA') AND AR_InvoiceHistoryHeader.ShipToState ='" & ReconState & "' AND AR_InvoiceHistoryHeader.TaxSchedule<>'MKTPLCFAC'"

Else
SQL = "SELECT * FROM AR_InvoiceHistoryHeader WHERE (AR_InvoiceHistoryHeader.InvoiceDate BETWEEN '" & StartDate & "' AND '" & EndDate & "') AND (AR_InvoiceHistoryHeader.ShipToCountryCode ='US' OR AR_InvoiceHistoryHeader.ShipToCountryCode ='USA') AND AR_InvoiceHistoryHeader.ShipToState ='" & ReconState & "'"
End If
End If

    ' ...
End Sub

Private Sub Step_05_Secondary_LoadData()
    Dim Conn As New ADODB.Connection
    Dim Recset As ADODB.RecordSet
    Dim i As Long

    Conn.Open ConString
    Set Recset = New ADODB.RecordSet
    Recset.Open SQL, Conn
    Sage.Cells(2, 1).CopyFromRecordset Recset

    With Recset
        For i = 1 To .Fields.Count
            Sage.Cells(1, i).Value = Recset.Fields(i - 1).Name
        Next i
    End With
    Recset.Close
    Conn.Close
    Sage.Columns.AutoFit

    FormatColumns
End Sub

Private Sub FormatColumns()
    Dim i As Long
    For i = 1 To WorksheetFunction.CountA(Sage.Range("1:1"))
        If InStr(Sage.Cells(1, i).Value, "date") Or InStr(Sage.Cells(1, i).Value, "Date") Then
            Sage.Cells(1, i).EntireColumn.NumberFormat = "mm/dd/yyyy"
        End If
    Next i
End Sub

Sub Step_07_Main_Convert_Invoice_Numbers()

    Dim i As Long
    Dim JurTypeDict As Object
    Dim Loop_JurType As Range

    ' Add "Trimmed Invoice No" column to Vertex and populate it
    With Vertex
        V_Last_Col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        V_Last_Row = .Cells(.Rows.Count, WorksheetFunction.Match("Transaction ID", .Range("1:1"), 0)).End(xlUp).row
        .Cells(1, V_Last_Col + 1).Value = "Trimmed Invoice No"
        
        ' Populate "Trimmed Invoice No" by taking right 6 digits of "Document Number"
        For i = 2 To V_Last_Row
            .Cells(i, V_Last_Col + 1).Value = Right(Left(.Cells(i, .Cells.Find("Document Number").Column), 7), 6)
        Next i
        
        .Columns.AutoFit
    End With

    ' Add "Trimmed Invoice No" column to Sage and populate it
    With Sage
        S_Last_Row = .Cells(.Rows.Count, WorksheetFunction.Match("InvoiceNo", .Range("1:1"), 0)).End(xlUp).row
        S_Last_Col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Cells(1, S_Last_Col + 1).Value = "Trimmed Invoice No"
        
        ' Populate "Trimmed Invoice No" by taking right 6 digits of "InvoiceNo"
        For i = 2 To S_Last_Row
            .Cells(i, S_Last_Col + 1).Value = Right(.Cells(i, 1).Value, 6)
        Next i
    End With

    ' Calculate Total Tax Amount for each invoice if required
    If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then
        Vertex.Cells(1, V_Last_Col + 2).Value = "Total Tax Amount"
        
        ' Create a dictionary for unique Jurisdiction Types
        Set JurTypeDict = CreateObject("Scripting.Dictionary")
        
        For Each Loop_JurType In Vertex.Range(Vertex.Cells(1, Vertex.Cells.Find("Jurisdiction Type").Column), Vertex.Cells(V_Last_Row, Vertex.Cells.Find("Jurisdiction Type").Column))
            If Not JurTypeDict.Exists(Loop_JurType.Value) Then
                JurTypeDict.Add Loop_JurType.Value, Loop_JurType.Value
            End If
        Next Loop_JurType
        
        ' Calculate Total Tax Amount for each invoice based on Jurisdiction Type
    For i = 2 To V_Last_Row
        Dim taxAmountSum As Double
        Dim jurisdictionTypeCount As Long
        
        ' Calculate the sum of Tax Amounts
        taxAmountSum = WorksheetFunction.SumIfs( _
            Vertex.Range(Vertex.Cells(1, Vertex.Cells.Find("Tax Amount").Column), Vertex.Cells(V_Last_Row, Vertex.Cells.Find("Tax Amount").Column)), _
            Vertex.Range(Vertex.Cells(1, Vertex.Cells.Find("Trimmed Invoice No").Column), Vertex.Cells(V_Last_Row, Vertex.Cells.Find("Trimmed Invoice No").Column)), _
            Vertex.Cells(i, V_Last_Col + 1).Value)
        
        ' Calculate the count of Jurisdiction Types
        jurisdictionTypeCount = WorksheetFunction.CountIfs( _
            Vertex.Range(Vertex.Cells(1, Vertex.Cells.Find("Jurisdiction Type").Column), Vertex.Cells(V_Last_Row, Vertex.Cells.Find("Jurisdiction Type").Column)), _
            Vertex.Cells(i, Vertex.Cells.Find("Jurisdiction Type").Column), _
            Vertex.Range(Vertex.Cells(1, Vertex.Cells.Find("Trimmed Invoice No").Column), Vertex.Cells(V_Last_Row, Vertex.Cells.Find("Trimmed Invoice No").Column)), _
            Vertex.Cells(i, V_Last_Col + 1))
        
        ' Calculate the Total Tax Amount
        Vertex.Cells(i, V_Last_Col + 2).Value = Round(taxAmountSum, 2)
    Next i
    
End If

End Sub

Sub Step_08_Main_Prepare_Data_For_Pivots()
    ' This subroutine prepares data for creating pivot tables by setting up PivotCaches for Sage and Vertex data.
    
    ' Set the ranges for Vertex and Sage data
    Set V_Pivot = Range(Vertex.Cells(1, 1), Vertex.Cells(V_Last_Row, V_Last_Col + 1))
    Set S_Pivot = Range(Sage.Cells(1, 1), Sage.Cells(S_Last_Row, S_Last_Col + 1))
    
    ' Create a PivotCache for Sage data
    Set Sage_Data_Cache = Taxbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=S_Pivot)
    
    ' Create a PivotCache for Vertex data based on whether the Sales Tax is included in Gross or not
    If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then
        ' Update the V_Pivot range to include the extra column for Total Tax Amount
        Set V_Pivot = Range(Vertex.Cells(1, 1), Vertex.Cells(V_Last_Row, V_Last_Col + 2))
    End If

    ' Create a PivotCache for Vertex data
    Set Vertex_Data_Cache = Taxbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=V_Pivot)
End Sub


Sub Step_13_FN_Create_Sage_Pivot()

Dim DestRange As Range

' Determine whether tax is to be included in gross sales
If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then


Application.Calculation = xlCalculationManual

' Set A range for the pivot table
Set DestRange = Sage_Pivot.Range("A5")

'Create the Pivot Table Data Cache
Set Sage_Data_Cache = Taxbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=Range(Sage.Cells(1, 1), Sage.Cells(S_Last_Row, S_Last_Col + 1)))
    
'Create the pivot table from the Pivot Table Data Cache
Set Sage_Data = Sage_Data_Cache.CreatePivotTable( _
    TableDestination:=Sage_Pivot.Range("A5"), _
    TableName:="Sage_Sales_Data")
    
'Define the Pivotfields
Set S_T_InvoiceNo_PF = Sage_Data.PivotFields("Trimmed Invoice No")
Set S_Taxable_Sales_PF = Sage_Data.PivotFields("TaxableSalesAmt")
Set S_Freight_PF = Sage_Data.PivotFields("FreightAmt")
Set S_Tax_Amt_PF = Sage_Data.PivotFields("SalesTaxAmt")
Set S_TaxSch_PF = Sage_Data.PivotFields("TaxSchedule")

'Create Calculated Fields for Gross Sales and Sage Non-Taxable Sales that is inclusive of Tax
Set S_Non_Taxable_Sales_PF = Sage_Data.CalculatedFields.Add("Sage Gross_NonTax_(With Tax)", "='NonTaxableSalesAmt'+'SalesTaxAmt'")
Set S_Gross_Sales_PF = Sage_Data.CalculatedFields.Add("Sage Gross Sales_NonMktplc", "='TaxableSalesAmt'+'NonTaxableSalesAmt'+'FreightAmt'+'SalesTaxAmt'")


'I believe these variables may be redundant
S_Sum_Tx_Sales = "Sage Taxable Sales"
S_Sum_NTx_Sales = "Sage Non-Taxable Sales"
S_Sum_Freight_Sales = "Sage Freight"
S_Sum_Tax = "Sage Tax Amount"

'Set Pivot Table Fields
S_T_InvoiceNo_PF.Orientation = xlRowField
S_TaxSch_PF.Orientation = xlPageField
S_TaxSch_PF.EnableMultiplePageItems = True

'Isolate Non-Marketplace Sales
With S_TaxSch_PF
    On Error Resume Next
    .PivotItems("MKTPLCFAC").Visible = False
    
End With

'Populate the Pivot Data
Sage_Data.AddDataField S_Gross_Sales_PF, "Sage Gross Sales", xlSum
Sage_Data.AddDataField S_Taxable_Sales_PF, "Sage Taxable Sales", xlSum
Sage_Data.AddDataField S_Non_Taxable_Sales_PF, "Sage Non-Taxable Sales", xlSum
Sage_Data.AddDataField S_Freight_PF, "Sage Freight", xlSum
Sage_Data.AddDataField S_Tax_Amt_PF, "Sage Tax Amount", xlSum

'Activate the Sheet with the Pivot Table to define ranges
Sage_Pivot.Activate

'Define Ranges for Later Reference
Sage_Data.PivotSelect "Sage Gross Sales", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Gross = Selection
Sage_Data.PivotSelect "Sage Taxable Sales", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Taxable = Selection
Sage_Data.PivotSelect "Sage Non-Taxable Sales", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Nontaxable = Selection
Sage_Data.PivotSelect "Sage Freight", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Freight = Selection
Sage_Data.PivotSelect "Sage Tax Amount", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Tax = Selection
Sage_Data.rowRange.Select
Set Lookup_Sage_NonMarketplace_InvoiceNo = Selection

Sage_Pivot.Columns.AutoFit

' Create a Second Worksheet Called Sage Gross Compare
Set Sage_Gross = Taxbook.Sheets.Add
Sage_Gross.Name = "Sage_Gross Compare"

'Create another pivot cache for a Sage Pivot Table
Set Sage_Gross_Pivot = Sage_Data_Cache.CreatePivotTable( _
    TableDestination:=Sage_Gross.Range("A5"), _
    TableName:="Sage_GrossSales_Data")
    
' Define Pivotfields
Set SGP_InvoiceNo_PF = Sage_Gross_Pivot.PivotFields("Trimmed Invoice No")
Set SGP_Taxable_Sales_PF = Sage_Gross_Pivot.PivotFields("TaxableSalesAmt")
Set SGP_Nontaxable_Sales_PF = Sage_Gross_Pivot.PivotFields("Sage Gross_NonTax_(With Tax)")
Set SGP_Freight_PF = Sage_Gross_Pivot.PivotFields("FreightAmt")
Set SGP_Tax_Amt_PF = Sage_Gross_Pivot.PivotFields("SalesTaxAmt")
Set SGP_TaxSch_PF = Sage_Gross_Pivot.PivotFields("TaxSchedule")
Set SGP_Sage_Gross_Sales = Sage_Gross_Pivot.PivotFields("Sage Gross Sales_NonMktplc")

'Define Field Orientation
SGP_InvoiceNo_PF.Orientation = xlRowField
SGP_TaxSch_PF.Orientation = xlPageField
S_TaxSch_PF.EnableMultiplePageItems = True
    
'Populate the Pivot Data
Sage_Gross_Pivot.AddDataField SGP_Sage_Gross_Sales, "Sage Gross", xlSum
Sage_Gross_Pivot.AddDataField SGP_Taxable_Sales_PF, "Sage Taxable Sales", xlSum
Sage_Gross_Pivot.AddDataField SGP_Nontaxable_Sales_PF, "Sage Nontaxable Sales", xlSum
Sage_Gross_Pivot.AddDataField SGP_Freight_PF, "Sage Freight", xlSum
Sage_Gross_Pivot.AddDataField SGP_Tax_Amt_PF, "Sage Tax Amount", xlSum

'Activate the Sheet
Sage_Gross.Activate

'Define Ranges for Later Use
Sage_Gross_Pivot.PivotSelect "Sage Gross", xlDataAndLabel
Set Lookup_Sage_Gross = Selection
Sage_Gross_Pivot.PivotSelect "Sage Taxable Sales", xlDataAndLabel
Set Lookup_Sage_Taxable = Selection
Sage_Gross_Pivot.PivotSelect "Sage Nontaxable Sales", xlDataAndLabel
Set Lookup_Sage_Nontaxable = Selection
Sage_Gross_Pivot.PivotSelect "Sage Freight", xlDataAndLabel
Set Lookup_Sage_Freight = Selection
Sage_Gross_Pivot.PivotSelect "Sage Tax Amount", xlDataAndLabel
Set Lookup_SageTax = Selection
Sage_Gross_Pivot.rowRange.Select
Set Lookup_Sage_InvoiceNo = Selection
Application.Calculation = xlAutomatic

Else 'If Sales tax isn't included in gross sales then

'Create a Pivot Table of Sage Sales Data for comparison (excludes marketplace originated sales)

Application.Calculation = xlCalculationManual

' Set A range for the pivot table
Set DestRange = Sage_Pivot.Range("A5")

'Create a Sage Data Pivot Cache
Set Sage_Data_Cache = Taxbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=Range(Sage.Cells(1, 1), Sage.Cells(S_Last_Row, S_Last_Col + 1)))
    
'Create a Sage Data Pivot Table
Set Sage_Data = Sage_Data_Cache.CreatePivotTable( _
    TableDestination:=Sage_Pivot.Range("A5"), _
    TableName:="Sage_Sales_Data")

'Define Pivot Fields
Set S_T_InvoiceNo_PF = Sage_Data.PivotFields("Trimmed Invoice No")
Set S_Taxable_Sales_PF = Sage_Data.PivotFields("TaxableSalesAmt")
Set S_Non_Taxable_Sales_PF = Sage_Data.PivotFields("NonTaxableSalesAmt")
Set S_Freight_PF = Sage_Data.PivotFields("FreightAmt")
Set S_Tax_Amt_PF = Sage_Data.PivotFields("SalesTaxAmt")
Set S_TaxSch_PF = Sage_Data.PivotFields("TaxSchedule")

'Add a calculated field
Set S_Gross_Sales_PF = Sage_Data.CalculatedFields.Add("Sage Gross Sales_NonMktplc", "='TaxableSalesAmt'+'NonTaxableSalesAmt'+'FreightAmt'")

'Possibly redundant variables
S_Sum_Tx_Sales = "Sage Taxable Sales"
S_Sum_NTx_Sales = "Sage Non-Taxable Sales"
S_Sum_Freight_Sales = "Sage Freight"
S_Sum_Tax = "Sage Tax Amount"

'Define field orientations
S_T_InvoiceNo_PF.Orientation = xlRowField
S_TaxSch_PF.Orientation = xlPageField
S_TaxSch_PF.EnableMultiplePageItems = True

'Filter the Pivot Table to Exclude Marketplace Sales
With S_TaxSch_PF
    On Error Resume Next
    .PivotItems("MKTPLCFAC").Visible = False
    
End With

'Populate the Pivot Table with Gross, Taxable, Nontaxable, Freight, and Tax Amounts
Sage_Data.AddDataField S_Gross_Sales_PF, "Sage Gross Sales", xlSum
Sage_Data.AddDataField S_Taxable_Sales_PF, "Sage Taxable Sales", xlSum
Sage_Data.AddDataField S_Non_Taxable_Sales_PF, "Sage Non-Taxable Sales", xlSum
Sage_Data.AddDataField S_Freight_PF, "Sage Freight", xlSum
Sage_Data.AddDataField S_Tax_Amt_PF, "Sage Tax Amount", xlSum

'Activate the Sheet where the Pivot Table is stored
Sage_Pivot.Activate

'Define Ranges for Later Use
Sage_Data.PivotSelect "Sage Gross Sales", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Gross = Selection
Sage_Data.PivotSelect "Sage Taxable Sales", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Taxable = Selection
Sage_Data.PivotSelect "Sage Non-Taxable Sales", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Nontaxable = Selection
Sage_Data.PivotSelect "Sage Freight", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Freight = Selection
Sage_Data.PivotSelect "Sage Tax Amount", xlDataAndLabel
Set Lookup_Sage_NonMarketplace_Tax = Selection
Sage_Data.rowRange.Select
Set Lookup_Sage_NonMarketplace_InvoiceNo = Selection

'Fit Column Widths to Contents
Sage_Pivot.Columns.AutoFit


'Create a second sage pivot table on a sheet called sage gross compare that includes marketplace sales
'Create a worksheet to store the pivot table
Set Sage_Gross = Taxbook.Sheets.Add
Sage_Gross.Name = "Sage_Gross Compare"
Sage_Gross.Activate
Sage_Gross.Range("A7").Select
ActiveWindow.FreezePanes = True

'Create a pivot table of the Sage data
Set Sage_Gross_Pivot = Sage_Data_Cache.CreatePivotTable( _
    TableDestination:=Sage_Gross.Range("A5"), _
    TableName:="Sage_GrossSales_Data")
    
'Define the pivot fields
Set SGP_InvoiceNo_PF = Sage_Gross_Pivot.PivotFields("Trimmed Invoice No")
Set SGP_Taxable_Sales_PF = Sage_Gross_Pivot.PivotFields("TaxableSalesAmt")
Set SGP_Nontaxable_Sales_PF = Sage_Gross_Pivot.PivotFields("NonTaxableSalesAmt")
Set SGP_Freight_PF = Sage_Gross_Pivot.PivotFields("FreightAmt")
Set SGP_Tax_Amt_PF = Sage_Gross_Pivot.PivotFields("SalesTaxAmt")
Set SGP_TaxSch_PF = Sage_Gross_Pivot.PivotFields("TaxSchedule")

'Create a calculated field for gross sales
Set SGP_Sage_Gross_Sales = Sage_Gross_Pivot.CalculatedFields.Add("Sage Gross Sales", "='TaxableSalesAmt'+'NonTaxableSalesAmt'+'FreightAmt'")
    
'Orient the page fields
SGP_InvoiceNo_PF.Orientation = xlRowField
SGP_TaxSch_PF.Orientation = xlPageField
S_TaxSch_PF.EnableMultiplePageItems = True
    
'Populate the Pivot Table with Gross Sales, Taxable Sales, Nontaxable Sales, Freight, and Tax Amount
Sage_Gross_Pivot.AddDataField SGP_Sage_Gross_Sales, "Sage Gross", xlSum
Sage_Gross_Pivot.AddDataField SGP_Taxable_Sales_PF, "Sage Taxable Sales", xlSum
Sage_Gross_Pivot.AddDataField SGP_Nontaxable_Sales_PF, "Sage Nontaxable Sales", xlSum
Sage_Gross_Pivot.AddDataField SGP_Freight_PF, "Sage Freight", xlSum
Sage_Gross_Pivot.AddDataField SGP_Tax_Amt_PF, "Sage Tax Amount", xlSum

'Activate the worksheet the pivot table is stored on
Sage_Gross.Activate

'Define range variables for later reference
Sage_Gross_Pivot.PivotSelect "Sage Gross", xlDataAndLabel
Set Lookup_Sage_Gross = Selection
Sage_Gross_Pivot.PivotSelect "Sage Taxable Sales", xlDataAndLabel
Set Lookup_Sage_Taxable = Selection
Sage_Gross_Pivot.PivotSelect "Sage Nontaxable Sales", xlDataAndLabel
Set Lookup_Sage_Nontaxable = Selection
Sage_Gross_Pivot.PivotSelect "Sage Freight", xlDataAndLabel
Set Lookup_Sage_Freight = Selection
Sage_Gross_Pivot.PivotSelect "Sage Tax Amount", xlDataAndLabel
Set Lookup_SageTax = Selection
Sage_Gross_Pivot.rowRange.Select
Set Lookup_Sage_InvoiceNo = Selection

Application.Calculation = xlAutomatic

End If
End Sub

Sub Step_11_Main_Create_Vertex_Gross_Pivots()

 Dim Loopitem2 As PivotItem
 Dim VGP_TaxExempt As PivotField

'If Sales Tax Is to be included in sales amounts then:
If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then
Application.Calculation = xlCalculationAutomatic

Vertex_Pivot.Range("A4").Value = "Pivot of Vertex Sales Data for State level only to compare gross sales figures between datasets"

'Create a Pivot Table from the Vertex Data Cache
Set Vertex_Data = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_Pivot.Range("A5"), _
    TableName:="Vertex_Sales_Data")

'Define Pivotfields
Set V_Imposition_Type = Vertex_Data.PivotFields("Imposition Type")
Set V_Jurisdiction_Type = Vertex_Data.PivotFields("Jurisdiction Type")
Set V_FlexCode05 = Vertex_Data.PivotFields("Flex Code 5")
Set V_Product_Class_Code = Vertex_Data.PivotFields("Product Class Code")
Set V_Exempt_Amt_PF = Vertex_Data.PivotFields("Exempt Amount")
Set V_Taxable_Amt_PF = Vertex_Data.PivotFields("Taxable Amount")
Set V_NonTaxable_Amt_PF = Vertex_Data.PivotFields("Total Exempt Amount")
Set V_Tax_Amt_PF = Vertex_Data.PivotFields("Total Tax Amount")
Set V_InvoiceNo_PF = Vertex_Data.PivotFields("Trimmed Invoice No")

'Create Calculated Fields to include tax in gross amount
Set V_Gross_Amt_PF = Vertex_Data.CalculatedFields.Add("Gross Amount(With Tax)", "='Gross Amount'+ 'Total Tax Amount'")
Set V_TaxExempt = Vertex_Data.CalculatedFields.Add("Gross Nontaxable(All)", "='Exempt Amount'+ 'Non-Taxable Amount'+'Total Tax Amount'")
 
 ' Set orientations for the PivotFields
V_InvoiceNo_PF.Orientation = xlRowField
V_Jurisdiction_Type.Orientation = xlPageField
V_Imposition_Type.Orientation = xlPageField
V_Imposition_Type.EnableMultiplePageItems = True
V_Jurisdiction_Type.EnableMultiplePageItems = True

' Add data fields to the PivotTable
Vertex_Data.AddDataField V_Gross_Amt_PF, "Vertex Gross Sales", xlSum
Vertex_Data.AddDataField V_Taxable_Amt_PF, "Vertex Taxable Sale", xlSum
Vertex_Data.AddDataField V_TaxExempt, "Vertex Exempt Sales", xlSum

'Filter the Data for State Jurisdiction Type
For Each Loopitem2 In V_Jurisdiction_Type.PivotItems
    If Loopitem2.Name <> "STATE" Then
        Loopitem2.Visible = False
    End If
Next Loopitem2

'Filter the Data for General Sales and Use Tax imposition type
For Each Loopitem2 In V_Imposition_Type.PivotItems
    If Loopitem2.Name <> "General Sales and Use Tax" Then
        Loopitem2.Visible = False
    End If
Next Loopitem2

'Activate the sheet, and define ranges for future references
Vertex_Pivot.Activate
Vertex_Data.PivotSelect "Vertex Gross Sales", xlDataAndLabel
Set Lookup_Vertex_Gross_Gross = Selection
Vertex_Data.PivotSelect "Vertex Taxable Sale", xlDataAndLabel
Set Lookup_Vertex_Gross_Taxable = Selection
Vertex_Data.PivotSelect "Vertex Exempt Sales", xlDataAndLabel
Set Lookup_Vertex_Gross_Exempt = Selection
Vertex_Data.rowRange.Select
Set Lookup_Vertex_Gross_InvoiceNo = Selection

'Create a new worksheet and activate it for a second pivot table
Set Vertex_Gross_Sheet = Taxbook.Sheets.Add
Vertex_Gross_Sheet.Name = "Vertex_Gross Compare"
Vertex_Gross_Sheet.Activate
Vertex_Gross_Sheet.Range("A7").Select
ActiveWindow.FreezePanes = True

'Create a Second Pivot Table in a new sheet called "Vertex_Gross_Compare"
Set Vertex_Gross_Pivot = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_Gross_Sheet.Range("A5"), _
    TableName:="Vertex_GrossSales_Data")
    
'Define Pivotfields for the second pivot table
On Error Resume Next
Set VGP_TaxExempt = Vertex_Gross_Pivot.PivotFields("Total Exempt Amount")
Set VGP_Imposition_Type = Vertex_Gross_Pivot.PivotFields("Imposition Type")
Set VGP_Jurisdiction_Type = Vertex_Gross_Pivot.PivotFields("Jurisdiction Type")
Set VGP_Flexcode05 = Vertex_Gross_Pivot.PivotFields("Flex Code 5")
Set VGP_Product_Class_Code = Vertex_Gross_Pivot.PivotFields("Product Class Code")
Set VGP_Exempt_Amount = Vertex_Gross_Pivot.PivotFields("Exempt Amount")
Set VGP_Taxable_Sales_PF = Vertex_Gross_Pivot.PivotFields("Taxable Amount")
Set VGP_Nontaxable_Sales_PF = Vertex_Gross_Pivot.PivotFields("Non-Taxable Amount")
Set VGP_Tax_Amt_PF = Vertex_Gross_Pivot.PivotFields("Tax Amount")
Set VGP_Gross_Sales_PF = Vertex_Gross_Pivot.PivotFields("Gross Amount(With Tax)")
Set VGP_InvoiceNo_PF = Vertex_Gross_Pivot.PivotFields("Trimmed Invoice No")
 
 ' Set orientations for the PivotFields
VGP_InvoiceNo_PF.Orientation = xlRowField
VGP_Jurisdiction_Type.Orientation = xlPageField
VGP_Imposition_Type.Orientation = xlPageField
VGP_Jurisdiction_Type.EnableMultiplePageItems = True
    
Vertex_Gross_Pivot.AddDataField VGP_Gross_Sales_PF, "Vertex_Gross_Amt", xlSum
Vertex_Gross_Pivot.AddDataField VGP_Taxable_Sales_PF, "Vertex_Taxable_Amt", xlSum
Vertex_Gross_Pivot.AddDataField VGP_TaxExempt, "Vertex Exempt Amount", xlSum

'Check to see if the recon sate was TN, if it was not TN that add an additional data field for tax amount
If ReconState <> "TN" Then
Vertex_Gross_Pivot.AddDataField VGP_Tax_Amt_PF, "Vertex_Tax_Amt", xlSum
Else
End If

'Filter the pivot data for jurisdiction type state
For Each LoopItem In VGP_Jurisdiction_Type.PivotItems
    If LoopItem.Name <> "STATE" Then
    LoopItem.Visible = False
    End If
Next LoopItem

'Filter the pivot data to exclude Additional Sales and Use Tax
For Each LoopItem In VGP_Imposition_Type.PivotItems
    If LoopItem.Name = "Additional Sales and Use Tax" Then
        LoopItem.Visible = False
        Else
        LoopItem.Visible = True
        End If
Next LoopItem
    
'Define Ranges for Later Reference
Vertex_Gross_Pivot.PivotSelect "Vertex_Gross_Amt", xlDataAndLabel
Set Lookup_Vertex_GrossComp_Gross = Selection
Vertex_Gross_Pivot.PivotSelect "Vertex_Taxable_Amt", xlDataAndLabel
Set Lookup_Vertex_GrossComp_Taxable = Selection
Vertex_Gross_Pivot.PivotSelect "Vertex Exempt Amount", xlDataAndLabel
Set Lookup_Vertex_GrossComp_Exempt = Selection
Vertex_Gross_Pivot.rowRange.Select
Set Lookup_Vertex_GrossComp_InvoiceNo = Selection


'Create a Worksheet called Vertex_Nontaxable"
Set Vertex_NonTax_Compare = Taxbook.Sheets.Add
Vertex_NonTax_Compare.Name = "Vertex_Nontaxable"
Vertex_NonTax_Compare.Activate
Vertex_NonTax_Compare.Range("A7").Select
ActiveWindow.FreezePanes = True

'Create a third pivot table from vertex data called Vertex_NonTax_Sales_Data

Set Vertex_NonTaxable_Pivot = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_NonTax_Compare.Range("A5"), _
    TableName:="Vertex_NonTax_Sales_Data")

Vertex_NonTaxable_Pivot.RowAxisLayout xlTabularRow
Vertex_NonTaxable_Pivot.RepeatAllLabels xlRepeatLabels

'Define Pivotfields for vertex nontaxable pivot
Set VNTP_Imposition_Type = Vertex_NonTaxable_Pivot.PivotFields("Imposition Type")
Set VNTP_Jurisdiction_Type = Vertex_NonTaxable_Pivot.PivotFields("Jurisdiction Type")
Set VNTP_Flexcode05 = Vertex_NonTaxable_Pivot.PivotFields("Flex Code 5")
Set VNTP_Product_Class_Code = Vertex_NonTaxable_Pivot.PivotFields("Product Class Code")
Set VNTP_Exempt_Amount = Vertex_NonTaxable_Pivot.PivotFields("Exempt Amount")
Set VNTP_Taxable_Sales_PF = Vertex_NonTaxable_Pivot.PivotFields("Taxable Amount")
Set VNTP_Nontaxable_Sales_PF = Vertex_NonTaxable_Pivot.PivotFields("Non-Taxable Amount")
Set VNTP_Tax_Amt_PF = Vertex_NonTaxable_Pivot.PivotFields("Tax Amount")
Set VNTP_Gross_Sales_PF = Vertex_NonTaxable_Pivot.PivotFields("Gross Amount(With Tax)")
Set VNTP_InvoiceNo_PF = Vertex_NonTaxable_Pivot.PivotFields("Trimmed Invoice No")
Set VNTP_Product_Class = Vertex_NonTaxable_Pivot.PivotFields("Product Class Code")
Set VNTP_Gross_Nontax_PF = Vertex_NonTaxable_Pivot.PivotFields("Total Exempt Amount")
Set VNTP_PostalCode = Vertex_NonTaxable_Pivot.PivotFields("Destination Postal Code")
Set VNTP_SitusCity = Vertex_NonTaxable_Pivot.PivotFields("Situs City Name")

'Configure the pivot table subtotals and field orientation
VNTP_InvoiceNo_PF.Orientation = xlRowField
VNTP_InvoiceNo_PF.Subtotals(1) = True
VNTP_InvoiceNo_PF.Subtotals(1) = False
VNTP_PostalCode.Orientation = xlRowField
VNTP_PostalCode.Subtotals(1) = True
VNTP_PostalCode.Subtotals(1) = False
VNTP_SitusCity.Orientation = xlRowField
VNTP_SitusCity.Subtotals(1) = True
VNTP_SitusCity.Subtotals(1) = False
VNTP_Jurisdiction_Type.Orientation = xlPageField
VNTP_Jurisdiction_Type.EnableMultiplePageItems = True

'Check to see if the Recon State was Colorado, if it was not then filter imposition type to include General Sales and Use Tax only
If ReconState = "CO" Then
VNTP_Imposition_Type.Orientation = xlPageField
VNTP_Imposition_Type.EnableMultiplePageItems = True
    For Each LoopItem_CO In VNTP_Imposition_Type.PivotItems
            If LoopItem_CO.Name <> "General Sales and Use Tax" Then
                LoopItem_CO.Visible = False
            End If
    Next LoopItem_CO
End If

'Check to see if recon state is TN, if so then exclude additional sales and use tax from imposition type
If ReconState = "TN" Then
    VNTP_Imposition_Type.Orientation = xlPageField
    For Each LoopItem In VNTP_Imposition_Type.PivotItems
            If LoopItem.Name = "Additional Sales and Use Tax" Then
                LoopItem.Visible = False
        End If
Next LoopItem
    End If
    
'Filter the nontaxable pivot to include only jurisdiction type "STATE"
For Each VNTP_Loopitem In VNTP_Jurisdiction_Type.PivotItems

    If VNTP_Loopitem <> "STATE" Then
        VNTP_Loopitem.Visible = False
    End If
Next VNTP_Loopitem

VNTP_Product_Class.Orientation = xlPageField
VNTP_Product_Class.EnableMultiplePageItems = True

'Make Sure Freight is Included in the data (Why???)
For Each VNTP_Loopitem In VNTP_Product_Class.PivotItems
    If VNTP_Loopitem = "FREIGHT" Then
    VNTP_Loopitem.Visible = True
    End If
Next VNTP_Loopitem

VNTP_Flexcode05.Orientation = xlPageField
VNTP_Flexcode05.EnableMultiplePageItems = True


'Make Sure Marketplace Freight is Included in the data (Why???)
For Each VNTP_Loopitem In VNTP_Flexcode05.PivotItems
    If VNTP_Loopitem = "FREIGHT" Then
    VNTP_Loopitem.Visible = True
    End If
Next VNTP_Loopitem
            
 ' Populate Nontaxable Sales Pivot data
Vertex_NonTaxable_Pivot.AddDataField VNTP_Gross_Sales_PF, "Vertex_Gross_Amt", xlSum
Vertex_NonTaxable_Pivot.AddDataField VNTP_Taxable_Sales_PF, "Vertex_Taxable_Amt", xlSum
Vertex_NonTaxable_Pivot.AddDataField VNTP_Gross_Nontax_PF, "Vertex_Gross Nontaxable", xlSum

'Possibly Redundant but for each Pivotfield turn off subtotaling in the pivot table
For Each VNTP_LoopField In Vertex_NonTaxable_Pivot.PivotFields
    On Error Resume Next
    VNTP_LoopField.Subtotals(1) = True
    VNTP_LoopField.Subtotals(1) = False
Next VNTP_LoopField

'Check again if Recon State was TN
If ReconState = "TN" Then

'Filter data for TN to exclude Additional Sales and Use Tax (Redundent?)
For Each LoopItem In Vertex_NonTaxable_Pivot.PivotItems
    If LoopItem.Name = "Additional Sales and Use Tax" Then
    LoopItem.Visible = False
    End If
Next LoopItem

Else: End If

'Define Ranges for Later Reference
Vertex_NonTaxable_Pivot.PivotSelect "Vertex_Gross_Amt", xlDataAndLabel
Set Lookup_Vertex_Exempt_Gross = Selection
Vertex_NonTaxable_Pivot.PivotSelect "Vertex_Taxable_Amt", xlDataAndLabel
Set Lookup_Vertex_Exempt_Taxable = Selection
Vertex_NonTaxable_Pivot.PivotSelect "Vertex_Gross Nontaxable", xlDataAndLabel
Set Lookup_Vertex_Exempt_Exempt = Selection
Vertex_NonTaxable_Pivot.PivotFields("Trimmed Invoice No").DataRange.Select
Selection.Offset(Rowoffset:=-1).Select
Selection.Resize(RowSize:=WorksheetFunction.CountA(Selection) + 2).Select
Set Lookup_Vertex_Exempt_InvoiceNo = Selection
    


Else ' If Sales Tax Is Not included in Gross Sales then:

'Create a pivot table of vertex sales, filtered by state and imposition type general sales and use tax with invoice number as the primary key

    Application.Calculation = xlCalculationAutomatic
    
'Create a Pivot Table from the Vertex Data Cache

    Vertex_Pivot.Range("A4").Value = "Pivot of Vertex Sales Data for State level only to compare gross sales figures between datasets"
    Set Vertex_Data = Vertex_Data_Cache.CreatePivotTable( _
        TableDestination:=Vertex_Pivot.Range("A5"), _
        TableName:="Vertex_Sales_Data")
    
  'Define Pivotfields
  
    Set V_Imposition_Type = Vertex_Data.PivotFields("Imposition Type")
    Set V_Jurisdiction_Type = Vertex_Data.PivotFields("Jurisdiction Type")
    Set V_FlexCode05 = Vertex_Data.PivotFields("Flex Code 5")
    Set V_Product_Class_Code = Vertex_Data.PivotFields("Product Class Code")
    Set V_Exempt_Amt_PF = Vertex_Data.PivotFields("Exempt Amount")
    Set V_Taxable_Amt_PF = Vertex_Data.PivotFields("Taxable Amount")
    Set V_NonTaxable_Amt_PF = Vertex_Data.PivotFields("Non-Taxable Amount")
    Set V_Tax_Amt_PF = Vertex_Data.PivotFields("Tax Amount")
    Set V_Gross_Amt_PF = Vertex_Data.PivotFields("Gross Amount")
    Set V_InvoiceNo_PF = Vertex_Data.PivotFields("Trimmed Invoice No")
    
    ' Add a calculated field to combine exempt and nontaxable vertex sales
    Set V_TaxExempt = Vertex_Data.CalculatedFields.Add("Gross Nontaxable(All)", "='Exempt Amount'+ 'Non-Taxable Amount'")
    
    ' Set the orientation of the pivot fields
    V_InvoiceNo_PF.Orientation = xlRowField
    V_Jurisdiction_Type.Orientation = xlPageField
    V_Imposition_Type.Orientation = xlPageField
    V_Imposition_Type.EnableMultiplePageItems = True
    V_Jurisdiction_Type.EnableMultiplePageItems = True
    
    ' Add data fields to the PivotTable
    Vertex_Data.AddDataField V_Gross_Amt_PF, "Vertex Gross Sales", xlSum
    Vertex_Data.AddDataField V_Taxable_Amt_PF, "Vertex Taxable Sale", xlSum
    Vertex_Data.AddDataField V_TaxExempt, "Vertex Exempt Sales", xlSum
        
    'Filter Jurisdiction Type to be "STATE"
    For Each Loopitem2 In V_Jurisdiction_Type.PivotItems
        If Loopitem2.Name <> "STATE" Then
            Loopitem2.Visible = False
        End If
    Next Loopitem2
    
    'Filter Imposition Type to be "General Sales and Use Tax"
    For Each Loopitem2 In V_Imposition_Type.PivotItems
        If Loopitem2.Name <> "General Sales and Use Tax" Then
            Loopitem2.Visible = False
        End If
    Next Loopitem2
    
    'Activate the Worksheet with the pivot table on it
    Vertex_Pivot.Activate
    
    'Define Ranges for later reference
    Vertex_Data.PivotSelect "Vertex Gross Sales", xlDataAndLabel
    Set Lookup_Vertex_Gross_Gross = Selection
    Vertex_Data.PivotSelect "Vertex Taxable Sale", xlDataAndLabel
    Set Lookup_Vertex_Gross_Taxable = Selection
    Vertex_Data.PivotSelect "Vertex Exempt Sales", xlDataAndLabel
    Set Lookup_Vertex_Gross_Exempt = Selection
    Vertex_Data.rowRange.Select
    Set Lookup_Vertex_Gross_InvoiceNo = Selection
    
'Create a new worksheet called Vertex Gross Compare. This is a second pivot table created to run data comparison analysis on.
    Set Vertex_Gross_Sheet = Taxbook.Sheets.Add
    Vertex_Gross_Sheet.Name = "Vertex_Gross Compare"
    Vertex_Gross_Sheet.Activate
    Vertex_Gross_Sheet.Range("A7").Select
    ActiveWindow.FreezePanes = True
    
    'Create a pivot table from the vertex pivot cache
    
    Set Vertex_Gross_Pivot = Vertex_Data_Cache.CreatePivotTable( _
        TableDestination:=Vertex_Gross_Sheet.Range("A5"), _
        TableName:="Vertex_GrossSales_Data")
        
    'Define Pivotfields
    On Error Resume Next
    Set VGP_TaxExempt = Vertex_Gross_Pivot.CalculatedFields.Add("Gross Nontaxable(Compare)", "='Exempt Amount'+ 'Non-Taxable Amount'")
    Set VGP_Imposition_Type = Vertex_Gross_Pivot.PivotFields("Imposition Type")
    Set VGP_Jurisdiction_Type = Vertex_Gross_Pivot.PivotFields("Jurisdiction Type")
    Set VGP_Flexcode05 = Vertex_Gross_Pivot.PivotFields("Flex Code 5")
    Set VGP_Product_Class_Code = Vertex_Gross_Pivot.PivotFields("Product Class Code")
    Set VGP_Exempt_Amount = Vertex_Gross_Pivot.PivotFields("Exempt Amount")
    Set VGP_Taxable_Sales_PF = Vertex_Gross_Pivot.PivotFields("Taxable Amount")
    Set VGP_Nontaxable_Sales_PF = Vertex_Gross_Pivot.PivotFields("Non-Taxable Amount")
    Set VGP_Tax_Amt_PF = Vertex_Gross_Pivot.PivotFields("Tax Amount")
    Set VGP_Gross_Sales_PF = Vertex_Gross_Pivot.PivotFields("Gross Amount")
    Set VGP_InvoiceNo_PF = Vertex_Gross_Pivot.PivotFields("Trimmed Invoice No")
        
    ' Set the orientation of the pivot fields
      
    VGP_InvoiceNo_PF.Orientation = xlRowField
    VGP_Jurisdiction_Type.Orientation = xlPageField
    VGP_Imposition_Type.Orientation = xlPageField
    VGP_Jurisdiction_Type.EnableMultiplePageItems = True
        
    'Populate data fields
    Vertex_Gross_Pivot.AddDataField VGP_Gross_Sales_PF, "Vertex_Gross_Amt", xlSum
    Vertex_Gross_Pivot.AddDataField VGP_Taxable_Sales_PF, "Vertex_Taxable_Amt", xlSum
    Vertex_Gross_Pivot.AddDataField VGP_TaxExempt, "Vertex Exempt Amount", xlSum
    
    'Check to see if the Recon State was Colorado, if its not, then filter the pivot table for Imposition Type General Sales and Use Tax
    If ReconState = "CO" Then
    
    For Each Loopitem4 In VGP_Imposition_Type.PivotItems
        If Loopitem4.Name <> "General Sales and Use Tax" Then
        Loopitem4.Visible = False
        End If
    Next Loopitem4
      
    End If
    
    'Check to see if the Reconstate is TN, if its not TN then add a data field for tax amount
    If ReconState <> "TN" Then
    Vertex_Gross_Pivot.AddDataField VGP_Tax_Amt_PF, "Vertex_Tax_Amt", xlSum
    Else
    End If
    
    'Filter the Pivot table to include only Jurisdiction type State
    For Each LoopItem In VGP_Jurisdiction_Type.PivotItems
        If LoopItem.Name <> "STATE" Then
        LoopItem.Visible = False
        End If
    Next LoopItem
    
    'Loop through the imposition types to exclude Additional Sales and Use Tax and Additional Fee
    For Each LoopItem In VGP_Imposition_Type.PivotItems
        If LoopItem.Name = "Additional Sales and Use Tax" Or LoopItem.Name = "Additional Fee" Then
            LoopItem.Visible = False
            Else
            LoopItem.Visible = True
            End If
    Next LoopItem
        
    'Define Ranges for Later Use
    Vertex_Gross_Pivot.PivotSelect "Vertex_Gross_Amt", xlDataAndLabel
    Set Lookup_Vertex_GrossComp_Gross = Selection
    Vertex_Gross_Pivot.PivotSelect "Vertex_Taxable_Amt", xlDataAndLabel
    Set Lookup_Vertex_GrossComp_Taxable = Selection
    Vertex_Gross_Pivot.PivotSelect "Vertex Exempt Amount", xlDataAndLabel
    Set Lookup_Vertex_GrossComp_Exempt = Selection
    Vertex_Gross_Pivot.rowRange.Select
    Set Lookup_Vertex_GrossComp_InvoiceNo = Selection
    
    
 'Create a second Vertex Pivot Table for Nontaxable Sales only in a sheet called "Vertex_Nontaxable" that includes freight sales
 
'Create a Worksheet called "Vertex_NonTaxable"

    Set Vertex_NonTax_Compare = Taxbook.Sheets.Add
    Vertex_NonTax_Compare.Name = "Vertex_Nontaxable"
    
'Create a third Vertex Based Pivot Table for nontaxable sales

    Set Vertex_NonTaxable_Pivot = Vertex_Data_Cache.CreatePivotTable( _
        TableDestination:=Vertex_NonTax_Compare.Range("A5"), _
        TableName:="Vertex_NonTax_Sales_Data")
     
 'Modify the Pivot Table Settings to be tabular layout and to repeat label rows.
    Vertex_NonTaxable_Pivot.RowAxisLayout xlTabularRow
    Vertex_NonTaxable_Pivot.RepeatAllLabels xlRepeatLabels
    
    
   'Define Pivotfields
    Set VNTP_Imposition_Type = Vertex_NonTaxable_Pivot.PivotFields("Imposition Type")
    Set VNTP_Jurisdiction_Type = Vertex_NonTaxable_Pivot.PivotFields("Jurisdiction Type")
    Set VNTP_Flexcode05 = Vertex_NonTaxable_Pivot.PivotFields("Flex Code 5")
    Set VNTP_Product_Class_Code = Vertex_NonTaxable_Pivot.PivotFields("Product Class Code")
    Set VNTP_Exempt_Amount = Vertex_NonTaxable_Pivot.PivotFields("Exempt Amount")
    Set VNTP_Taxable_Sales_PF = Vertex_NonTaxable_Pivot.PivotFields("Taxable Amount")
    Set VNTP_Nontaxable_Sales_PF = Vertex_NonTaxable_Pivot.PivotFields("Non-Taxable Amount")
    Set VNTP_Tax_Amt_PF = Vertex_NonTaxable_Pivot.PivotFields("Tax Amount")
    Set VNTP_Gross_Sales_PF = Vertex_NonTaxable_Pivot.PivotFields("Gross Amount")
    Set VNTP_InvoiceNo_PF = Vertex_NonTaxable_Pivot.PivotFields("Trimmed Invoice No")
    Set VNTP_Product_Class = Vertex_NonTaxable_Pivot.PivotFields("Product Class Code")
    Set VNTP_Gross_Nontax_PF = Vertex_NonTaxable_Pivot.CalculatedFields.Add("Gross Nontaxable(NonTaxComp)", "='Exempt Amount'+ 'Non-Taxable Amount'")
    Set VNTP_PostalCode = Vertex_NonTaxable_Pivot.PivotFields("Destination Postal Code")
    Set VNTP_SitusCity = Vertex_NonTaxable_Pivot.PivotFields("Situs City Name")
    
    ' Set the orientation of the pivot fields and Turn Off Subtotaling (This seems to use Legacy Zipcode Import)
    
    VNTP_InvoiceNo_PF.Orientation = xlRowField
    VNTP_InvoiceNo_PF.Subtotals(1) = True
    VNTP_InvoiceNo_PF.Subtotals(1) = False
    VNTP_PostalCode.Orientation = xlRowField
    VNTP_PostalCode.Subtotals(1) = True
    VNTP_PostalCode.Subtotals(1) = False
    VNTP_SitusCity.Orientation = xlRowField
    VNTP_SitusCity.Subtotals(1) = True
    VNTP_SitusCity.Subtotals(1) = False
    VNTP_Jurisdiction_Type.Orientation = xlPageField
    VNTP_Jurisdiction_Type.EnableMultiplePageItems = True
    VNTP_Product_Class.Orientation = xlPageField
    VNTP_Product_Class.EnableMultiplePageItems = True
    VNTP_Flexcode05.Orientation = xlPageField
    VNTP_Flexcode05.EnableMultiplePageItems = True
    
'Check to see if the recon state was Colorado, if so then filter imposition type to include only general sales and use tax
     If ReconState = "CO" Then
        VNTP_Imposition_Type.Orientation = xlPageField
        For Each Loopitem6 In VNTP_Imposition_Type.PivotItems
                If Loopitem6.Name <> "General Sales and Use Tax" Then
                    Loopitem6.Visible = False
            End If
        Next Loopitem6
        End If
            
'Check to see if the reconstate was TN, if so filter Imposition Type to not include Additional Sales and Use Tax
    If ReconState = "TN" Then
        VNTP_Imposition_Type.Orientation = xlPageField
        For Each LoopItem In VNTP_Imposition_Type.PivotItems
                If LoopItem.Name = "Additional Sales and Use Tax" Then
                    LoopItem.Visible = False
            End If
    Next LoopItem
        End If
        
    'Filter the Pivot Table  to exclude jurisdictions that are not "STATE"
    For Each VNTP_Loopitem In VNTP_Jurisdiction_Type.PivotItems
    
        If VNTP_Loopitem <> "STATE" Then
            VNTP_Loopitem.Visible = False
        End If
    Next VNTP_Loopitem

    'Make sure Freight is included in the pivot data
    For Each VNTP_Loopitem In VNTP_Product_Class.PivotItems
        If VNTP_Loopitem = "FREIGHT" Then
        VNTP_Loopitem.Visible = True
        End If
    Next VNTP_Loopitem
    

    'Make sure Marketplace Freight is included in the pivot data
    For Each VNTP_Loopitem In VNTP_Flexcode05.PivotItems
        If VNTP_Loopitem = "FREIGHT" Then
        VNTP_Loopitem.Visible = True
        End If
    Next VNTP_Loopitem
    
    'Populate the pivot datra
                
    Vertex_NonTaxable_Pivot.AddDataField VNTP_Gross_Sales_PF, "Vertex_Gross_Amt", xlSum
    Vertex_NonTaxable_Pivot.AddDataField VNTP_Taxable_Sales_PF, "Vertex_Taxable_Amt", xlSum
    Vertex_NonTaxable_Pivot.AddDataField VNTP_Gross_Nontax_PF, "Vertex_Gross Nontaxable", xlSum
    
    'Turn off Subtotaling for each pivotfield
    For Each VNTP_LoopField In Vertex_NonTaxable_Pivot.PivotFields
        On Error Resume Next
        VNTP_LoopField.Subtotals(1) = True
        VNTP_LoopField.Subtotals(1) = False
    Next VNTP_LoopField
    
    
    'Check to see if ReconState is TN
    If ReconState = "TN" Then
    
    'If TN then filter out Additional Sales and Use Tax  (redundant, also incorrect reference?)
    For Each LoopItem In VNTP_Imposition_Type
        If LoopItem.Name = "Additional Sales and Use Tax" Then
        LoopItem.Visible = False
        End If
    Next LoopItem
    
    Else: End If
    
  'Define Ranges for Future Reference
    
    Vertex_NonTaxable_Pivot.PivotSelect "Vertex_Gross_Amt", xlDataAndLabel
    Set Lookup_Vertex_Exempt_Gross = Selection
    Vertex_NonTaxable_Pivot.PivotSelect "Vertex_Taxable_Amt", xlDataAndLabel
    Set Lookup_Vertex_Exempt_Taxable = Selection
    Vertex_NonTaxable_Pivot.PivotSelect "Vertex_Gross Nontaxable", xlDataAndLabel
    Set Lookup_Vertex_Exempt_Exempt = Selection
    Vertex_NonTaxable_Pivot.PivotFields("Trimmed Invoice No").DataRange.Select
    Selection.Offset(Rowoffset:=-1).Select
    Selection.Resize(RowSize:=WorksheetFunction.CountA(Selection) + 2).Select
    Set Lookup_Vertex_Exempt_InvoiceNo = Selection
End If
    End Sub

Sub Step_10_Main_Create_Vertex_Freight_Sales_Pivot()

'Creates a Pivot Table of Vertex Freight Sales on Sheet "Vertex Pivot". Part of the original code, additional vertex tables have since been added.

' Declare variables
Dim VFS_TaxExempt As PivotField

'Create the PivotTable from the Vertex Data Cache

Set Vertex_Freight_Sales = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_Pivot.Range("N5"), _
    TableName:="Vertex_Freight_Sales_Data")
  
' Set PivotFields

Set VFS_InvoiceNo = Vertex_Freight_Sales.PivotFields("Trimmed Invoice No")
Set VFS_Product_Class_Code = Vertex_Freight_Sales.PivotFields("Product Class Code")
Set VFS_Tax_Amount = Vertex_Freight_Sales.PivotFields("Tax Amount")
Set VFS_Flexcode05 = Vertex_Freight_Sales.PivotFields("Flex Code 5")
Set VFS_Gross_Amt = Vertex_Freight_Sales.PivotFields("Gross Amount")
Set VFS_Exempt_Amt = Vertex_Freight_Sales.PivotFields("Exempt Amount")
Set VFS_NonTax_Amt = Vertex_Freight_Sales.PivotFields("Non-Taxable Amount")
Set VFS_Taxable_Amt = Vertex_Freight_Sales.PivotFields("Taxable Amount")
Set VFS_Jurisdiction_Type = Vertex_Freight_Sales.PivotFields("Jurisdiction Type")
Set VFS_ImpositionType = Vertex_Freight_Sales.PivotFields("Imposition Type")

' Add a calculated field to Add Vertex Nontaxable and Exempt Fields to get total exempt and nontaxable sales

Set VFS_TaxExempt = Vertex_Freight_Sales.CalculatedFields.Add("Gross Nontaxable(Freight Only)", "='Exempt Amount'+ 'Non-Taxable Amount'")

' Set orientations of the fields

VFS_InvoiceNo.Orientation = xlRowField
VFS_Product_Class_Code.Orientation = xlPageField
VFS_Flexcode05.Orientation = xlPageField
VFS_Jurisdiction_Type.Orientation = xlPageField
VFS_ImpositionType.Orientation = xlPageField

' Filter ImpositionType field to include only General Sales and Use Tax
With VFS_ImpositionType
    .EnableMultiplePageItems = True
    On Error Resume Next
    .PivotItems("Additional Sales and Use Tax").Visible = False
    .PivotItems("General Sales and Use Tax").Visible = True
    .PivotItems("Additional Fee").Visible = False
End With

' Add data fields to the PivotTable
With Vertex_Freight_Sales
    .AddDataField VFS_Gross_Amt, "Vertex Gross Freight", xlSum
    .AddDataField VFS_Taxable_Amt, "Vertex Taxable Freight", xlSum
    .AddDataField VFS_TaxExempt, "Vertex Freight Exempt Tax", xlSum
End With

Dim PivotLoopItem As PivotItem
VFS_Jurisdiction_Type.EnableMultiplePageItems = True
VFS_Product_Class_Code.EnableMultiplePageItems = True
VFS_Jurisdiction_Type.EnableMultiplePageItems = True
 
 ' Filter the Jurisdiction_Type field to be State Only
For Each PivotLoopItem In VFS_Jurisdiction_Type.PivotItems
    If PivotLoopItem <> "STATE" Then
    PivotLoopItem.Visible = False
    Else
    PivotLoopItem.Visible = True
    End If
Next PivotLoopItem

 ' Set pivot title
Vertex_Pivot.Range("N5").Value = "Pivot of Sales that are freight related. Excludes anything except freight."

' Filter the Product_Class_Code field to exclude any transactions that weren't marketplace sales or freight sales
For Each PivotLoopItem In VFS_Product_Class_Code.PivotItems

    If _
    PivotLoopItem = "FREIGHT" _
        Or PivotLoopItem = "SEAR046_exempt" _
            Or PivotLoopItem = "AMAZ002_exempt" _
                Or PivotLoopItem = "EBAY001_exempt" _
                    Or PivotLoopItem = "WALM020_exempt" _
                        Or PivotLoopItem = "EBAY001_exempt" _
    Then
    PivotLoopItem.Visible = True
    Else
    PivotLoopItem.Visible = False
    End If
    
Next PivotLoopItem

'Filter the Flexcode05 fields to retain only marketplace freight to end with a pivot table report that includes only freight sales
VFS_Flexcode05.EnableMultiplePageItems = True

For Each PivotLoopItem In VFS_Flexcode05.PivotItems
    If PivotLoopItem = "FREIGHT" Or PivotLoopItem = "(blank)" Then
    PivotLoopItem.Visible = True
    Else
    PivotLoopItem.Visible = False
    End If
Next PivotLoopItem

' Activate the sheet to define ranges for future reference
Sheets("Vertex Pivot").Activate
    
' Define Ranges for Later Data Comparison (using the .PivotSelect method)
Vertex_Freight_Sales.PivotSelect "Vertex Gross Freight", xlDataAndLabel
Set Lookup_Vertex_Freight_Gross = Selection
Vertex_Freight_Sales.PivotSelect "Vertex Taxable Freight", xlDataAndLabel
Set Lookup_Vertex_Freight_Taxable = Selection
Vertex_Freight_Sales.PivotSelect "Vertex Freight Exempt Tax", xlDataAndLabel
Set Lookup_Vertex_Freight_Exempt = Selection
Vertex_Freight_Sales.rowRange.Select
Set Lookup_Vertex_Freight_Sales_InvoiceNo = Selection


End Sub


Sub Step_12_Create_Vertex_Tax_Pivot()

    ' If the recon state is Colorado, an additional process of determining the
    ' home rule tax is needed
    If ReconState = "CO" Then

        ' Declare variables for Colorado home rule tax
        Dim Datafile As String ' Directory location of the CO Administered City Tax File
        Dim HRLookup As Worksheet ' Sheet where the names of the cities where CO collects tax for the city
        Dim CopyFromSheet As Worksheet ' Sheet within the Home Rule City Data File
        Dim CopyToBook As Workbook ' Workbook where the home rule city list will be copied
        Dim HomeRuleList As Workbook ' Workbook that the home rule city data is stored in

        ' Set up home rule city data
        Set CopyToBook = ActiveWorkbook ' Set the CopytoBook to be the tax recon book
        Set HRLookup = CopyToBook.Sheets.Add ' Add a sheet into the recon file to hold home rule city data
        HRLookup.Name = "CO HR Cities" ' Rename the home rule city sheet to be CO HR Cities

        ' Prompt user to import CO administered city list
        MsgBox ("Import CO administered city list...")
        Datafile = Application.GetOpenFilename() ' Query User where the home rule cities data is

        ' Open home rule city data file and copy data
        Set HomeRuleList = Application.Workbooks.Open(Datafile) ' Set the HomeRuleList Workbook to be the queried file's windows directory
        Set CopyFromSheet = HomeRuleList.ActiveSheet ' Define the Copyfromsheet to be the home rule list
        CopyFromSheet.UsedRange.Copy
        HRLookup.Range("A1").PasteSpecial xlPasteAll
        HomeRuleList.Close

        ' Create home rule city tax pivot table
        Set Homerule = Taxbook.Sheets.Add
        Homerule.Name = "HR City Tax"

' Create Vertex Pivot Tables from Vertex Pivot Caches
Set Vertex_Tax_COHR = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Homerule.Range("A5"), _
    TableName:="Vertex_HRTax_Data")
        
Set Vertex_Tax = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_Pivot.Range("I5"), _
    TableName:="Vertex_Tax_Data")
         
' Define Pivot Fields for Vertex Tax Pivot
Set VT_FlexCode05 = Vertex_Tax.PivotFields("Flex Code 5")
Set VT_Imposition_Type = Vertex_Tax.PivotFields("Imposition Type")
Set VT_Product_Class_Code = Vertex_Tax.PivotFields("Product Class Code")
Set VT_Tax_InvNo = Vertex_Tax.PivotFields("Trimmed Invoice No")
Set VT_Tax = Vertex_Tax.PivotFields("Tax Amount")
Set VT_TAXJUR = Vertex_Tax.PivotFields("Jurisdiction Type")

'Define Pivot Fields for Vertex Tax Pivot HR Jurisdictions

Set VTC_FlexCode05 = Vertex_Tax_COHR.PivotFields("Flex Code 5")
Set VTC_Imposition_Type = Vertex_Tax_COHR.PivotFields("Imposition Type")
Set VTC_Product_Class_Code = Vertex_Tax_COHR.PivotFields("Product Class Code")
Set VTC_Tax_InvNo = Vertex_Tax_COHR.PivotFields("Trimmed Invoice No")
Set VTC_Tax = Vertex_Tax_COHR.PivotFields("Tax Amount")
Set VTC_TAXJUR = Vertex_Tax_COHR.PivotFields("Jurisdiction Type")
Set VTC_TAXCity = Vertex_Tax_COHR.PivotFields("Jurisdiction Name")

'Orient the page items and fields in the pivot table
VTC_TAXJUR.EnableMultiplePageItems = True
VTC_TAXCity.EnableMultiplePageItems = True
VTC_TAXCity.Orientation = xlRowField
VTC_TAXJUR.Orientation = xlPageField

'Create a calculated field in the city tax pivot table
Vertex_Tax_COHR.AddDataField VTC_Tax, "Vertex Gross Tax"
VT_Tax_InvNo.Orientation = xlRowField
Vertex_Tax.AddDataField VT_Tax, "Vertex Gross Tax"

'Filter the Home Rule Pivot to include only city jurisdictions
With VTC_TAXJUR
    On Error Resume Next
    .PivotItems("STATE").Visible = False
    On Error Resume Next
    .PivotItems("COUNTY").Visible = False
    On Error Resume Next
    .PivotItems("SALES TAX DISTRICT").Visible = False
    On Error Resume Next
    .PivotItems("LOCAL IMPROVEMENT DISTRICT").Visible = False
    On Error Resume Next
    .PivotItems("CITY").Visible = True
End With

Dim Value_1 As String
Dim Value_2 As String
Dim Item As PivotItem

' Loop Through the cities in the pivot items and filter them out if they are NOT home rule jurisdictions
For Each Item In VTC_TAXCity.PivotItems
    Value_2 = WorksheetFunction.IfError(WorksheetFunction.CountIf(HRLookup.Range("A:A"), Item.Name), 0)
    If Value_2 > 0 Then
    Item.Visible = True
    Else
    Item.Visible = False
    End If
Next Item

Else

'If there is no colorado home rule jurisdiction to worry about then:

' Create a pivot table of vertex tax per document number on the sheet "Vertex Pivot".

Vertex_Pivot.Range("I4").Value = "Pivot of total tax charged (local and state) by related invoice txn no."
Set Vertex_Tax = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_Pivot.Range("I5"), _
    TableName:="Vertex_Tax_Data")
    
Set VT_FlexCode05 = Vertex_Tax.PivotFields("Flex Code 5")
Set VT_Imposition_Type = Vertex_Tax.PivotFields("Imposition Type")
Set VT_Product_Class_Code = Vertex_Tax.PivotFields("Product Class Code")
Set VT_Tax_InvNo = Vertex_Tax.PivotFields("Trimmed Invoice No")
Set VT_Tax = Vertex_Tax.PivotFields("Tax Amount")

VT_Tax_InvNo.Orientation = xlRowField
Vertex_Tax.AddDataField VT_Tax, "Vertex Gross Tax"

End If
Vertex_Pivot.Activate

Vertex_Tax.PivotSelect "Vertex Gross Tax", xlDataAndLabel
Selection.Offset(Rowoffset:=-1).Select
Selection.Resize(RowSize:=WorksheetFunction.CountA(Selection) + 1).Select
Set Lookup_Vertex_Tax_Tax = Selection
Vertex_Tax.rowRange.Select
Set Lookup_Vertex_Tax_InvoiceNo = Selection

End Sub

Sub Step_14_Create_Mktplc_Sage_Pivot()

'Define Variables
Dim Mktplc_LC As Long
Dim Mktplc_LR As Long

'Determine whether tax should be included within gross sales
If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then

'Create a Pivot Table from Sage Data Cache
Set MktplcPiv = Sage_Data_Cache.CreatePivotTable( _
    TableDestination:=Mktplc.Range("A5"), _
    TableName:="Sage_Mktplc_Sales_Data")
    
'Define Pivot Fields
    
Set M_T_InvoiceNo_PF = MktplcPiv.PivotFields("Trimmed Invoice No")
Set M_Taxable_Sales_PF = MktplcPiv.PivotFields("TaxableSalesAmt")
Set M_Non_Taxable_Sales_PF = MktplcPiv.PivotFields("Sage Gross_NonTax_(With Tax)")
Set M_Freight_PF = MktplcPiv.PivotFields("FreightAmt")
Set M_Tax_Amt_PF = MktplcPiv.PivotFields("SalesTaxAmt")
Set M_TaxSch_PF = MktplcPiv.PivotFields("TaxSchedule")
Set MKTPLC_Imposition_Type = MktplcPiv.PivotFields("Imposition Type")
Set M_Gross_Sales_PF = MktplcPiv.PivotFields("Sage Gross Sales_NonMktplc")

'Orient Pivot Fields
M_T_InvoiceNo_PF.Orientation = xlRowField
M_TaxSch_PF.Orientation = xlPageField
M_TaxSch_PF.EnableMultiplePageItems = True

'If Recon State is CO then Filter for only General Sales and Use Tax
If ReconState = "CO" Then
MKTPLC_Imposition_Type.Orientation = xlPageField
MKTPLC_Imposition_Type.EnableMultiplePageItems = True
For Each Loopitem5 In MKTPLC_Imposition_Type.PivotItems
    If Loopitem5.Name <> "General Sales and Use Tax" Then
        Loopitem5.Visible = False
    End If
Next Loopitem5
End If
Set Loopitem5 = Nothing

' Filter Sage Data for Marketplace Facilitator Transactions
For Each Loopitem5 In M_TaxSch_PF.PivotItems
    If Loopitem5 <> "MKTPLCFAC" Then
    Loopitem5.Visible = False
End If
Next Loopitem5

'Populate the pivot table with Gross Sales, Taxable Sales, NonTaxable Sales, Freight Sales, and Tax Amount
MktplcPiv.AddDataField M_Gross_Sales_PF, "Sage Gross Mkplc Sales", xlSum
MktplcPiv.AddDataField M_Taxable_Sales_PF, "Sage Taxable Sales", xlSum
MktplcPiv.AddDataField M_Non_Taxable_Sales_PF, "Sage Non-Taxable Sales", xlSum
MktplcPiv.AddDataField M_Freight_PF, "Sage Freight", xlSum
MktplcPiv.AddDataField M_Tax_Amt_PF, "Sage Tax Amount", xlSum
    
'Activate the worksheet that the pivot table was stored on
Mktplc.Activate


'Define Range variables for future reference
MktplcPiv.PivotSelect "Sage Gross Mkplc Sales", xlDataAndLabel
Set Lookup_Sage_Marketplace_Gross = Selection
MktplcPiv.PivotSelect "Sage Taxable Sales", xlDataAndLabel
Set Lookup_Sage_Marketplace_Taxable = Selection
MktplcPiv.PivotSelect "Sage Non-Taxable Sales", xlDataAndLabel
Set Lookup_Sage_Marketplace_Nontaxable = Selection
MktplcPiv.PivotSelect "Sage Freight", xlDataAndLabel
Set Lookup_Sage_Marketplace_Freight = Selection
MktplcPiv.PivotSelect "Sage Tax Amount", xlDataAndLabel
Set Lookup_Sage_Marketplace_Tax = Selection
MktplcPiv.rowRange.Select
Set Lookup_Sage_Marketplace_InvoiceNo = Selection

Else ' If Tax is not to be included in gross sales then...

'Create a pivot table from sage pivot data cache
Set MktplcPiv = Sage_Data_Cache.CreatePivotTable( _
    TableDestination:=Mktplc.Range("A5"), _
    TableName:="Sage_Mktplc_Sales_Data")
    
'Define Pivotfields

Set M_T_InvoiceNo_PF = MktplcPiv.PivotFields("Trimmed Invoice No")
Set M_Taxable_Sales_PF = MktplcPiv.PivotFields("TaxableSalesAmt")
Set M_Non_Taxable_Sales_PF = MktplcPiv.PivotFields("NonTaxableSalesAmt")
Set M_Freight_PF = MktplcPiv.PivotFields("FreightAmt")
Set M_Tax_Amt_PF = MktplcPiv.PivotFields("SalesTaxAmt")
Set M_TaxSch_PF = MktplcPiv.PivotFields("TaxSchedule")

'Create Calculated Field for Gross Marketplace Sales
Set M_Gross_Sales_PF = MktplcPiv.CalculatedFields.Add("Sage Gross(Mktplc)", "='TaxableSalesAmt'+'NonTaxableSalesAmt'+'FreightAmt'")


'Orient the pivot fields
M_T_InvoiceNo_PF.Orientation = xlRowField
M_TaxSch_PF.Orientation = xlPageField
M_TaxSch_PF.EnableMultiplePageItems = True

'Filter the Pivot table for marketplace facilitated transactions
For Each Loopitem5 In M_TaxSch_PF.PivotItems
    If Loopitem5 <> "MKTPLCFAC" Then
    Loopitem5.Visible = False
End If
Next Loopitem5


'Populate the pivot table data with Gross Sales, Taxable Sales, NonTaxable Sales, Freight Sales, Tax Amount
MktplcPiv.AddDataField M_Gross_Sales_PF, "Sage Gross Mkplc Sales", xlSum
MktplcPiv.AddDataField M_Taxable_Sales_PF, "Sage Taxable Sales", xlSum
MktplcPiv.AddDataField M_Non_Taxable_Sales_PF, "Sage Non-Taxable Sales", xlSum
MktplcPiv.AddDataField M_Freight_PF, "Sage Freight", xlSum
MktplcPiv.AddDataField M_Tax_Amt_PF, "Sage Tax Amount", xlSum
    
'Define Range Variables for later reference

Mktplc.Activate
MktplcPiv.PivotSelect "Sage Gross Mkplc Sales", xlDataAndLabel
Set Lookup_Sage_Marketplace_Gross = Selection
MktplcPiv.PivotSelect "Sage Taxable Sales", xlDataAndLabel
Set Lookup_Sage_Marketplace_Taxable = Selection
MktplcPiv.PivotSelect "Sage Non-Taxable Sales", xlDataAndLabel
Set Lookup_Sage_Marketplace_Nontaxable = Selection
MktplcPiv.PivotSelect "Sage Freight", xlDataAndLabel
Set Lookup_Sage_Marketplace_Freight = Selection
MktplcPiv.PivotSelect "Sage Tax Amount", xlDataAndLabel
Set Lookup_Sage_Marketplace_Tax = Selection
MktplcPiv.rowRange.Select
Set Lookup_Sage_Marketplace_InvoiceNo = Selection

End If


End Sub
Sub Step_16_Secondary_5_Loop_And_Link_Section_1()
Error_Report.Activate
Trimmed_Invoice_Col = WorksheetFunction.Match("Trimmed Invoice No", Sage.Range("1:1"), False)
Trimmed_Invoice_Col_Letter = Split(Cells(1, Trimmed_Invoice_Col).Address, "$")(1)
Lookuplocation = Trimmed_Invoice_Col_Letter & ":" & Trimmed_Invoice_Col_Letter

'loop through all the errors and lookup their information in save
For i = SageCompareStart + 3 To SageEnd

    
    ' Link Column A to the raw data for each document number
            
            Set LinkCell = Error_Report.Range("A" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!A" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
        'Lookup Invoice Date from Sage AR Data
        
            Error_Report.Range("B" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$E:$E," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            Error_Report.Range("B" & i).NumberFormat = "mm/dd/yyyy"
        
            
        'Link Column B (Invoice Date) to Source Data
            
            Set LinkCell = Error_Report.Range("B" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!E" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup Customer Number based on Invoice Number from Sage AR Data
            Error_Report.Range("C" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$H:$H," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
        
            'Link Column C (Customer No.) to Source Data
        
            Set LinkCell = Error_Report.Range("C" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!H" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup Vertex Customer name based on Invoice Number from Exemption Report
            Error_Report.Range("D" & i).Value = "=xlookup(C" & i & ",'Exemption Report'!B:B" & ",'Exemption Report'!$D:$D," & Chr(34) & Chr(34) & ",0)"
            
            'Link Column D (Vertex Customer Name) to Source Data
            
            Set LinkCell = Error_Report.Range("D" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("C" & i).Value & ",'Exemption Report'!B:B,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "''Exemption Report'!D" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup ShipToState from Sage AR Data (Error Report Column E)
            Error_Report.Range("E" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AQ:$AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
            
            'Link Column E (Ship To State) to Source Data
            
            Set LinkCell = Error_Report.Range("E" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "''Sage AR Data'!AQ" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup ShipToZipCode from Sage AR Data (Column F)
            Error_Report.Range("F" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AR:$AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
            'Link Column F (Ship To Zipcode) to Source Data
            Set LinkCell = Error_Report.Range("F" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AR" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
             
          'Lookup TaxSchedule from Sage AR Data (Column G)
           Error_Report.Range("G" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$J:$J," & Chr(34) & "ERROR" & Chr(34) & ",0)"
           
          'Link Column G (Tax Schedule) to Source Data
            Set LinkCell = Error_Report.Range("G" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!J" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
          'Lookup Order Manager from Sage AR Data (Column H)
          Error_Report.Range("H" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$CY:$CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
        
          'Link Column H (Order Manager) to Source Data (Column H)
            Set LinkCell = Error_Report.Range("H" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!CY" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
                                                      
          'Lookup ShipToCity Column I) from Sage AR Data
          Error_Report.Range("I" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AP:$AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
          
            'Link ShipToCity to Source Data (Column I)
            Set LinkCell = Error_Report.Range("I" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
           LinkAddress = "'Sage AR Data'!AP" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
          'Lookup ShiptoAddress 1 from Sage AR Data (Column J)
           Error_Report.Range("J" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AM:$AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
                                 
          'Link Column J (Ship To Address 1) to Source Data (Column J)
            Set LinkCell = Error_Report.Range("J" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AM" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
        
          'Lookup Ship To Address 2 from Sage AR Data (Column K)
           Error_Report.Range("K" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!AN:$AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
          'Link Column J (Ship To Address 2) to Source Data (Column K)
            Set LinkCell = Error_Report.Range("K" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AN" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
        'Lookup Ship To Address 3 from Sage AR Data (Column L)
        Error_Report.Range("L" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!AO:$AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
        'Link Column J (Ship To Address 3) to Source Data (Column L)
            Set LinkCell = Error_Report.Range("L" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AO" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
        
        'Lookup Exemption Reason from Exemption Report (Column M)
        
        Error_Report.Range("M" & i).Value = "=xlookup(C" & i & ",'Exemption Report'!B:B,'Exemption Report'!F:F," & Chr(34) & "ERROR" & Chr(34) & ",0)&" & Chr(34) & "-" & Chr(34) & "&" & "xlookup(C" & i & ",'Exemption Report'!B:B,'Exemption Report'!M:M," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
        
        'Link Column M (Exemption Reason) to Source Data (Column M)
            Set LinkCell = Error_Report.Range("M" & i)
            MatchString = "IFError(MATCH(" & Chr(34) & Error_Report.Range("C" & i).Value & Chr(34) & ",'Exemption Report'!B:B,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Exemption Report'!F" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
        
        'Link Column N to Source Data (Sage Gross Sales)
            
            Set LinkCell = Error_Report.Range("N" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage_Gross Compare'!B" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
         
           'Link Column O to Source Data (Sage Taxable Sales)
            
            Set LinkCell = Error_Report.Range("O" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage_Gross Compare'!C" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Link Column P to Source Data (Sage Nontaxble Sales)
            
            Set LinkCell = Error_Report.Range("P" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage_Gross Compare'!D" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Link Column Q to Source Data (Sage Freight)
            Set LinkCell = Error_Report.Range("Q" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage_Gross Compare'!E" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Link Column R to Source Data (Sage Tax)
            
            Set LinkCell = Error_Report.Range("R" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage_Gross Compare'!F" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Link Column S to Source Data (Vertex Gross)
            
            Set LinkCell = Error_Report.Range("S" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Vertex_Gross Compare'!B" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
             End If
            
            'Link Column U to Source Data (Vertex Taxable Sales)
            
            Set LinkCell = Error_Report.Range("U" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Vertex_Gross Compare'!C" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Link Column W to Source Data (Vertex Non-Taxable Sales)
            
            Set LinkCell = Error_Report.Range("W" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Vertex_Gross Compare'!D" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
           'Link Column Y to Source Data (Vertex Freight Sales)
           
            Set LinkCell = Error_Report.Range("Y" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Vertex_Gross Compare'!F" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Link Column AA to Source Data (Vertex Tax)

            Set LinkCell = Error_Report.Range("AA" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Vertex_Gross Compare'!G" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
   Next i
       
  Step_16_Secondary_6_Fill_Notes_Section_1

End Sub
Sub Step_16_Secondary_6_Fill_Notes_Section_1()

        ' Attempt to Determine Reason for Tax Variance.
   For i = SageCompareStart + 3 To SageEnd
            
            Dim Message As String
            
            
         'Check to see if a  transaction was created with a default tax schedule and check to see if there is a valid exemption on file. If there is a valid exemption of file fill in the note:
        If Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> "ERROR-ERROR" Or Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> "" Then
        Message = "Transaction had Tax Schedule Default and an exemption on file for " & Error_Report.Range("M" & i).Value & " as of this reconciliation indicating a correcting transaction needs posted to vertex. Potentially refund tax as appropriate."
        
        Error_Report.Range("AC" & i).Value = "Transaction had Tax Schedule Default and an exemption on file for " & Error_Report.Range("M" & i).Value & " as of this reconciliation indicating a correcting transaction needs posted to vertex. Potentially refund tax as appropriate."
        
        'If there is a transaction with Tax Schedule Default and No Exemption on File fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value = "ERROR-ERROR" Then
        Message = "Transaction had Tax Schedule Default But no exemption of file. Make sure to update the vertex configuration report. If the report is up to date then customer was likely issued a refund without a certificate on file otherwise the configuration report needs updated. Contact Sales/Cust Service for resolution."
        
         Error_Report.Range("AC" & i).Value = "Transaction had Tax Schedule Default But no exemption of file. Make sure to update the vertex configuration report. If the report is up to date then customer was likely issued a refund without a certificate on file otherwise the configuration report needs updated. Contact Sales/Cust Service for resolution."
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the tax amount per the document, then whether Vertex tax was higher then sage tax. If so fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("R" & i).Value <> Error_Report.Range("AA" & i).Value _
        And Error_Report.Range("AA" & i).Value > Error_Report.Range("R" & i).Value _
        And Error_Report.Range("AA" & i).Value > 0 _
        And Error_Report.Range("R" & i).Value > 0 _
        Then
        Message = "The tax calculations between Sage and Vertex do not agree. While both have tax, Vertex produced a higher tax calculation. This could be a timing difference in a tax rate change, or some other error. The rule of thumb is to use Vertex as the guide so the transactional tax in Sage for this invoice should be adjusted to match."
        Error_Report.Range("AC" & i).Value = "The tax calculations between Sage and Vertex do not agree. While both have tax, Vertex produced a higher tax calculation. This could be a timing difference in a tax rate change, or some other error. The rule of thumb is to use Vertex as the guide so the transactional tax in Sage for this invoice should be adjusted to match."
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the tax amount per the document, then whether Vertex tax was higher then sage tax. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("R" & i).Value <> Error_Report.Range("AA" & i).Value _
        And Error_Report.Range("AA" & i).Value < Error_Report.Range("R" & i).Value _
        And Error_Report.Range("AA" & i).Value > 0 _
        And Error_Report.Range("R" & i).Value > 0 _
        Then
        Message = "The tax calculations between Sage and Vertex do not agree. While both have tax, Vertex produced a lower tax calculation. This could be a timing difference in a tax rate change, or some other error. The rule of thumb is to use Vertex as the guide so the transactional tax in Sage for this invoice should be adjusted to match."
        Error_Report.Range("AC" & i).Value = "The tax calculations between Sage and Vertex do not agree. While both have tax, Vertex produced a lower tax calculation. This could be a timing difference in a tax rate change, or some other error. The rule of thumb is to use Vertex as the guide so the transactional tax in Sage for this invoice should be adjusted to match."
        
    'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the tax amount per the document, then whether Vertex tax was 0 and sage tax greater then 0. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("R" & i).Value <> Error_Report.Range("AA" & i).Value _
        And Error_Report.Range("AA" & i).Value > 0 _
        And Error_Report.Range("R" & i).Value = 0 _
        Then
        Message = "The tax calculations between Sage and Vertex do not agree. Only Vertex returned tax. This could be a timing difference due to the expiration of a tax certificate or uploading of a tax certificate. Generally, Sage uses Sales Order Date for its tax logic when determining tax due, and Vertex uses document date. If the certificate was expired prior the document date, but not the sales order date, then Vertex would charge tax and sage would not. Check the status of the customer and their exemption to determine next steps."
        Error_Report.Range("AC" & i).Value = "The tax calculations between Sage and Vertex do not agree. Only Vertex returned tax. This could be a timing difference due to the expiration of a tax certificate or uploading of a tax certificate. Generally, Sage uses Sales Order Date for its tax logic when determining tax due, and Vertex uses document date. If the certificate was expired prior the document date, but not the sales order date, then Vertex would charge tax and sage would not. Check the status of the customer and their exemption to determine next steps."
        
End If
        
'Attempt to determine reason for gross sales variance
       
               'Check to see if a  transaction was created with a default tax schedule and check to see if there is a valid exemption on file. If there is a valid exemption of file fill in the note:
        If Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> "ERROR-ERROR" _
        Or Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> 0 _
        And Error_Report.Range("T" & i).Value <> 0 Then
            Message = "Transaction had Tax Schedule Default and an exemption on file for " & Error_Report.Range("M" & i).Value & " as of this reconciliation indicating a correcting transaction needs posted to vertex. The variance in gross sales may be because this refund is for a prior period tax. Investigate the originating document and determine if a tax refund is appropriate."
        
                If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
                            Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & " Transaction had Tax Schedule Default and an exemption on file for " & Error_Report.Range("M" & i).Value & " as of this reconciliation indicating a correcting transaction needs posted to vertex. The variance in gross sales may be because this refund is for a prior period tax. Investigate the originating document and determine if a tax refund is appropriate.")
                End If
                
        'If there is a transaction with Tax Schedule Default and No Exemption on File fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value = "ERROR-ERROR" And Error_Report.Range("T" & i).Value <> 0 Then
         
      If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
         Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & " Transaction had Tax Schedule Default But no exemption of file. Make sure to update the vertex configuration report and rerun the reconciliation. If the report is up to date then customer was likely issued a refund without a certificate on file or had a default tax schedule assigned to their ship to address, or otherwise the configuration report needs updated. Contact Sales/Cust Service for resolution if the report is up to date to obtain a cert from the customer.")
        End If
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the gross amount per the document, then whether Sage gross was higher then Vertex gross. If so fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("N" & i).Value <> Error_Report.Range("S" & i).Value _
        And Error_Report.Range("N" & i).Value > Error_Report.Range("S" & i).Value _
        And Error_Report.Range("N" & i).Value > 0 _
        And Error_Report.Range("S" & i).Value > 0 _
        Then
        
        If Error_Report.Range("AC" & i).Value <> "" Then
             Else
                 Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Gross Sales Amount between vertex and sage disagree. Sage has a higher value then vertex, but both have an amount. Check that any adjusting transactions made were posted correctly for this document.")
        End If
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the gross amount per the document, then whether Vertex gross was higher then sage gross. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("N" & i).Value <> Error_Report.Range("S" & i).Value _
        And Error_Report.Range("N" & i).Value < Error_Report.Range("S" & i).Value _
        And Error_Report.Range("N" & i).Value > 0 _
        And Error_Report.Range("S" & i).Value > 0 _
        Then
        
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Gross Sales Amount between vertex and sage disagree. Vertex has a higher value then Sage, but both have an amount. Check that any adjusting transactions made were posted correctly for this document. ")
        End If
        
    'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the tax amount per the document, then whether Vertex tax was 0 and sage tax greater then 0. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("N" & i).Value <> Error_Report.Range("S" & i).Value _
        And Error_Report.Range("N" & i).Value > 0 _
        And Error_Report.Range("S" & i).Value = 0 _
        Then
        
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Gross Sales Amount between vertex and sage disagree. Vertex has a null value but sage has an amount. Check that the ship to state in sage is correct. If correct, check to see if customer zip code is jurisdiction STATE DEPARTMENT which are not included in vertex gross sales.")
        
       End If
        End If
        
'Attempt to determine reason for taxable sales variance


          'Check to see if a  transaction was created with a default tax schedule and check to see if there is a valid exemption on file. If there is a valid exemption of file fill in the note:
        If Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> "ERROR-ERROR" _
        Or Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> 0 _
        And Error_Report.Range("T" & i).Value <> 0 Then
     If Error_Report.Range("AC" & i).Value <> "" Then
        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & " Transaction had Tax Schedule Default and an exemption on file for " & Error_Report.Range("M" & i).Value & " as of this reconciliation indicating a correcting transaction needs posted to vertex. The variance in gross sales may be because this refund is for a prior period tax. Investigate the originating document and determine if a tax refund is appropriate.")
    End If
        
        'If there is a transaction with Tax Schedule Default and No Exemption on File fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value = "ERROR-ERROR" And Error_Report.Range("T" & i).Value <> 0 Then
        
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
         Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & " Transaction had Tax Schedule Default But no exemption of file. Make sure to update the vertex configuration report and rerun the reconciliation. If the report is up to date then customer was likely issued a refund without a certificate on file or had a default tax schedule assigned to their ship to address, or otherwise the configuration report needs updated. Contact Sales/Cust Service for resolution if the report is up to date to obtain a cert from the customer.")
        
        End If
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the gross amount per the document, then whether Sage gross was higher then Vertex gross. If so fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("O" & i).Value <> Error_Report.Range("U" & i).Value _
        And Error_Report.Range("O" & i).Value > Error_Report.Range("U" & i).Value _
        And Error_Report.Range("O" & i).Value > 0 _
        And Error_Report.Range("U" & i).Value > 0 _
        Then
        If Error_Report.Range("AC" & i).Value <> "" Then
            Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Taxable Sales Amount between vertex and sage disagree. Sage has a higher value then vertex, but both have an amount. Check that any adjusting transactions made were posted correctly for this document.")
        End If
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the gross amount per the document, then whether Vertex gross was higher then sage gross. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("O" & i).Value <> Error_Report.Range("U" & i).Value _
        And Error_Report.Range("O" & i).Value < Error_Report.Range("U" & i).Value _
        And Error_Report.Range("O" & i).Value > 0 _
        And Error_Report.Range("U" & i).Value > 0 _
        Then
        
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Taxable Sales Amount between vertex and sage disagree. Vertex has a higher value then Sage, but both have an amount. Check that any adjusting transactions made were posted correctly for this document. ")
        End If
        
    'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the tax amount per the document, then whether Vertex tax was 0 and sage tax greater then 0. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("O" & i).Value <> Error_Report.Range("U" & i).Value _
        And Error_Report.Range("O" & i).Value > 0 _
        And Error_Report.Range("U" & i).Value = 0 _
        Then
        If Error_Report.Range("AC" & i).Value <> "" Then
            Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Taxable Sales Amount between vertex and sage disagree. Vertex has a null value but sage has an amount. This indicates a timing issue with a certificate. Check the certificate to make sure it was active for the sales order date")
      End If

    'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the tax amount per the document, then whether Vertex tax was 0 and sage tax greater then 0. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("O" & i).Value <> Error_Report.Range("U" & i).Value _
        And Error_Report.Range("U" & i).Value > 0 _
        And Error_Report.Range("O" & i).Value = 0 _
        Then
        If Error_Report.Range("AC" & i).Value <> "" Then
         Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Taxable Sales Amount between vertex and sage disagree. Vertex has a  value but sage has an null amount. This indicates a timing issue with a certificate. Check the certificate to make sure it was active for the sales order date or did not expire between the sales order date and invoice date")
        End If
        End If

'Attempt to determine reason for Nontaxable Sale Variance

'Attempt to determine reason for taxable sales variance
          'Check to see if a  transaction was created with a default tax schedule and check to see if there is a valid exemption on file. If there is a valid exemption of file fill in the note:
        If Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> "ERROR-ERROR" _
        Or Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value <> 0 _
        And Error_Report.Range("T" & i).Value <> 0 Then
                
If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & " Transaction had Tax Schedule Default and an exemption on file for " & Error_Report.Range("M" & i).Value & " as of this reconciliation indicating a correcting transaction needs posted to vertex. The variance in gross sales may be because this refund is for a prior period tax. Investigate the originating document and determine if a tax refund is appropriate.")
 End If
    
        
        'If there is a transaction with Tax Schedule Default and No Exemption on File fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "DEFAULT" And Error_Report.Range("M" & i).Value = "ERROR-ERROR" And Error_Report.Range("T" & i).Value <> 0 Then
         Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & " Transaction had Tax Schedule Default But no exemption of file. Make sure to update the vertex configuration report and rerun the reconciliation. If the report is up to date then customer was likely issued a refund without a certificate on file or had a default tax schedule assigned to their ship to address, or otherwise the configuration report needs updated. Contact Sales/Cust Service for resolution if the report is up to date to obtain a cert from the customer.")
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the gross amount per the document, then whether Sage gross was higher then Vertex gross. If so fill in the note
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("P" & i).Value <> Error_Report.Range("W" & i).Value _
        And Error_Report.Range("P" & i).Value > Error_Report.Range("W" & i).Value _
        And Error_Report.Range("P" & i).Value > 0 _
        And Error_Report.Range("W" & i).Value > 0 _
        Then
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Nontaxable or Exempt Sales Amount between vertex and sage disagree. Sage has a higher value then vertex, but both have an amount. Check that any adjusting transactions made were posted correctly for this document. Check for a certificate timing issue.")
        End If
        
        'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the gross amount per the document, then whether Vertex gross was higher then sage gross. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("P" & i).Value <> Error_Report.Range("W" & i).Value _
        And Error_Report.Range("P" & i).Value < Error_Report.Range("W" & i).Value _
        And Error_Report.Range("P" & i).Value > 0 _
        And Error_Report.Range("W" & i).Value > 0 _
        Then
        
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Nontaxable or Exempt Sales Amount between vertex and sage disagree. Vertex has a higher value then Sage, but both have an amount. Check that any adjusting transactions made were posted correctly for this document. Check for certificate timing differences.")
        
        End If
        
    'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the exempt amount per the document, then whether Vertex exempt was 0 and sage tax greater then 0. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("P" & i).Value <> Error_Report.Range("W" & i).Value _
        And Error_Report.Range("P" & i).Value > 0 _
        And Error_Report.Range("W" & i).Value = 0 _
        Then
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Nontaxable or Exempt Sales Amount between vertex and sage disagree. Vertex has a null value but sage has an amount. This indicates a timing issue with a certificate. Check the certificate to make sure it was active for the sales order date")
    
    End If

    'Check to see if the Tax Schedule is Vertex, and whether there is a variance in the exempt amount per the document, then whether Vertex exempt was 0 and sage tax greater then 0. If so fill in the note as appropriate to the conditions
        ElseIf Error_Report.Range("G" & i).Value = "VERTEX" _
        And Error_Report.Range("P" & i).Value <> Error_Report.Range("W" & i).Value _
        And Error_Report.Range("W" & i).Value > 0 _
        And Error_Report.Range("P" & i).Value = 0 _
        Then
        If Error_Report.Range("AC" & i).Value <> "" Then
                        Else
        Error_Report.Range("AC" & i).Value = WorksheetFunction.Trim(Error_Report.Range("AC" & i).Value & "The Nontaxable or Exempt Sales Amount between vertex and sage disagree. Vertex has a  value but sage has an null amount. This indicates a timing issue with a certificate. Check the certificate to make sure it was active for the sales order date or did not expire between the sales order date and invoice date")
        
      End If
        End If
               
Next i

For i = SageCompareStart + 3 To SageEnd
       'Change Number Format to Accounting Format
        Error_Report.Range("N" & i & ":" & "AB" & i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
            
        If Error_Report.Range("K" & i).Value = 0 Then
        Error_Report.Range("K" & i).Value = ""
        End If
        If Error_Report.Range("L" & i).Value = 0 Then
        Error_Report.Range("L" & i).Value = ""
        End If
        If Error_Report.Range("M" & i).Value = 0 Then
        Error_Report.Range("M" & i).Value = ""
        End If
        
Next i


End Sub


Sub Step_16_Secondary_7_Loop_And_Link_Section_2()
For i = MarketplaceStart To MarketplaceEnd - 1
 If MarketplaceStart = MarketplaceEnd Then
 Else
       ' Link Column A to Source Data
            
            Set LinkCell = Error_Report.Range("A" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!A" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
        'Lookup Invoice Date from Sage AR Data
        
            Error_Report.Range("B" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$E:$E," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            Error_Report.Range("B" & i).NumberFormat = "mm/dd/yyyy"
        
            
        'Link Column B (Invoice Date) to Source Data
            
            Set LinkCell = Error_Report.Range("B" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!E" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup Customer Number based on Invoice Number from Sage AR Data
            Error_Report.Range("C" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$H:$H," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
        
            'Link Column C (Customer No.) to Source Data
        
            Set LinkCell = Error_Report.Range("C" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!H" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup Vertex Customer name based on Invoice Number from Exemption Report
            Error_Report.Range("D" & i).Value = "=xlookup(C" & i & ",'Exemption Report'!B:B" & ",'Exemption Report'!$D:$D," & Chr(34) & Chr(34) & ",0)"
            
            'Link Column D (Vertex Customer Name) to Source Data
            
            Set LinkCell = Error_Report.Range("D" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("C" & i).Value & ",'Exemption Report'!B:B,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "''Exemption Report'!D" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup ShipToState from Sage AR Data (Error Report Column E)
            Error_Report.Range("E" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AQ:$AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
            
            'Link Column E (Ship To State) to Source Data
            
            Set LinkCell = Error_Report.Range("E" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "''Sage AR Data'!AQ" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
            'Lookup ShipToZipCode from Sage AR Data (Column F)
            Error_Report.Range("F" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AR:$AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
            'Link Column F (Ship To Zipcode) to Source Data
            Set LinkCell = Error_Report.Range("F" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AR" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
             
          'Lookup TaxSchedule from Sage AR Data (Column G)
           Error_Report.Range("G" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$J:$J," & Chr(34) & "ERROR" & Chr(34) & ",0)"
           
          'Link Column G (Tax Schedule) to Source Data
            Set LinkCell = Error_Report.Range("G" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!J" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
          'Lookup Order Manager from Sage AR Data (Column H)
          Error_Report.Range("H" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$CY:$CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
        
          'Link Column H (Order Manager) to Source Data (Column H)
            Set LinkCell = Error_Report.Range("H" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!CY" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
                                                      
          'Lookup ShipToCity Column I) from Sage AR Data
          Error_Report.Range("I" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AP:$AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
          
            'Link ShipToCity to Source Data (Column I)
            Set LinkCell = Error_Report.Range("I" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
           LinkAddress = "'Sage AR Data'!AP" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
          'Lookup ShiptoAddress 1 from Sage AR Data (Column J)
           Error_Report.Range("J" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AM:$AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
                                 
          'Link Column J (Ship To Address 1) to Source Data (Column J)
            Set LinkCell = Error_Report.Range("J" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AM" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
        
          'Lookup Ship To Address 2 from Sage AR Data (Column K)
           Error_Report.Range("K" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!AN:$AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
            
          'Link Column J (Ship To Address 2) to Source Data (Column K)
            Set LinkCell = Error_Report.Range("K" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AN" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
            
        'Lookup Ship To Address 3 from Sage AR Data (Column L)
        Error_Report.Range("L" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!AO:$AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
        'Link Column J (Ship To Address 3) to Source Data (Column L)
            Set LinkCell = Error_Report.Range("L" & i)
            MatchString = "IFError(Match(" & Error_Report.Range("A" & i).Value & ",'Sage AR Data'!" & Lookuplocation & ",0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Sage AR Data'!AO" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
        
        'Lookup Exemption Reason from Exemption Report (Column M)
        
        Error_Report.Range("M" & i).Value = "=xlookup(C" & i & ",'Exemption Report'!B:B,'Exemption Report'!F:F," & Chr(34) & "ERROR" & Chr(34) & ",0)&" & Chr(34) & "-" & Chr(34) & "&" & "xlookup(C" & i & ",'Exemption Report'!B:B,'Exemption Report'!M:M," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
        
        'Link Column M (Exemption Reason) to Source Data (Column M)
            Set LinkCell = Error_Report.Range("M" & i)
            MatchString = "IFError(MATCH(" & Chr(34) & Error_Report.Range("C" & i).Value & Chr(34) & ",'Exemption Report'!B:B,0),1)"
            MatchString = Evaluate(MatchString)
            If MatchString = "1" Then
            Else
            LinkAddress = "'Exemption Report'!F" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            End If
        
        
                   Set LinkCell = Error_Report.Range("N" & i)
            LinkAddress = "'Sage_Gross Compare'!B" & Sheets("Sage_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("O" & i)
            LinkAddress = "'Sage_Gross Compare'!$C$" & Sheets("Sage_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("P" & i)
            LinkAddress = "'Sage_Gross Compare'!$D$" & Sheets("Sage_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("Q" & i)
            LinkAddress = "'Sage_Gross Compare'!$E$" & Sheets("Sage_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("R" & i)
            LinkAddress = "'Sage_Gross Compare'!$F$" & Sheets("Sage_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("S" & i)
            MatchString = "IfError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            LinkAddress = "''Vertex_Gross Compare'!B" & MatchString
            
            LinkAddress = "'Vertex_Gross Compare'!$B$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
             Set LinkCell = Error_Report.Range("U" & i)
            MatchString = "IfError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            LinkAddress = "'Vertex_Gross Compare'!$C$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("W" & i)
             MatchString = "IfError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            LinkAddress = "'Vertex_Gross Compare'!$D$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
           
            Set LinkCell = Error_Report.Range("Y" & i)
            MatchString = "IfError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            LinkAddress = "'Vertex_Gross Compare'!$F$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

            Set LinkCell = Error_Report.Range("AA" & i)
            MatchString = "IfError(Match(" & Error_Report.Range("A" & i).Value & ",'Vertex_Gross Compare'!A:A,0),1)"
            MatchString = Evaluate(MatchString)
            LinkAddress = "'Vertex_Gross Compare'!$G$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
        
        Error_Report.Range("N" & i & ":" & "AB" & i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
        
        If Error_Report.Range("K" & i).Value = 0 Then
        Error_Report.Range("K" & i).Value = ""
        End If
        If Error_Report.Range("L" & i).Value = 0 Then
        Error_Report.Range("L" & i).Value = ""
        End If
        If Error_Report.Range("M" & i).Value = 0 Then
        Error_Report.Range("M" & i).Value = ""
        End If
        
        
If MatchString = "1" Then
Error_Report.Range("AC" & i).Value = "Vertex did not have this transaction within its database indicating that the shiptostate on the sage transaction is incorrect. Check the Sage transaction for address validation erros."
End If
      
      End If
Next i

Error_Report.Columns.AutoFit

End Sub
Sub Step_16_Secondary_8_Loop_And_Link_Section_3()
Error_Report.Range("D" & VertexCompareStart - 1).Value = "Customer Name"
For i = VertexCompareStart To VertexEnd

    Error_Report.Range("B" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$E:$E," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    Error_Report.Range("B" & i).NumberFormat = "mm/dd/yyyy"
            Set LinkCell = Error_Report.Range("A" & i)
            LinkAddress = Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).Address(external:=True)
            LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
    
       Error_Report.Range("C" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$H:$H," & Chr(34) & "ERROR" & Chr(34) & ",0)"
       Error_Report.Range("D" & i).Value = "=xlookup(C" & i & ",'Exemption Report'!B:B" & ",'Exemption Report'!D:D," & Chr(34) & "ERROR" & Chr(34) & ",0)"
       
               If Len(Error_Report.Range("D" & i).Value) > 1 Then
            Set LinkCell = Error_Report.Range("C" & i)
            On Error GoTo Error_1
            MatchString = "IfError(Match(" & Error_Report.Range("A" & i).Value & ",'Exemption Report'!B:B,0),1)"
            MatchString = Evaluate(MatchString)
            LinkAddress = "'Exemption Report'!B" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
                End If
Error_1:
      Error_Report.Range("E" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AQ:$AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
       Error_Report.Range("F" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AR:$AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("G" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$J:$J," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("H" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$CY:$CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("I" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AP:$AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("J" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!$AM:$AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("K" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!AN:$AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("L" & i).Value = "=xlookup(A" & i & ",'Sage AR Data'!" & Lookuplocation & ",'Sage AR Data'!AO:$AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        Error_Report.Range("M" & i).Value = "=xlookup(C" & i & ",'Exemption Report'!B:B" & ",'Exemption Report'!F:F," & Chr(34) & "ERROR" & Chr(34) & ",0)"
        
                   Set LinkCell = Error_Report.Range("N" & i)
            LinkAddress = "'Vertex_Gross Compare'!B" & Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("O" & i)
            LinkAddress = "'Vertex_Gross Compare'!$C$" & Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("P" & i)
            LinkAddress = "'Vertex_Gross Compare'!$D$" & Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("Q" & i)
            LinkAddress = "'Vertex_Gross Compare'!$E$" & Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("R" & i)
            LinkAddress = "'Vertex_Gross Compare'!$F$" & Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("S" & i)
            LinkAddress = "'Vertex_Gross Compare'!$G$" & Sheets("Vertex_Gross Compare").Cells.Find(Error_Report.Range("A" & i)).row
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
             Set LinkCell = Error_Report.Range("T" & i)
             
             MatchString = "Iferror(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
             MatchString = Evaluate(MatchString)
            LinkAddress = "'Sage_Gross Compare'!$B$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Set LinkCell = Error_Report.Range("V" & i)
             
             MatchString = "Iferror(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
             MatchString = Evaluate(MatchString)
            LinkAddress = "'Sage_Gross Compare'!$C$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
             Set LinkCell = Error_Report.Range("X" & i)
             
             MatchString = "Iferror(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
             MatchString = Evaluate(MatchString)
            LinkAddress = "'Sage_Gross Compare'!$D$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

             Set LinkCell = Error_Report.Range("Z" & i)
             MatchString = "Iferror(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
             MatchString = Evaluate(MatchString)
            LinkAddress = "'Sage_Gross Compare'!$E$" & MatchString
            LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
             Set LinkCell = Error_Report.Range("AB" & i)
             MatchString = "Iferror(Match(" & Error_Report.Range("A" & i).Value & ",'Sage_Gross Compare'!A:A,0),1)"
             MatchString = Evaluate(MatchString)
             LinkAddress = "'Sage_Gross Compare'!$F$" & MatchString
             LinkCell.Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
            
            Error_Report.Range("N" & i & ":" & "AB" & i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
            
            
 If MatchString = "1" Then
Error_Report.Range("AC" & i).Value = "Sage did not have this transaction within its database indicating that the shiptostate or country code on the sage transaction is incorrect. Check the Sage transaction for address validation erros."
End If
            
       
Next i
End Sub

Sub Step_16_Main_Setup_Error_Report()

Step_16_Secondary_1_Setup_Section_1

Step_16_Secondary_2_Setup_Section_2

Step_16_Secondary_3_Setup_Section_3

Step_16_Secondary_4_Setup_Sage_Headers

Step_16_Secondary_5_Loop_And_Link_Section_1

Step_16_Secondary_7_Loop_And_Link_Section_2

Step_16_Secondary_8_Loop_And_Link_Section_3



For i = VertexCompareStart To VertexEnd
If Error_Report.Range("U" & i).Value = 0 And Error_Report.Range("W" & i).Value = 0 And Error_Report.Range("Y" & i).Value = 0 And Error_Report.Range("AA" & i).Value = 0 Then
Range("A" & i).EntireRow.Delete
If IsEmpty(Error_Report.Range("U" & i).Value) = False Then

i = i - 1
End If


Else
End If

Next i

'Delete Duplicate Records
For i = VertexCompareStart To VertexEnd
   
    If WorksheetFunction.CountIf(Error_Report.Range("A:A"), Error_Report.Range("A" & i).Value) > 1 Then
        Error_Report.Rows(i & ":" & i).Delete
        i = i - 1

    End If
Next i
Error_Report.Columns.AutoFit
       
    
    Step_16_Secondary_9_ReOrder_Columns_Section_3
 
 Error_Report.Range("AC:AC").WrapText = True
Error_Report.Range("AC:AC").ColumnWidth = 65
End Sub

Sub Step_16_Secondary_9_ReOrder_Columns_Section_3()
 'Re-Order Columns to match other report data items

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Sage Gross Sales", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Sage Gross Sales", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Sage Taxable Sales", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Sage Taxable Sales", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Sage NonTaxable Sales", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Sage NonTaxable Sales", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Sage Freight Amount", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Sage Freight Amount", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Sage Tax Amount", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Sage Tax Amount", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Gross_Amt", HeaderRow, 0)))

'Move the Vertex Gross Amount Column into alignment with the other headers in the report
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Gross Sales Variance", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Gross Sales Variance", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Taxable_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Taxable_Amt", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Taxable Sales Variance", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Taxable Sales Variance", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex Exempt Amount", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex Exempt Amount", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Delete Extra Vertex Column
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex_Tax_Amt", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex_Tax_Amt", HeaderRow, 0)))

VertexCutRange.Delete xlShiftToLeft

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("NonTaxable Sales Variance", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("NonTaxable Sales Variance", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex Freight", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex Freight", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

'Define Cut Range
Set VertexCutRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Freight Variance", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Freight Variance", HeaderRow, 0)))

'Define Paste Range
Set VertexPasteRange = Error_Report.Range(Error_Report.Cells(HeaderRow.row, _
WorksheetFunction.Match("Vertex Tax", HeaderRow, 0)) _
, Error_Report.Cells(VertexEnd, WorksheetFunction.Match("Vertex Tax", HeaderRow, 0)))

'Move Column
VertexCutRange.Cut
VertexPasteRange.Insert xlShiftToRight

End Sub
Sub Step_16_Secondary_4_Setup_Sage_Headers()

'Create Headers to Pull in Transaction Information such as customer no, exemption reason, shipping details, tax schedule etc.
Error_Report.Range("B:M").Insert xlShiftToRight
Error_Report.Range("B" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Invoice Date"
Error_Report.Range("C" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Customer No."
Error_Report.Range("D" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Vertex Customer Name"
Error_Report.Range("E" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Ship To State"
Error_Report.Range("F" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Ship to ZipCode"
Error_Report.Range("G" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Tax Schedule"
Error_Report.Range("H" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Order Manager"
Error_Report.Range("I" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Ship To City"
Error_Report.Range("J" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Ship To Address 1"
Error_Report.Range("K" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Ship To Address 2"
Error_Report.Range("L" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Ship To Address 3"
Error_Report.Range("M" & Error_Report.Cells.Find("Sage Gross Sales", Error_Report.Range("A1")).row).Value = "Exemption Reason"

Error_Report.Range("B" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Invoice Date"
Error_Report.Range("C" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Customer No."
Error_Report.Range("D" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Customer Name"
Error_Report.Range("E" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Ship To State"
Error_Report.Range("F" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Ship to ZipCode"
Error_Report.Range("G" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Tax Schedule"
Error_Report.Range("H" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Order Manager"
Error_Report.Range("I" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Ship To City"
Error_Report.Range("J" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Ship To Address 1"
Error_Report.Range("K" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Ship To Address 2"
Error_Report.Range("L" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Ship To Address 3"
Error_Report.Range("M" & Error_Report.Cells.Find("Sage Gross Mkplc Sales", Error_Report.Range("A1")).row).Value = "Exemption Reason"

Error_Report.Range("B" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Invoice Date"
Error_Report.Range("C" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Customer No."
Error_Report.Range("D" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship To State"
Error_Report.Range("E" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship To State"
Error_Report.Range("F" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship to ZipCode"
Error_Report.Range("G" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Tax Schedule"
Error_Report.Range("H" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Order Manager"
Error_Report.Range("I" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship To City"
Error_Report.Range("J" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship To Address 1"
Error_Report.Range("K" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship To Address 2"
Error_Report.Range("L" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Ship To Address 3"
Error_Report.Range("M" & Error_Report.Cells.Find("Vertex_Gross_Amt", Error_Report.Range("A1")).row).Value = "Exemption Reason"


Error_Report.Columns.AutoFit
End Sub

Sub Step_16_Secondary_3_Setup_Section_3()
'Fill in a description of the data in the section
Error_Report.Range("A" & Counter + 2).Value = "The following transactions are sourced from a camparison against Vertex Data, if an invoice appears here its because it exists in vertex but does not exist in Sage."

'Format the section description and apply formatting
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).Merge
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).WrapText = True
Error_Report.Rows(Counter + 2 & ":" & Counter + 2).RowHeight = 35
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).Cells.BorderAround 1, xlThick
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).Cells.Interior.Color = RGB(175, 177, 193)
'Calculate the next blank row accounting for merged rows( 3 merged rows and add one to get next blank row)
Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 4
'Copy the report headers from the veretex gross compare sheet
Vertex_Gross_Sheet.Range("6:6").Copy
Error_Report.Range("A" & Counter).PasteSpecial xlPasteAll
Error_Report.Range("A" & Counter).Value = "Trimmed Invoice Number"
Error_Report.Range("R" & Counter).Value = "Note"
Error_Report.Range("A" & Counter & ":R" & Counter).WrapText = True
Error_Report.Range("A" & Counter & ":R" & Counter).HorizontalAlignment = xlCenter
Error_Report.Range("A" & Counter & ":R" & Counter).VerticalAlignment = xlCenter
Error_Report.Range("A" & Counter & ":R" & Counter).BorderAround 1, xlThick


Set HeaderRow = Error_Report.Range("A" & Counter & ":R" & Counter)

'Loop through the header rows and format if vertex source, sage source, or calculation
For Each HeaderCell In HeaderRow

If InStr(HeaderCell.Value, "Vertex") > 0 And InStr(HeaderCell.Value, "Variance") = 0 Then

HeaderCell.Font.Color = RGB(4, 60, 116)
HeaderCell.Interior.Color = RGB(195, 213, 66)
HeaderCell.BorderAround 1, xlThin


ElseIf InStr(HeaderCell.Value, "Sage") > 0 And InStr(HeaderCell.Value, "Variance") = 0 Then

HeaderCell.Font.Color = RGB(4, 220, 4)
HeaderCell.Interior.Color = RGB(248, 244, 244)
HeaderCell.BorderAround 1, xlThin

ElseIf InStr(HeaderCell.Value, "Variance") > 0 Then
HeaderCell.Font.Color = RGB(250, 125, 0)
HeaderCell.Interior.Color = RGB(242, 242, 242)
HeaderCell.BorderAround 1, xlThin

Else
HeaderCell.Interior.Color = RGB(248, 244, 244)
HeaderCell.BorderAround 1, xlThin

End If

Next HeaderCell
VertexCompareStart = Error_Report.Cells(Rows.Count, 1).End(xlUp).row + 1

'Loop Through Vertex Gtoww Sheet looking at variances
For X = Vertex_Gross_Sheet.Cells.Find("Vertex_Gross_Amt", After:=Vertex_Gross_Sheet.Range("A1")).row + 1 To Vertex_Pivot_LR

    If Vertex_Gross_Sheet.Range(Split(Cells(1, Vertex_Gross_Sheet.Cells.Find("Gross Sales Variance").Column).Address, "$")(1) & X).Value <> 0 _
    Or Vertex_Gross_Sheet.Range(Split(Cells(1, Vertex_Gross_Sheet.Cells.Find("Taxable Sales Variance").Column).Address, "$")(1) & X).Value <> 0 _
    Or Vertex_Gross_Sheet.Range(Split(Cells(1, Vertex_Gross_Sheet.Cells.Find("NonTaxable Sales Variance").Column).Address, "$")(1) & X).Value <> 0 _
    Or Vertex_Gross_Sheet.Range(Split(Cells(1, Vertex_Gross_Sheet.Cells.Find("Freight Variance").Column).Address, "$")(1) & X).Value <> 0 _
    Or Vertex_Gross_Sheet.Range(Split(Cells(1, Vertex_Gross_Sheet.Cells.Find("Tax Variance").Column).Address, "$")(1) & X).Value <> 0 Then
    Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 4
        Vertex_Gross_Sheet.Range("A" & X & ":P" & X).Copy Destination:=Error_Report.Range("A" & Counter)
        Vertex_Gross_Sheet.Range("A" & X & ":P" & X).Style = "Bad"
         For Each HeaderCell In Error_Report.Range("A" & Counter & ":Q" & Counter)
            HeaderCell.BorderAround 1, xlThin
        Next HeaderCell
        End If
Next X
    

    

    
    
VertexEnd = Error_Report.Cells(Rows.Count, 1).End(xlUp).row

End Sub

Sub Step_16_Secondary_2_Setup_Section_2()
'Write a description of the report section
Error_Report.Range("A" & Counter + 1).Value = "The following transactions are Marketplace Facilitator Transactions that show an error, if transactions appear here then the Sage ZipCode does not match the ShipToState of the Sage invoice, or the transaction is missing the country code, or the transaction hasn't yet been uploaded to vertex."

'Format the section header cells
Error_Report.Range("A" & Counter + 1 & ":Q" & Counter + 1).Merge
Error_Report.Range("A" & Counter + 1 & ":Q" & Counter + 1).WrapText = True
Error_Report.Rows(Counter + 1 & ":" & Counter + 1).RowHeight = 35
Error_Report.Range("A" & Counter + 1 & ":Q" & Counter + 1).Cells.BorderAround 1, xlThick
Error_Report.Range("A" & Counter + 1 & ":Q" & Counter + 1).Cells.Interior.Color = RGB(175, 177, 193)

'Calculate the next blank row
Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 1

'Copy the headers from the marketplace facilitator table
Mktplc.Range("6:6").Copy
Error_Report.Range("A" & Counter + 2).PasteSpecial xlPasteAll
'Add headers for trimmed invoice no and note columns
Error_Report.Range("A" & Counter + 2).Value = "Trimmed Invoice Number"
Error_Report.Range("Q" & Counter + 2).Value = "Note"
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).WrapText = True
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).HorizontalAlignment = xlCenter
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).VerticalAlignment = xlCenter
Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2).BorderAround 1, xlThick

'Set the coordinates of the header row for the marketplace facilitator section
Set HeaderRow = Error_Report.Range("A" & Counter + 2 & ":Q" & Counter + 2)

'Loop through each of the header cells and format the cells based on their origin (Verex, Sage, Calculation)
For Each HeaderCell In HeaderRow

If InStr(HeaderCell.Value, "Vertex") > 0 And InStr(HeaderCell.Value, "Variance") = 0 Then

HeaderCell.Font.Color = RGB(4, 60, 116)
HeaderCell.Interior.Color = RGB(195, 213, 66)
HeaderCell.BorderAround 1, xlThin


ElseIf InStr(HeaderCell.Value, "Sage") > 0 And InStr(HeaderCell.Value, "Variance") = 0 Then

HeaderCell.Font.Color = RGB(4, 220, 4)
HeaderCell.Interior.Color = RGB(248, 244, 244)
HeaderCell.BorderAround 1, xlThin

ElseIf InStr(HeaderCell.Value, "Variance") > 0 Then
HeaderCell.Font.Color = RGB(250, 125, 0)
HeaderCell.Interior.Color = RGB(242, 242, 242)
HeaderCell.BorderAround 1, xlThin

Else
HeaderCell.Interior.Color = RGB(248, 244, 244)
HeaderCell.BorderAround 1, xlThin

End If
Next HeaderCell

'Determine the next blank row to start the copying of data
MarketplaceStart = Cells(Rows.Count, 1).End(xlUp).row + 1

For i = Mktplc.Cells.Find("Sage Gross Mkplc Sales", After:=Mktplc.Range("A1")).row + 1 To Mktplc_LR - 1
    
        If Mktplc.Range("H" & i).Value <> 0 Or Mktplc.Range("J" & i).Value <> 0 Or Mktplc.Range("L" & i).Value <> 0 Or Mktplc.Range("N" & i).Value <> 0 Or Mktplc.Range("P" & i).Value <> 0 Then
        Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 3
        Mktplc.Range("A" & i & ":P" & i).Copy Destination:=Error_Report.Range("A" & Counter)
        Mktplc.Range("A" & i & ":P" & i).Style = "Bad"
         For Each HeaderCell In Error_Report.Range("A" & Counter & ":Q" & Counter)
            HeaderCell.BorderAround 1, xlThin
        Next HeaderCell
        

End If
Next i

'Calculate the coordinates of the end of the marketplace facilitated sales section
MarketplaceEnd = Error_Report.Cells(Rows.Count, 1).End(xlUp).row + 1
 
 
 'Create the Vertex vs Sage Section of the report
 
 'Set Counter equal to the next blank row to start the next error report section
Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 2
End Sub

Sub Step_16_Secondary_1_Setup_Section_1()
SageAR_LastRow = WorksheetFunction.CountA(Sheets("Sage AR Data").Range("A:A"))

Sheets("Sage AR Data").Cells.Find ("InvoiceDate")

'Activate The Error Report Sheet
Error_Report.Activate

'Define The first row
SageCompareStart = Error_Report.Range("A1").row


'Describe the Report Section
Error_Report.Range("A1").Value = "Transactions with Variances Below are sourced with Sage Transactions as the Base. Items here will not include any Invoice Numbers Included in Vertex but not In the sage download from the ERP. " & _
"Causes of such variances could include but are not limited to transactions having incorrect ShipToStates or Country Codes in Sage, or a Zipcode that does not match the ship to state in Sage as transactions run through" & _
"Vertex overide to the correct ShiptoState as validated by ZipCode. Variances in the amount of tax could be due to the timing of the receipt of exemption certificates, and/ or Sales" & _
"Tax refunds processed in the period (Typically will Have Tax Schedule Default) or the start date of a certificate being subsequent to the order date of the transaction. If the tax Schedule of an Invoice is " & Chr(34) & "Default" & Chr(34) & " it is a correction that needs uploaded to Vertex. If the tax amount is different, it is likely the result of a tax rate change that was not calculated at the time of invoicing due to a known issue with Sage100 Connector and Starship where the transaction calls the order date for tax calculation instead of invoice date."

'Format the Sage Comparison Report Section
Error_Report.Range("A1:Q2").Merge
Error_Report.Range("A1:Q2").HorizontalAlignment = xlLeft
Error_Report.Range("A1:Q2").VerticalAlignment = xlCenter
Error_Report.Range("A1:Q2").VerticalAlignment = xlCenter
Error_Report.Range("A1:Q2").WrapText = True
Error_Report.Range("A1:Q2").Cells.BorderAround 1, xlThick
Error_Report.Range("A1:Q2").Cells.Interior.Color = RGB(175, 177, 193)
Rows("1:1").RowHeight = 60

'Determine the Last Row of the Error Report Section Plus 2 to account for the merged cell above
Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 2

'Copy The headers from the Sage Pivot Table Comparison
Sage_Pivot.Range("6:6").Copy

'Paste the Headers into the Error Report
Error_Report.Range("A" & Counter).PasteSpecial xlPasteAll
Error_Report.Range("A3").Value = "Trimmed Invoice Number"
Error_Report.Range("Q" & Counter).Value = "Note"
Error_Report.Range("A:Q").ColumnWidth = 13
Error_Report.Range("A" & Counter & ":Q" & Counter).WrapText = True
Error_Report.Range("A" & Counter & ":Q" & Counter).HorizontalAlignment = xlCenter
Error_Report.Range("A" & Counter & ":Q" & Counter).VerticalAlignment = xlCenter
Error_Report.Range("A" & Counter & ":Q" & Counter).BorderAround 1, xlThick


'Define the Header Row for Formatting
Set HeaderRow = Error_Report.Range("A" & Counter & ":Q" & Counter)

'Loop Through Each Cell in the header row

For Each HeaderCell In HeaderRow

    'Format Headers for Vertex Related Sourced Data
    If InStr(HeaderCell.Value, "Vertex") > 0 And InStr(HeaderCell.Value, "Variance") = 0 Then
    HeaderCell.Font.Color = RGB(4, 60, 116)
    HeaderCell.Interior.Color = RGB(195, 213, 66)
    HeaderCell.BorderAround 1, xlThin
    
    'Format Headers for Sage Related Sourced Data
    ElseIf InStr(HeaderCell.Value, "Sage") > 0 And InStr(HeaderCell.Value, "Variance") = 0 Then
    
    HeaderCell.Font.Color = RGB(4, 220, 4)
    HeaderCell.Interior.Color = RGB(248, 244, 244)
    HeaderCell.BorderAround 1, xlThin
    
    'Format Headers for Variance Cells
    ElseIf InStr(HeaderCell.Value, "Variance") > 0 Then
    HeaderCell.Font.Color = RGB(250, 125, 0)
    HeaderCell.Interior.Color = RGB(242, 242, 242)
    HeaderCell.BorderAround 1, xlThin
    
    'Format All Other Headers
    Else
    HeaderCell.Interior.Color = RGB(248, 244, 244)
    HeaderCell.BorderAround 1, xlThin
    
    End If
    
Next HeaderCell

'Once the first set of headers has been formatted, loop through each cell on the Sage_Pivot Tab
For i = Sage_Pivot.Cells.Find("Sage Gross Sales", Sage_Pivot.Range("A1")).row + 1 To Sage_Pivot.Cells.Find("Grand Total").row


'Check Sage against Vertex Excluding Marketplace Sales

'Check to see if There is a sales variance on a gross sales basis, taxable sales variance, non-taxable sales variance, a freight variance, or a tax variance

    If Round(Sage_Pivot.Range("H" & i).Value, 2) <> 0 _
        Or Round(Sage_Pivot.Range("J" & i).Value, 2) <> 0 _
            Or Round(Sage_Pivot.Range("L" & i).Value, 2) <> 0 _
                Or Round(Sage_Pivot.Range("N" & i).Value, 2) <> 0 _
                    Or Round(Sage_Pivot.Range("P" & i).Value, 2) <> 0 Then
                    
'If there is a variance then
  'Determine the last row written to in the error report range:
        Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 1
        Sage_Pivot.Range("A" & i & ":P" & i).Style = "Bad"
        Sage_Pivot.Range("A" & i & ":P" & i).Copy
        Error_Report.Range("A" & Counter + 1).PasteSpecial xlPasteValues
        For Each HeaderCell In Error_Report.Range("A" & Counter + 1 & ":Q" & Counter + 1)
            HeaderCell.BorderAround 1, xlThin
        Next HeaderCell
    End If

Next i

'Once The Sage Versus Vertex Exclusive of Marketplace Transactions has been looped through establish what the end row of the partially filled out report is as a reference to start the next section

SageEnd = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 1

'Determine the last row in the marketplace sales pivot table
Mktplc_LR = Mktplc.Cells.Find("Grand Total", Mktplc.Range("A1")).row

'Set the begining of the next report section
Counter = SageEnd + 1
End Sub

Sub T_FN_Compare_Sage_To_Vertex_Mktplcfac()

Application.Calculation = xlCalculationManual
Mktplc.Range("F" & Mktplc_HR + 1).Value = "Taxable Freight"
Mktplc.Range("G" & Mktplc_HR + 1).Value = "Non-Taxable Freight"
Mktplc.Range("H" & Mktplc_HR + 1).Value = "Vertex Taxable Sales Excluding Freight"
Mktplc.Range("I" & Mktplc_HR + 1).Value = "Vertex Non-Taxable Sales Excluding Freight"
Mktplc.Range("J" & Mktplc_HR + 1).Value = "Taxable Sales Variance"
Mktplc.Range("K" & Mktplc_HR + 1).Value = "Non-Taxable Sales Variance"
Mktplc.Range("L" & Mktplc_HR + 1).Value = "Vertex Tax"
Mktplc.Range("M" & Mktplc_HR + 1).Value = "Vertex Tax Variance"

For X = 7 To Mktplc_LR
Counter = WorksheetFunction.CountA(Error_Report.Range("A:A")) + 1

MktplcLookupFreight = "ROUND(IFERROR(VLOOKUP(A" & X & ",MKTPLCFAC!$A$6:$E$" & Mktplc_LR & ",4,0),0),2)"
MktplcLookupTax = "ROUND(IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$I$6:$J$" & Vertex_Tax_Last_Row & ",2,0),0),2)"
FreightLookupTax = "ROUND(IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$N$6:$R$" & Vertex_Freight_Sales_Last_Row & ",3,0),0),2)"
FreightLookupNonTax = "ROUND(IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$N$6:$R$" & Vertex_Freight_Sales_Last_Row & ",4,0),0),2)"
FreightLookupExemptTax = "ROUND(IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$N$6:$R$" & Vertex_Freight_Sales_Last_Row & ",5,0),0),2)"
VertexTaxSales = "ROUND(IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$A$6:$E$" & Vertex_Sales_Last_Row & ",3,0),0),2)"
VertexNoNTaxSales = "ROUND(IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$A$6:$E$" & Vertex_Sales_Last_Row & ",4,0),0)+IFERROR(VLOOKUP(A" & X & ",'Vertex Pivot'!$A$6:$E$" & Vertex_Sales_Last_Row & ",5,0),0),2)"

Mktplc.Range("F" & X).Value = "=" & FreightLookupTax
Mktplc.Range("G" & X).Value = "=" & FreightLookupNonTax & "+" & FreightLookupExemptTax & "+" & MktplcLookupFreight
Mktplc.Range("H" & X).Value = "=" & VertexTaxSales & "-" & "F" & X
On Error Resume Next
Mktplc.Range("I" & X).Value = WorksheetFunction.VLookup(Mktplc.Range("A" & X).Value, Vertex_Nontax_Exclude_Freight.TableRange2, 4, 0)
Mktplc.Range("J" & X).Formula = "=ROUND(H" & X & "-" & "B" & X & ",2)"
Mktplc.Range("K" & X).Formula = "=ROUND(I" & X & "-" & "C" & X & ",2)"
Mktplc.Range("L" & X).Value = "=" & MktplcLookupTax
Mktplc.Range("M" & X).Value = "=ROUND(E" & X & "-" & "L" & X & ",2)"
Mktplc.Columns.AutoFit

Next X

Mktplc.Range("F" & Mktplc_LR + 1).Value = "=ROUND(SUM(F" & Mktplc_HR & ":F" & Mktplc_LR & "),2)"
Mktplc.Range("G" & Mktplc_LR + 1).Value = "=ROUND(SUM(G" & Mktplc_HR & ":G" & Mktplc_LR & "),2)"
Mktplc.Range("H" & Mktplc_LR + 1).Value = "=ROUND(SUM(H" & Mktplc_HR & ":H" & Mktplc_LR & "),2)"
Mktplc.Range("I" & Mktplc_LR + 1).Value = "=ROUND(SUM(I" & Mktplc_HR & ":I" & Mktplc_LR & "),2)"
Mktplc.Range("J" & Mktplc_LR + 1).Value = "=ROUND(SUM(J" & Mktplc_HR & ":J" & Mktplc_LR & "),2)"
Mktplc.Range("K" & Mktplc_LR + 1).Value = "=ROUND(SUM(K" & Mktplc_HR & ":K" & Mktplc_LR & "),2)"

Application.Calculation = xlAutomatic

End Sub
Sub Step_17_Main_Setup_Summary_Sheet()
Dim LinkRow As Long
Dim LinkAddress As String
Dim LinkCell As Range



'Load the "Summary" Worksheet
Summary.Activate

'Fill in Report Header Information
Summary.Range("A1").Value = "Fotronic Corporation"
Summary.Range("A2").Value = "Comparison of Tax Return, Vertex, and Sage"
Summary.Range("A3").Value = ReconState & " for Period Ended " & Format(EndDate, "yyyy-mm-dd")
Summary.Range("C1").Value = "Return Data"
Summary.Range("E1").Value = "Vertex Data"
Summary.Range("G1").Value = "Sage Data"
Summary.Columns.AutoFit

'Bold the Titles and Header Information
Summary.Range("1:3").Font.Bold = True
Summary.Range("1:3").Font.Size = 14
Summary.Range("1:3").WrapText = True

'Autofit the columns to contents
Summary.Columns.AutoFit

'data lables
Summary.Range("A4").Value = "Gross Sales:"
Summary.Range("A5").Value = "Taxable Sales:"
Summary.Range("A6").Value = "NonTaxable or Excluded Sales:"
Summary.Range("A7").Value = "Sales Tax:"

'Determine whether Sales Tax should be included in the Gross Totals or Not
If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then

Summary.Range("E4").Value = WorksheetFunction.VLookup("Grand Total", Vertex_Gross_Pivot.TableRange2, 2, 0) + WorksheetFunction.VLookup("Grand Total", Vertex_Tax.TableRange2, 2, 0)

Summary.Range("E5").Value = WorksheetFunction.VLookup("Grand Total", Vertex_Gross_Pivot.TableRange2, 3, 0)
Summary.Range("E6").Value = WorksheetFunction.VLookup("Grand Total", Vertex_Gross_Pivot.TableRange2, 2, 0) - WorksheetFunction.VLookup("Grand Total", Vertex_Gross_Pivot.TableRange2, 3, 0)
Summary.Range("E7").Value = WorksheetFunction.VLookup("Grand Total", Vertex_Gross_Pivot.TableRange2, 6, 0)

Summary.Range("G4").Value = WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 2, 0) + WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 3, 0) + WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 4, 0) + WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 5, 0)
Summary.Range("G5").Value = WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 2, 0) + WorksheetFunction.VLookup("Grand Total", Vertex_Freight_Sales.TableRange2, 3, 0)
Summary.Range("G6").Value = WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 3, 0) + WorksheetFunction.VLookup("Grand Total", Vertex_Freight_Sales.TableRange2, 2, 0) - WorksheetFunction.VLookup("Grand Total", Vertex_Freight_Sales.TableRange2, 3, 0)
Summary.Range("G7").Value = WorksheetFunction.VLookup("Grand Total", Sage_Gross_Pivot.TableRange2, 5, 0)

Summary.Range("C4:G7").Style = "Comma"
Summary.Range("I1").Value = "Variance - Sage Data to Vertex Data"
Summary.Range("I4").Value = "=E4-G4"
Summary.Range("I5").Value = "=E5-G5"
Summary.Range("I6").Value = "=E6-G6"
Summary.Range("I7").Value = "=E7-G7"

Else
'If Gross Amount does not include sales tax then...

''Populate Gross Vertex Sales
Summary.Range("E4").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Vertex_GrossComp_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_GrossComp_Gross.Address(external:=True) & ",0,0)"


' Link the Vertex Gross Sales to its source data
Summary.Range("E4").Select
LinkRow = Vertex_Gross_Sheet.Cells.Find("Grand Total", After:=Vertex_Gross_Sheet.Range("A1")).row
Set LinkCell = Selection
LinkAddress = Vertex_Gross_Sheet.Range("B" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

'Format the data into an accounting format
Summary.Range("E4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"

'Populate Vertex Taxable Sales
Summary.Range("E5").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Vertex_GrossComp_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_GrossComp_Taxable.Address(external:=True) & ",0,0)"

'Link Vertex Taxable Data to Source
Summary.Range("E5").Select
LinkRow = Vertex_Gross_Sheet.Cells.Find("Grand Total", After:=Vertex_Gross_Sheet.Range("A1")).row
Set LinkCell = Selection
LinkAddress = Vertex_Gross_Sheet.Range("C" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

'Format the data to accounting format
Summary.Range("E5").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"

'Populate Vertex Exempt Sales
Summary.Range("E6").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Vertex_GrossComp_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_GrossComp_Exempt.Address(external:=True) & ",0,0)"

'Link Vertex Exempt Sales to Source Data
Summary.Range("E6").Select
LinkRow = Vertex_Gross_Sheet.Cells.Find("Grand Total", After:=Vertex_Gross_Sheet.Range("A1")).row
Set LinkCell = Selection
LinkAddress = Vertex_Gross_Sheet.Range("D" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

'Format the data to accounting format
Summary.Range("E6").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"


'Populate Vertex Sales Tax
Summary.Range("E7").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Vertex_Tax_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Tax_Tax.Address(external:=True) & ",0,0)"

'Link Vertex Sales Tax to source data
Summary.Range("E7").Select
LinkRow = WorksheetFunction.CountA(Vertex_Pivot.Range("I:I")) + 3
Set LinkCell = Selection
LinkAddress = Vertex_Pivot.Range("J" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
Summary.Range("E7").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"


'Populate Sage Gross Sales
Summary.Range("G4").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Gross.Address(external:=True) & ",0,0)"

'Link Sage Gross Sales to source data
Summary.Range("G4").Select
LinkRow = WorksheetFunction.CountA(Sheets("Sage_Gross Compare").Range("A:A")) + 4
Set LinkCell = Selection
LinkAddress = Sheets("Sage_Gross Compare").Range("B" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

' format Sage Gross Sales as Accounting
Summary.Range("G4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"


' Create a second view of Sage Sales starting with Taxable
Summary.Range("J3").Value = "Sage Taxable Sales:"
Summary.Range("J3").Style = "Normal"

'Populate the Sage Taxable Sales Value
Summary.Range("K3").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Taxable.Address(external:=True) & ",0,0)"
Summary.Range("K3").Select
LinkRow = WorksheetFunction.CountA(Sheets("Sage_Gross Compare").Range("A:A")) + 4
Set LinkCell = Selection
LinkAddress = Sheets("Sage_Gross Compare").Range("C" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

'Set the Number Format for the Sage Taxable Sales Amount
Summary.Range("K3").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"

'Populate the Vertex Taxable Freight
Summary.Range("J4").Value = "Vertex Taxable Freight:"
Summary.Range("K4").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Taxable.Address(external:=True) & ",0,0)"
Summary.Range("K4").Select

'Link the summary back to the data source for the Vertex data
LinkRow = WorksheetFunction.CountA(Sheets("Vertex Pivot").Range("P:P")) + 6
Set LinkCell = Selection
LinkAddress = Sheets("Vertex Pivot").Range("P" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

'Format the number as accounting/currency
Summary.Range("K4").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"


Summary.Range("J5").Value = "Total Sage Taxable Sales:"
Summary.Range("K5").Style = "Normal"
Summary.Range("K5").Value = "=K3+K4"
Summary.Range("K5").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"

Summary.Range("J7").Value = "Sage Non-Taxable Sales"
Summary.Range("K7").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Nontaxable.Address(external:=True) & ",0,0)"
Summary.Range("K7").Select
LinkRow = WorksheetFunction.CountA(Sheets("Sage_Gross Compare").Range("A:A")) + 4
Set LinkCell = Selection
LinkAddress = Sheets("Sage_Gross Compare").Range("D" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
Summary.Range("K7").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"


Summary.Range("J8").Value = "Vertex Non-Taxable Freight"
Summary.Range("K8").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Exempt.Address(external:=True) & ",0,0)"
Summary.Range("K8").Select
LinkRow = WorksheetFunction.CountA(Sheets("Vertex Pivot").Range("Q:Q")) + 6
Set LinkCell = Selection
LinkAddress = Sheets("Vertex Pivot").Range("Q" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
Summary.Range("K8").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"

Summary.Range("J9").Value = "Total Sage Non-Taxable"
Summary.Range("K9").Value = "=K7+K8"
Summary.Range("K9").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"


Summary.Range("G5").Value = "=K5"
Summary.Range("G5").Select
LinkRow = 5
Set LinkCell = Selection
LinkAddress = Sheets("Summary").Range("K" & LinkRow).Address(external:=True)
LinkAddress = "K5"
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

Summary.Range("G6").Value = "=K9"
Summary.Range("G6").Select
LinkRow = 5
Set LinkCell = Selection
LinkAddress = Sheets("Summary").Range("K" & LinkRow).Address(external:=True)
LinkAddress = "K9"
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)

Summary.Range("G7").Value = "=Xlookup(" & Chr(34) & "Grand Total" & Chr(34) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_SageTax.Address(external:=True) & ",0,0)"
Summary.Range("G7").Select
LinkRow = WorksheetFunction.CountA(Sheets("Sage_Gross Compare").Range("A:A")) + 4
Set LinkCell = Selection
LinkAddress = Sheets("Sage_Gross Compare").Range("F" & LinkRow).Address(external:=True)
LinkAddress = "'" & Right(LinkAddress, Len(LinkAddress) - InStr(LinkAddress, "]"))
ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkAddress, TextToDisplay:=CStr(Selection.Value)
Summary.Range("K3").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"



Summary.Range("I1").Value = "Variance - Sage Data to Vertex Data"
Summary.Range("I4").Value = "=E4-G4"
Summary.Range("I5").Value = "=E5-G5"
Summary.Range("I6").Value = "=E6-G6"
Summary.Range("I7").Value = "=E7-G7"

 End If
If User_Interface.CheckBox1.Value = True Then
Summary.Range("A10").Value = "Credits Held"
Summary.Range("A11").Value = "Credits Applied"
Summary.Range("C10").Value = "=-SUMIF('Credit Report'!D:D," & Chr(34) & ReconState & Chr(34) & ",'Credit Report'!H:H)"
Summary.Range("C11").Value = "=-SUMIF('Credit Report'!D:D," & Chr(34) & ReconState & Chr(34) & ",'Credit Report'!I:I)"
Summary.Range("E10").Value = "=SUMIF('Credit Report'!D:D," & Chr(34) & ReconState & Chr(34) & ",'Credit Report'!H:H)"
Summary.Range("E11").Value = "=SUMIF('Credit Report'!D:D," & Chr(34) & ReconState & Chr(34) & ",'Credit Report'!I:I)"
Summary.Range("G10").Value = "=SUMIF('Credit Report'!D:D," & Chr(34) & ReconState & Chr(34) & ",'Credit Report'!H:H)"
Summary.Range("G11").Value = "=SUMIF('Credit Report'!D:D," & Chr(34) & ReconState & Chr(34) & ",'Credit Report'!I:I)"
Summary.Range("A13").Value = "Adjusted Tax"
Summary.Range("C13").Value = "=SUM(C7,C10,C11)"
Summary.Range("E13").Value = "=SUM(E7,E10,E11)"
Summary.Range("G13").Value = "=SUM(G7,G10,G11)"
Summary.Range("A14").Value = "Return Credit Adjusted Tax Variance To Reconciled Tax"
Summary.Range("C14").Value = "=E13-C7"
Else
End If

Summary.Range("A15").Value = "Vendor Credit/Discount"

Dim Sage_AR_Data_NonTaxable As Range
Dim Sage_AR_Data_Taxable As Range
Dim Sage_AR_Data_Freight As Range
Dim Sage_AR_Data_Tax As Range


'Sage_AR_Data_NonTaxable = Sage.Range(Sage.Cells(1, WorksheetFunction.Match("TaxableSalesAmt", Sage.Range("1:1"), 0)), _
'Sage.Cells(WorksheetFunction.CountA(Sage.Range("A:A")), WorksheetFunction.Match("NonTaxableSalesAmt", Sage.Range("1:1"), 0)))

'Sage_AR_Data_Taxable = Sage.Range(Sage.Cells(1, WorksheetFunction.Match("TaxableSalesAmt", Sage.Range("1:1"), 0)), _
'Sage.Cells(WorksheetFunction.CountA(Sage.Range("A:A")), WorksheetFunction.Match("NonTaxableSalesAmt", Sage.Range("1:1"), 0)))

'Sage_AR_Data_Freight = Sage.Range(Sage.Cells(1, WorksheetFunction.Match("FreightAmt", Sage.Range("1:1"), 0)), _
Sage.Cells(WorksheetFunction.CountA(Sage.Range("A:A")), WorksheetFunction.Match("FreightAmt", Sage.Range("1:1"), 0)))

'Sage_AR_Data_Tax = Sage.Range(Sage.Cells(1, WorksheetFunction.Match("SalesTaxAmt", Sage.Range("1:1"), 0)), _
Sage.Cells(WorksheetFunction.CountA(Sage.Range("A:A")), WorksheetFunction.Match("SalesTaxAmt", Sage.Range("1:1"), 0)))

Summary.Range("A20").Value = "Non-Taxable Sales Excluding Mktplcfac"
Summary.Range("A20").Font.Bold = True
Summary.Range("E20").Value = "=E6-'MKTPLCFAC'!C" & Mktplc.Cells.Find("Grand Total", Mktplc.Range("A1")).row & "-'MKTPLCFAC'!D" & Mktplc.Cells.Find("Grand Total", Mktplc.Range("A1")).row
Summary.Range("A21").Value = "eBay Sales"
Summary.Range("E21").Value = "=SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "EBAY001" & Chr(34) & "),'Sage AR Data'!BV:BV)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "EBAY001" & Chr(34) & "),'Sage AR Data'!BW:BW)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "EBAY001" & Chr(34) & "),'Sage AR Data'!BX:BX)"
Summary.Range("A22").Value = "Amazon Sales"
Summary.Range("E22").Value = "=SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "AMAZ002" & Chr(34) & "),'Sage AR Data'!BV:BV)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "AMAZ002" & Chr(34) & "),'Sage AR Data'!BW:BW)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "AMAZ002" & Chr(34) & "),'Sage AR Data'!BX:BX)"
Summary.Range("A23").Value = "Google Sales"
Summary.Range("E23").Value = "=SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "GOOG007" & Chr(34) & "),'Sage AR Data'!BV:BV)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "GOOG007" & Chr(34) & "),'Sage AR Data'!BW:BW)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "GOOG007" & Chr(34) & "),'Sage AR Data'!BX:BX)"
Summary.Range("A24").Value = "Walmart Sales"
Summary.Range("E24").Value = "=SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "WALM020" & Chr(34) & "),'Sage AR Data'!BV:BV)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "WALM020" & Chr(34) & "),'Sage AR Data'!BW:BW)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "WALM020" & Chr(34) & "),'Sage AR Data'!BX:BX)"
Summary.Range("A25").Value = "Sears Sales"
Summary.Range("E25").Value = "=SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "SEAR046" & Chr(34) & "),'Sage AR Data'!BV:BV)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "SEAR046" & Chr(34) & "),'Sage AR Data'!BW:BW)+SUMPRODUCT(--('Sage AR Data'!H:H=" & Chr(34) & "SEAR046" & Chr(34) & "),'Sage AR Data'!BX:BX)"
Summary.Range("A26").Value = "Total Marketplace Facilitator Sales"
Summary.Range("A26").Font.Bold = True
Summary.Range("E26").Value = "=SUM(E21:E25)"
Summary.Range("A29").Value = "Deduction Categorical Summary"
Summary.Range("A29").Font.Bold = True

Dim FreightCheckLoop As Range
Dim ValCheck As Double
Dim Vtex_Nt As Worksheet

Set Vtex_Nt = Taxbook.Sheets("Vertex_Nontaxable")
Vtex_Nt.Name = "Vertex_Nontaxable"

Dim DeductionTypes As Range
Dim Loop_DeductionType As Range
Dim DeductionTypeDict As Object
Dim Key As Variant
Dim LastRow As Long
Dim Nth_ded As Long


With Vtex_Nt
LastRow = Vtex_Nt.Cells.Find("Grand Total", Vtex_Nt.Range("A1"), xlValues).row

Set DeductionTypes = Vtex_Nt.Range("N" & Vtex_Nt.Cells.Find("Trimmed Invoice No", After:=Vtex_Nt.Range("A1")).row + 1 & ":N" & LastRow)
End With
Set DeductionTypeDict = CreateObject("Scripting.Dictionary")


For Each Loop_DeductionType In DeductionTypes
    If Not DeductionTypeDict.Exists(Loop_DeductionType.Value) Then
        DeductionTypeDict.Add Loop_DeductionType.Value, Loop_DeductionType.Value
    End If
Next Loop_DeductionType

Nth_ded = Summary.Cells.Find("Deduction Categorical Summary", After:=Summary.Range("A1")).row
Summary.Range("E" & Nth_ded).Value = "Amount Excluding Freight"
Summary.Range("E" & Nth_ded).WrapText = True
Summary.Range("E" & Nth_ded).Interior.Color = RGB(180, 198, 231)
Summary.Range("G" & Nth_ded).Value = "Freight"
Summary.Range("G" & Nth_ded).WrapText = True
Summary.Range("G" & Nth_ded).Interior.Color = RGB(180, 198, 231)
Summary.Range("I" & Nth_ded).Value = "Total"
Summary.Range("I" & Nth_ded).WrapText = True
Summary.Range("I" & Nth_ded).Interior.Color = RGB(180, 198, 231)

With Summary.Range("A" & Nth_ded & ":I" & Nth_ded)
.Interior.Color = RGB(180, 198, 231)
.Font.Bold = True

End With

For Each Key In DeductionTypeDict.Keys
Nth_ded = Nth_ded + 1
    Summary.Range("A" & Nth_ded).Value = DeductionTypeDict(Key)
Next Key

For i = Summary.Cells.Find("Deduction Categorical Summary", After:=Summary.Range("A1")).row + 1 To Nth_ded - 1
Summary.Range("E" & i).Value = "=Sumproduct(--(" & Chr(34) & Summary.Range("A" & i).Value & Chr(34) & "='Vertex_Nontaxable'!N:N" & "),'Vertex_Nontaxable'!O:O" & ")"
Summary.Range("G" & i).Value = "=Sumproduct(--(" & Chr(34) & Summary.Range("A" & i).Value & Chr(34) & "='Vertex_Nontaxable'!N:N" & "),'Vertex_Nontaxable'!D:D" & ")" & "-" & "Sumproduct(--(" & Chr(34) & Summary.Range("A" & i).Value & Chr(34) & "='Vertex_Nontaxable'!N:N" & "),'Vertex_Nontaxable'!O:O" & ")"
Summary.Range("I" & i).Value = "=E" & i & "+G" & i
Next i

Summary.Range("A" & Summary.Cells(Rows.Count, 1).End(xlUp).row + 2).Value = "Total Vertex Direct Pay Permits Applied"

 Summary.Range("E" & Summary.Cells(Rows.Count, 1).End(xlUp).row).Value = "=SUMPRODUCT(--((Vertex!AZ:AZ=" & Chr(34) & "STATE" & Chr(34) & ")" & Chr(42) & "(Vertex!BY:BY=" & Chr(34) & "Direct Pay Permit Default" & Chr(34) & ")),Vertex!AU:AU)"
Taxbook.Sheets("Zipcodes").Visible = False





End Sub

Sub Q_FN_Setup_Mktplc_Formulas()


Mktplc_HR = MktplcPiv.TableRange1.row
Mktplc_LR = MktplcPiv.DataBodyRange.Rows.Count + Mktplc_HR
Mktplc_LC = WorksheetFunction.CountA(Mktplc.Range(Mktplc_HR & ":" & Mktplc_HR))

Mktplc_Taxable = "MKTPLCFAC!" & Mktplc.Range("B" & Mktplc_LR + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)
Mktplc_NonTaxable = "MKTPLCFAC!" & Mktplc.Range("C" & Mktplc_LR + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)
Mktplc_Freight = "MKTPLCFAC!" & Mktplc.Range("D" & Mktplc_LR + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)
Mktplc_Tax = "MKTPLCFAC!" & Mktplc.Range("E" & Mktplc_LR + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)



End Sub
Sub FN_Create_Error_Report()

Error_Report.Range("A1").Value = "Fotronic Corporation"
Error_Report.Range("A2").Value = EndDate
Error_Report.Range("A3").Value = "Transactions From Sage Where the Vertex Record Disagrees"
Sage_Pivot.Range(Sage_Sales_First_Row & ":" & Sage_Sales_First_Row).Copy
Error_Report.Range("A" & WorksheetFunction.CountA(Error_Report.Range("A:A") + 1)).PasteSpecial xlPasteValuesAndNumberFormats


End Sub
Sub FN_Add_Filters_Sage_Vtx()
Dim Vertex As Worksheet
Dim Sage As Worksheet
Dim Error_Report As Worksheet

Set Vertex = ActiveWorkbook.Sheets("Vertex")
Set Sage = ActiveWorkbook.Sheets("Summary")
Set Error_Report = ActiveWorkbook.Sheets("Error Report")

Dim Error_Sheet_LR As Long

Error_Sheet_LR = Error_Report.Cells.Find("Row Labels", Error_Report.Range("B1")).row - 1

For i = 2 To Error_Sheet_LR

Next i


End Sub
Sub Step_01_Main_Filter_And_Extract_Vertex_data()


'This Subroutine does the following:
'1.) Creates a Sheet within the Vertex Transactional Report Datafile called Settings
'2.) Coppies the Situs Main Division Column Into the created sheet
'3.) Copies the Destination Main Division Column into the created sheet
'4.) Copies the Destination Countrry into the created sheet
'5.) Formats the country code to be USA, and deletes any rows where the country is not USA or US
'6.) Copies the posting dates and removed duplicates on all data sets. This should leave you with a setting sheet that lists the range of dates in the vertex data, their destination state, and desination country.
'These lists are used when activating the Userform at the end that will use these lists to let a user configure the parameters of the reconciliation

'' Initialize variables and settings
Dim Function_Last_Row As Long
Dim i As Long
Dim Config_LastRow As Long
Dim Config_Date_LastRow As Long

Application.DisplayAlerts = False

Set Vertex = ActiveSheet
Set SourceData = ActiveWorkbook

DropboxDir = "C:\Users\christopher.bartus\Fotronic Dropbox\Christopher Bartus\"

' Create a new worksheet called "Settings" and clear any existing ones
On Error Resume Next
Vertex.ShowAllData
Set Config = Sheets.Add
On Error Resume Next
ActiveWorkbook.Sheets("Settings").Delete
Config.Name = "Settings"

' Rename the Vertex Data sheet
Vertex.Name = Format(WorksheetFunction.Min(Vertex.Range("BR:BR")), "YYYY-MM") & " - " & Format(WorksheetFunction.Max(Vertex.Range("BR:BR")), "YYYY-MM") & " Vertex Data"

' Copy Situs Main Division column to the Settings sheet
Copy_Column_To_Settings_Sheet Vertex, Config, "Situs Main Division", "A1", Function_Last_Row

' Copy Destination Main Division column to the Settings sheet
Copy_Column_To_Settings_Sheet Vertex, Config, "Destination Main Division", "B1", Function_Last_Row

' Copy Destination Country column to the Settings sheet
Copy_Column_To_Settings_Sheet Vertex, Config, "Situs Country Code", "C1", Function_Last_Row

' Remove duplicates and sort
Config.Range("A:C").RemoveDuplicates Columns:=Array(1), header:=xlNo
Config.Range("A:C").Sort key1:=Config.Range("A1"), header:=xlYes

' Filter rows with "US" or "USA" in the Destination Country column
Config_LastRow = WorksheetFunction.CountA(Config.Range("A:A"))

For i = 2 To Config_LastRow
    If InStr(Config.Range("C" & i).Value, "USA") <> 0 Then
        Config.Range("C" & i).Value = "US"
    End If
    If InStr(Config.Range("C" & i).Value, "US") = 0 Then
        Config.Range("C" & i).EntireRow.Delete
        If Config.Range("C" & i).Value <> "" Then
            i = i - 1
        End If
    End If
Next i

' Copy Posting Date column to the Settings sheet
Copy_Column_To_Settings_Sheet Vertex, Config, "Posting Date", "D1", Function_Last_Row

' Format and sort Posting Date column
Config.Range("D:D").NumberFormat = "mm/dd/yyyy"
Config.Range("D:D").Sort key1:=Config.Range("D1"), header:=xlYes
Config.Range("D:D").RemoveDuplicates Columns:=Array(1), header:=xlNo
Config_Date_LastRow = WorksheetFunction.CountA(Config.Range("D:D"))

' Hide the Settings sheet and show the User_Interface form
Config.Visible = False
User_Interface.Show
Windows(ThisWorkbook.FullName).WindowState = xlMinimized

End Sub

Sub Copy_Column_To_Settings_Sheet(ByRef Vertex As Worksheet, ByRef Config As Worksheet, ByVal ColumnName As String, ByVal TargetCell As String, ByRef Function_Last_Row As Long)

    Dim ColumnRange As Range
    Dim DestinationColumn As Range

    ' Find the specified column in the Vertex Data sheet
Set ColumnRange = Vertex.Rows(1).Find(What:=ColumnName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
   
    ' If the column is found
    If Not ColumnRange Is Nothing Then
        ' Define the column range
        Set ColumnRange = Range(ColumnRange, ColumnRange.End(xlDown))
        ' Copy the range
        ColumnRange.Copy
        ' Update Function_Last_Row
        Function_Last_Row = WorksheetFunction.CountA(ColumnRange)
        ' Paste the range into the Settings sheet at the specified target cell
        Config.Range(TargetCell).PasteSpecial xlPasteAll
    End If

End Sub
Sub X_FN_Formatting()

Summary.Columns.AutoFit
Summary.Range("B:K").Style = "Comma"

Sage.Range("1:1").Interior.ColorIndex = 15
Sage.Columns.AutoFit

Vertex.Range("1:1").Interior.ColorIndex = 15
Vertex.Columns.AutoFit

End Sub




Sub W_FN_Input_Return_Data()

If CDate(EndDate) < WorksheetFunction.EoMonth(CDate(StartDate), 0) Then

Else
    Dim prompt As String
    Dim title As String
    Dim response As VbMsgBoxResult
    Dim timeout As Long
    
    prompt = "Is return data available?"
    title = "Data Availability"
    timeout = Now + TimeValue("00:00:15")
    
    response = MsgBox(prompt, vbQuestion + vbYesNo, title)
    
    Do While response = vbRetry And Now < timeout
        response = MsgBox(prompt, vbQuestion + vbYesNo, title)
    Loop
    
    If response = vbYes Then
        'Code to retrieve and display the available data
        TaxReturn_Details.Show
    Else
        'Code to handle the case when the user answers No or when the timeout is reached
    End If
    End If
End Sub



Sub FN_Filter_for_Errors()

Dim ER_LR As Long
Dim ER_HR As Long
Dim ER_LC As Long



ER_LR = WorksheetFunction.CountA(Error_Report.Range("A:A"))
ER_HR = 1
ER_LC = WorksheetFunction.CountA(Error_Report.Range("1:1"))

'Error_Report.Range("A1").Value = Vertex.Range("EX1").Value
Vertex.UsedRange.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Error_Report.Range("A" & ER_HR & ":A" & ER_LR)
Sage.UsedRange.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Error_Report.Range("A" & ER_HR & ":A" & ER_LR)


End Sub

Sub FN_Create_Vendor_Upload()

Dim Upload As Worksheet
Dim Tax_Vendor_Lookup As Workbook
Dim MasterBook As Workbook
Dim SaveLocation As String
Dim Last_Row As Long
Dim Tax_Vendor_Row As Long


MsgBox ("Select the Exepense Upload Workbook.")

On Error Resume Next

Set MasterBook = Application.Workbooks.Open(Application.GetOpenFilename)
Set Upload = MasterBook.Sheets(1)
Set Tax_Vendor_Lookup = Application.Workbooks.Open(DropboxDir & "Multi-State Sales Tax\Tax Jursidictions\Settings\Tax Vendor To Vertex Jurisdiction Mapping.xlsx")
Last_Row = WorksheetFunction.CountA(MasterBook.Sheets(1).Range("A:A"))

Tax_Vendor_Row = Tax_Vendor_Lookup.Sheets(1).Cells.Find(ReconState, After:=Tax_Vendor_Lookup.Sheets(1).Range("A1"), SearchOrder:=xlByColumns).row
Upload.Range("A" & Last_Row + 1).Value = "H"
Upload.Range("A" & Last_Row + 2).Value = "L"
Upload.Range("A" & Last_Row + 3).Value = "L"
Upload.Range("A" & Last_Row + 4).Value = "L"


Upload.Range("B" & Last_Row + 1).Value = Tax_Vendor_Lookup.Sheets(1).Cells(Tax_Vendor_Row, 4).Value

Upload.Range("C" & Last_Row + 1).Value = ReconState & " " & Format(StartDate, "yyyy-mm")

Upload.Range("D" & Last_Row + 1).Value = Format(EndDate, "mm/dd/yyyy")

Upload.Range("B" & Last_Row + 2).Value = "9999900A3"
Upload.Range("B" & Last_Row + 3).Value = "243000000"
Upload.Range("B" & Last_Row + 4).Value = "99999009X"


Upload.Range("C" & Last_Row + 2).Value = -Total_Discount_Amt

Upload.Range("D" & Last_Row + 2).Value = Format(EndDate, "yyyy-mm-dd") & " " & ReconState & " Discount of Tax Collected"

Upload.Range("C" & Last_Row + 4).Value = "=-1*(" & Summary.Range("E10").Value - Round(Total_Discount_Amt, 2) - Round(Total_Return_Tax, 2) & ")"
Upload.Range("D" & Last_Row + 3).Value = Format(EndDate, "yyyy-mm-dd") & " " & ReconState & " Tax Liability Payment"
Upload.Range("D" & Last_Row + 4).Value = Format(EndDate, "yyyy-mm-dd") & " " & ReconState & " Rounding Adjustment"
Upload.Columns.AutoFit

Upload.Range("C" & Last_Row + 3).Value = Summary.Range("E10").Value

Upload.Range("C:C").Style = "Comma"
MasterBook.Save
MasterBook.Close
Tax_Vendor_Lookup.Close



End Sub

Sub Step_06_Main_ImportAndFilter_Customer_Exemptions()

    ' Declare variables
    Dim Recon As Workbook
    Dim Cust_Info As Worksheet
    Dim Exemption_Source As Workbook
    Dim Exemptions As Worksheet
    Dim Customer_Data As Range
    Dim i As Long

    ' Set workbook references
    Set Recon = Taxbook
    Set Exemption_Source = Workbooks.Open(DropboxDir & "Multi-State Sales Tax\Tax Jursidictions\Settings\ConfigReport.xlsx")
    Set Exemptions = Recon.Sheets.Add

    ' Set the name of the new sheet
    Exemptions.Name = "Exemption Report"
    Set Cust_Info = Exemption_Source.Sheets("Customer Information")

    ' Add Exemption Reason column
    Step_06_Secondary_AddExemptionReasonColumn Cust_Info

    ' Insert and format Customer ID column
    Step_06_Secondary_AddFormattedCustomerIDColumn Cust_Info

    ' Set a range for the used range in the worksheet
    Set Customer_Data = Cust_Info.UsedRange

    ' Apply filters to the range
    Step_06_Secondary_ApplyFilters Customer_Data

    ' Copy the filtered data and paste it into the Exemptions worksheet
    Step_06_Secondary_CopyFilteredDataToExemptions Customer_Data, Exemptions

    ' Clear the clipboard and autofit columns in the Exemptions worksheet
    Application.CutCopyMode = False
    Exemptions.Columns.AutoFit

    ' Close the Exemption_Source workbook without saving changes
    Exemption_Source.Close SaveChanges:=False

End Sub

Sub Step_06_Secondary_ApplyFilters(Customer_Data As Range)
    ' Apply filters to the range
    Customer_Data.AutoFilter 6, "<>", xlFilterValues
    Customer_Data.AutoFilter 10, ">=" & DateValue(EndDate), xlOr, "="
End Sub

Sub Step_06_Secondary_CopyFilteredDataToExemptions(Customer_Data As Range, Exemptions As Worksheet)
    ' Copy the filtered data and paste it into the Exemptions worksheet
    Customer_Data.SpecialCells(xlCellTypeVisible).Copy
    Exemptions.Range("A1").PasteSpecial xlPasteAll

    ' Clear the clipboard and autofit columns in the Exemptions worksheet
    Application.CutCopyMode = False
    Exemptions.Columns.AutoFit
End Sub

Sub Step_06_Secondary_AddExemptionReasonColumn(Cust_Info As Worksheet)
    ' Insert a new column and set the header for Exemption Reason
    Cust_Info.Range("E:E").Insert
    Cust_Info.Range("E1").Value = "Exemption Reason"

    ' Loop through the rows and set Exemption Reason based on the criteria
    For i = 1 To WorksheetFunction.CountA(Cust_Info.Range("A:A"))
        If InStr(Cust_Info.Range("F" & i).Value, "Exempt") > 0 Then
            Cust_Info.Range("E" & i).Value = Cust_Info.Range("G" & i).Value
        ElseIf InStr(Cust_Info.Range("F" & i).Value, "Taxable") > 0 Then
            Cust_Info.Range("E" & i).Value = Cust_Info.Range("P" & i).Value
        End If
    Next i
End Sub

Sub Step_06_Secondary_AddFormattedCustomerIDColumn(Cust_Info As Worksheet)
    ' Insert a new column and format the values for the new column
    Cust_Info.Range("B:B").Insert

    For i = 2 To WorksheetFunction.CountA(Cust_Info.Range("A:A"))
        If IsNumeric(Cust_Info.Range("C" & i).Value) = True Then
            Cust_Info.Range("B" & i).Value = "'" & Right(Cust_Info.Range("C" & i).Value, 7)
        Else
            Cust_Info.Range("B" & i).Value = Right(Cust_Info.Range("C" & i).Value, 7)
        End If
    Next i
End Sub

Sub Step_04_Main_Import_Zipcode_Database()

'Define the workbook object Import_ZIPCODE_CSV for the zipcode database to be the zipcode database stored on the shared drive
Set Import_ZIPCODE_CSV = Workbooks.Open("T:\Christopher.Bartus\Tax\Zipcode Database.csv")

' Define the sheet object "ZipCode_Source" to be the first sheet
Set ZipCode_Source = Import_ZIPCODE_CSV.Sheets(1)

'If the user indicated that the reconciliation state was "All" and not a specific state then do not filter the zipcode database
If InStr(ReconState, "ALL") = 0 Then

ZipCode_Source.UsedRange.AutoFilter 3, "=" & ReconState
ZipCode_Source.UsedRange.SpecialCells(xlCellTypeVisible).Copy
ZipCode_Data.Range("A1").PasteSpecial (xlPasteValues)
Application.CutCopyMode = False
Import_ZIPCODE_CSV.Saved = True
Import_ZIPCODE_CSV.Close

'Otherwise filter the zipcode database by the state entered by the script user upon execution
Else

ZipCode_Source.UsedRange.SpecialCells(xlCellTypeVisible).Copy
ZipCode_Data.Range("A1").PasteSpecial (xlPasteValues)
Application.CutCopyMode = False
Import_ZIPCODE_CSV.Saved = True
Import_ZIPCODE_CSV.Close
End If

End Sub

Sub Z_FN_Lookup_Exemption_Reason()


End Sub


Sub ZB_FN_Format_Headers()


Dim HeaderCell As Range
Dim HeaderRow As Range
Dim DataRow As Range
Dim DataCell As Range



Set DataRow = Error_Report.Range(Error_Report.Cells(2, 2), Error_Report.Cells(2, WorksheetFunction.CountA(Error_Report.Range("1:1"))))

'For Each DataCell In DataRow
    'If IsNumeric(DataCell.Value) = True And InStr(ActiveSheet.Cells(1, DataCell.Column).Value, "Trimmed Invoice No") = 0 Then
    '    DataCell.EntireColumn.Style = "Comma"
    'ElseIf IsDate(DataCell.Value) And InStr(DataCell.Value, "Invoice No") = True And InStr(ActiveSheet.Cells(1, DataCell.Column).Value, "Date") > 0 Then
    '    DataCell.EntireColumn.NumberFormat = "mm/dd/yyyy"
    '  End If
'Next DataCell

End Sub

Sub ZA_FN_Format_Summary()


Dim Summarysheet As Worksheet
Dim SumLastRow As Long
Set Summarysheet = ActiveSheet
Dim Cell As Range


SumLastRow = Summary.Cells(Summary.Rows.Count, 1).End(xlUp).row

With Summarysheet.PageSetup
.Zoom = False
.FitToPagesTall = 1
.FitToPagesWide = 1
.PrintGridlines = True
.Orientation = xlPortrait
End With
Summarysheet.Columns("I").ColumnWidth = 23.5
Summarysheet.Range("A3").Value = ReconState
With Summarysheet.PageSetup
.PrintArea = "$A$1:$I$" & SumLastRow
.LeftMargin = Application.InchesToPoints(0.25)
.RightMargin = Application.InchesToPoints(0.25)
End With


Summary.Range("A15").Value = TaxReturn_Details.Total_Discount_TB.Value
Summarysheet.Range("A3").Value = ReconState & " Between " & Format(StartDate, "mm/dd/yyyy") & "and " & Format(EndDate, "mm/dd/yyyy")
Summarysheet.Range("1:3").WrapText = True
Summarysheet.Range("A:A").WrapText = True
Summarysheet.Range("1:3").Font.Bold = True
Summarysheet.Range("1:3").Font.Size = 14
Summarysheet.Range("1:3").WrapText = True
Summarysheet.Range("J3:K3").Style = "Normal"
Summary.Range("E1").Font.Color = RGB(4, 60, 116)
Summary.Range("E1").Interior.Color = RGB(195, 213, 66)
Summary.Range("G1").Font.Color = RGB(4, 220, 4)
Summary.Range("G1").Interior.Color = RGB(248, 244, 244)
Summary.Range("B:B").ColumnWidth = 1
Summary.Range("C:C").ColumnWidth = 16
Summary.Range("D:D").ColumnWidth = 1
Summary.Range("E:E").ColumnWidth = 16
Summary.Range("F:F").ColumnWidth = 1
Summary.Range("G:G").ColumnWidth = 16
Summary.Range("H:H").ColumnWidth = 1
Summary.Range("E20").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
Summary.Range("E26").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
Summary.Range("15:18").EntireRow.Delete
Summary.Range("A1:D1").Interior.Color = RGB(180, 198, 231)
Summary.Range("F1").Interior.Color = RGB(180, 198, 231)
Summary.Range("H1:I1").Interior.Color = RGB(180, 198, 231)
For Each Cell In Summary.Range("A1:I" & SumLastRow)
    
    With Cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
    End With
    
    With Cell.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
    End With
    
    With Cell.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
    End With
    
    With Cell.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
    End With
    
Next Cell


If ReconState = "NC" Then
Summary.Range("A4").EntireRow.Insert
Summary.Range("A4").Value = "North Caroline Acct #"
Summary.Range("C4") = "601267099"
Summary.Range("C4").NumberFormat = "General"
Summary.Range("C:C").EntireColumn.AutoFit
Summary.Range("A4").EntireRow.Insert
Summary.Range("A4").Value = "Fotronic Tax ID #"
Summary.Range("C4") = "04-3215575"
Summary.Range("C4").NumberFormat = "General"

End If

Summary.Range("J" & WorksheetFunction.CountA(Summary.Range("J:J")) + 5).Value = Format(EndDate, "YYYY-MM-DD") & " " & ReconState & " Sales Tax Return as Filed"
Summary.Range("J" & WorksheetFunction.CountA(Summary.Range("J:J")) + 5).Value = Format(EndDate, "YYYY-MM-DD") & " " & ReconState & " Sales Tax Payment Confirmation"
Summary.Range("J" & WorksheetFunction.CountA(Summary.Range("J:J")) + 5).Value = "C:\Users\christopher.bartus\Fotronic Dropbox\Christopher Bartus\Multi-State Sales Tax\Tax Jursidictions\" & ReconState & "\" & Format(StartDate, "YYYY") & "\" & Format(StartDate, "YYYY-MM")

Error_Report.Range("A" & WorksheetFunction.CountA(Error_Report.Range("A:A")) + 10).Value = "5 Commonwealth Avenue, Unit 6"
Error_Report.Range("A" & WorksheetFunction.CountA(Error_Report.Range("A:A")) + 10).Value = "Woburn"
Error_Report.Range("A" & WorksheetFunction.CountA(Error_Report.Range("A:A")) + 10).Value = "'01801-1069"



End Sub


Sub Step_09_Main_Vertex_Create_NonTaxable_Pivot_Exclude_Freight()
Dim Filter_PivotItem As PivotItem

'Creates a pivot table report of the Vertex data that excludes freight-related transactions.
'The pivot table is used to analyze non-taxable and exempt sales transactions while excluding freight.
'The code first checks if Sales Tax is included in the Gross Amount.
'Based on the check, the pivot table is created with different calculated fields and sets various pivot fields as row, page, or data fields.
'Filters are applied to exclude freight-related transactions from the report.

'Check if Sales Tax is included in the Gross Amount
If User_Interface.Toggle_SalesTax_In_Gross.Value = True Then

'Create a new worksheet for the Vertex data that excludes freight
Set Vertex_NonTax_Exclude_Freight_Sheet = Taxbook.Sheets.Add
Vertex_NonTax_Exclude_Freight_Sheet.Name = "Vertex ET Non Freight"

' Create the pivot table using the Vertex data cache
    Set Vertex_Nontax_Exclude_Freight = Vertex_Data_Cache.CreatePivotTable( _
        TableDestination:=Vertex_NonTax_Exclude_Freight_Sheet.Range("A5"), _
        TableName:="Vertex_Nontax_Exclude_Freight")

' Set the pivot fields
Set Vertex_Nontax_Exclude_Freight_InvoiceNo = Vertex_Nontax_Exclude_Freight.PivotFields("Trimmed Invoice No")
Set Vertex_Nontax_Exclude_Freight_Product_Class_Code = Vertex_Nontax_Exclude_Freight.PivotFields("Product Class Code")
Set Vertex_Nontax_Exclude_Freight_Amount = Vertex_Nontax_Exclude_Freight.PivotFields("Tax Amount")
Set Vertex_Nontax_Exclude_Freight_Flexcode05 = Vertex_Nontax_Exclude_Freight.PivotFields("Flex Code 5")
Vertex_Nontax_Exclude_Freight.CalculatedFields.Add "GrossAmount (Non Freight)", "= Gross Amount  + 'Total Tax Amount'"
Set Vertex_Nontax_Exclude_Freight_Gross_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("GrossAmount (Non Freight)")
Set Vertex_Nontax_Exclude_Freight_Exempt_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Exempt Amount")
Set Vertex_Nontax_Exclude_Freight_NonTax_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Non-Taxable Amount")
Set Vertex_Nontax_Exclude_Freight_Taxable_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Taxable Amount")
Set Vertex_Nontax_Exclude_Freight_Jurisdiction_Type = Vertex_Nontax_Exclude_Freight.PivotFields("Jurisdiction Type")
Set Vertex_Nontax_Exclude_Freight_ImpositionType = Vertex_Nontax_Exclude_Freight.PivotFields("Imposition Type")

 ' Add calculated fields to the pivot table
Vertex_Nontax_Exclude_Freight.CalculatedFields.Add "Total Exempt Amount", "= 'Exempt Amount' + 'Non-Taxable Amount' + 'Total Tax Amount'"
Set Vertex_Nontax_Exclude_Freight_TotalExempt = Vertex_Nontax_Exclude_Freight.PivotFields("Total Exempt Amount")

' Set the orientation of the pivot fields
Vertex_Nontax_Exclude_Freight_InvoiceNo.Orientation = xlRowField
Vertex_Nontax_Exclude_Freight_Product_Class_Code.Orientation = xlPageField
Vertex_Nontax_Exclude_Freight_Flexcode05.Orientation = xlPageField
Vertex_Nontax_Exclude_Freight_Jurisdiction_Type.Orientation = xlPageField
Vertex_Nontax_Exclude_Freight_ImpositionType.Orientation = xlPageField

'Add data fields to the pivot table
With Vertex_Nontax_Exclude_Freight
    .AddDataField Vertex_Nontax_Exclude_Freight_Gross_Amt, "Vertex Gross Sales (No Freight)", xlSum
    .AddDataField Vertex_Nontax_Exclude_Freight_Taxable_Amt, "Vertex Taxable Sales", xlSum
    .AddDataField Vertex_Nontax_Exclude_Freight_TotalExempt, "Vertex Total Exempt and Excluded Sales", xlSum
End With

 ' Apply filters on the pivot fields
 
 'Filter for General Sales and Use Tax out of the vertex data
With Vertex_Nontax_Exclude_Freight_ImpositionType
    .EnableMultiplePageItems = True
    On Error Resume Next
    .PivotItems("Additional Sales and Use Tax").Visible = False
    .PivotItems("General Sales and Use Tax").Visible = True
End With

'Isolate Data to Include only Jurisdiction Type "State"
For Each Filter_PivotItem In Vertex_Nontax_Exclude_Freight_Jurisdiction_Type.PivotItems
    If Filter_PivotItem.Name <> "STATE" Then
        Filter_PivotItem.Visible = False
    End If
Next Filter_PivotItem

' Isolate Data to filter out Marketplace Freight Sales (Recorded as Code Freight in Flexcode 05)
Vertex_Nontax_Exclude_Freight_Flexcode05.EnableMultiplePageItems = True

For Each Filter_PivotItem In Vertex_Nontax_Exclude_Freight_Flexcode05.PivotItems
    If Filter_PivotItem.Name = "FREIGHT" Then
        Filter_PivotItem.Visible = False
        End If
Next Filter_PivotItem

' Isolate data to filter out non-marketplace freight sales (Recorded as "FREIGHT" in Product Class Code in Vertex Data
For Each Filter_PivotItem In Vertex_Nontax_Exclude_Freight_Product_Class_Code.PivotItems
    If Filter_PivotItem.Name = "FREIGHT" Then
        Filter_PivotItem.Visible = False
        End If
Next Filter_PivotItem

'Define Ranges for Later Data Comparison
Vertex_Nontax_Exclude_Freight.PivotSelect "Vertex Gross Sales (No Freight)", xlDataAndLabel
Set Lookup_Vertex_NontaxNoFreight_Gross = Selection
Vertex_Nontax_Exclude_Freight.PivotSelect "Vertex Taxable Sales", xlDataAndLabel
Set Lookup_Vertex_NontaxNoFreight_Taxable = Selection
Vertex_Nontax_Exclude_Freight.PivotSelect "Vertex Total Exempt and Excluded Sales", xlDataAndLabel
Set Lookup_Vertex_NontaxNoFreight_Exempt = Selection
Vertex_Nontax_Exclude_Freight.rowRange.Select
Set Lookup_Vertex_NontaxNoFreight_InvoiceNo = Selection


Else  ' If Sales Tax is NOT included in the Gross Amount

' Create a new worksheet for the Vertex data that excludes freight
Set Vertex_NonTax_Exclude_Freight_Sheet = Taxbook.Sheets.Add
Vertex_NonTax_Exclude_Freight_Sheet.Name = "Vertex ET Non Freight"

 ' Create the pivot table using the Vertex data cache
Set Vertex_Nontax_Exclude_Freight = Vertex_Data_Cache.CreatePivotTable( _
    TableDestination:=Vertex_NonTax_Exclude_Freight_Sheet.Range("A5"), _
    TableName:="Vertex_Nontax_Exclude_Freight")
    
' Set the pivot fields
    
Set Vertex_Nontax_Exclude_Freight_InvoiceNo = Vertex_Nontax_Exclude_Freight.PivotFields("Trimmed Invoice No")
Set Vertex_Nontax_Exclude_Freight_Product_Class_Code = Vertex_Nontax_Exclude_Freight.PivotFields("Product Class Code")
Set Vertex_Nontax_Exclude_Freight_Amount = Vertex_Nontax_Exclude_Freight.PivotFields("Tax Amount")
Set Vertex_Nontax_Exclude_Freight_Flexcode05 = Vertex_Nontax_Exclude_Freight.PivotFields("Flex Code 5")
Set Vertex_Nontax_Exclude_Freight_Gross_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Gross Amount")
Set Vertex_Nontax_Exclude_Freight_Exempt_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Exempt Amount")
Set Vertex_Nontax_Exclude_Freight_NonTax_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Non-Taxable Amount")
Set Vertex_Nontax_Exclude_Freight_Taxable_Amt = Vertex_Nontax_Exclude_Freight.PivotFields("Taxable Amount")
Set Vertex_Nontax_Exclude_Freight_Jurisdiction_Type = Vertex_Nontax_Exclude_Freight.PivotFields("Jurisdiction Type")
Set Vertex_Nontax_Exclude_Freight_ImpositionType = Vertex_Nontax_Exclude_Freight.PivotFields("Imposition Type")

'Add calculated fields to the pivot table

Vertex_Nontax_Exclude_Freight.CalculatedFields.Add "Total Exempt Amount", "= Gross Amount - Taxable Amount"
Set Vertex_Nontax_Exclude_Freight_TotalExempt = Vertex_Nontax_Exclude_Freight.PivotFields("Total Exempt Amount")

'Set the orientation of the pivot fields
Vertex_Nontax_Exclude_Freight_InvoiceNo.Orientation = xlRowField
Vertex_Nontax_Exclude_Freight_Product_Class_Code.Orientation = xlPageField
Vertex_Nontax_Exclude_Freight_Flexcode05.Orientation = xlPageField
Vertex_Nontax_Exclude_Freight_Jurisdiction_Type.Orientation = xlPageField
Vertex_Nontax_Exclude_Freight_ImpositionType.Orientation = xlPageField


'Add data fields to the pivot table

With Vertex_Nontax_Exclude_Freight
    .AddDataField Vertex_Nontax_Exclude_Freight_Gross_Amt, "Vertex Gross Sales (No Freight)", xlSum
    .AddDataField Vertex_Nontax_Exclude_Freight_Taxable_Amt, "Vertex Taxable Sales", xlSum
    .AddDataField Vertex_Nontax_Exclude_Freight_TotalExempt, "Vertex Total Exempt and Excluded Sales", xlSum
End With

' Filter Imposition Type to be General Sales and Use Tax Only
With Vertex_Nontax_Exclude_Freight_ImpositionType
    .EnableMultiplePageItems = True
    On Error Resume Next
    .PivotItems("Additional Sales and Use Tax").Visible = False
    .PivotItems("Additional Fee").Visible = False
    .PivotItems("General Sales and Use Tax").Visible = True
End With

' Filter Jurisdiction Type to be State Only
For Each Filter_PivotItem In Vertex_Nontax_Exclude_Freight_Jurisdiction_Type.PivotItems
    If Filter_PivotItem.Name <> "STATE" Then
        Filter_PivotItem.Visible = False
    End If
Next Filter_PivotItem

'Filter FlexCode05 to exclude Freight Lines
Vertex_Nontax_Exclude_Freight_Flexcode05.EnableMultiplePageItems = True

For Each Filter_PivotItem In Vertex_Nontax_Exclude_Freight_Flexcode05.PivotItems
    If Filter_PivotItem.Name = "FREIGHT" Then
        Filter_PivotItem.Visible = False
        End If
Next Filter_PivotItem

'Filter Product Class Code to exclude freight items
For Each Filter_PivotItem In Vertex_Nontax_Exclude_Freight_Product_Class_Code.PivotItems
    If Filter_PivotItem.Name = "FREIGHT" Then
        Filter_PivotItem.Visible = False
        End If
Next Filter_PivotItem


'Define Ranges for Later Data Comparison
Vertex_Nontax_Exclude_Freight.PivotSelect "Vertex Gross Sales (No Freight)", xlDataAndLabel
Set Lookup_Vertex_NontaxNoFreight_Gross = Selection
Vertex_Nontax_Exclude_Freight.PivotSelect "Vertex Taxable Sales", xlDataAndLabel
Set Lookup_Vertex_NontaxNoFreight_Taxable = Selection
Vertex_Nontax_Exclude_Freight.PivotSelect "Vertex Total Exempt and Excluded Sales", xlDataAndLabel
Set Lookup_Vertex_NontaxNoFreight_Exempt = Selection
Vertex_Nontax_Exclude_Freight.rowRange.Select
Set Lookup_Vertex_NontaxNoFreight_InvoiceNo = Selection

End If
End Sub

Sub Delete_Empty_Columns()

Dim LastRow As Long
Dim lastCol As Long
Dim i As Integer


Dim Rng As Range
Dim Col As Range

For i = 1 To WorksheetFunction.CountA(ActiveSheet.Range("1:1"))
    If WorksheetFunction.CountA(ActiveSheet.Columns(i)) = 1 Then
    ActiveSheet.Columns(i).EntireColumn.Delete
    i = i - 1
    Else
    End If
Next i

End Sub


Sub Step_15_Main_Compare_Data()

'Compare Taxable and Nontaxable and Gross Sales per document number between Sage and Vertex in various views

Step_15_Secondary_Sage_Compare_NonFreight
Step_15_Secondary_Sage_Compare_Gross
Step_15_Secondary_Sage_Compare_Marketplace
Step_15_Secondary_Vertex_Compare_Gross
Step_15_Secondary_Vertex_Compare_Exempt

End Sub
Sub Step_15_Secondary_Vertex_Compare_Exempt()
'Vertex Nontaxable Sales Pivot Compare

' Get the last column index in the Vertex_NonTax_Compare sheet
Vertex_NonTax_Compare_LC = WorksheetFunction.CountA(Vertex_NonTax_Compare.Range("7:7"))

' Find the row index of "Grand Total" in the Vertex_NonTax_Compare sheet
Vertex_NonTax_Compare_LR = Vertex_NonTax_Compare.Cells.Find("Grand Total", Vertex_NonTax_Compare.Range("A1")).row

' Add column headers to the Vertex_NonTax_Compare sheet
With Vertex_NonTax_Compare
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 1).Value = "Customer Number"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 2).Value = "Sage Gross Sales"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 3).Value = "Gross Sales Variance"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 4).Value = "Sage Taxable Sales"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 5).Value = "Taxable Sales Variance"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 6).Value = "Sage Nontaxable Sales"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 7).Value = "Sage Nontaxable Variance"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 8).Value = "Exemption Reason"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 9).Value = "Vertex Gross Amount Excluding Freight"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 10).Value = "Vertex Exempt/Nontaxable Amount Excluding Freight"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 11).Value = "Vertex Taxable Amount Excluding Freight"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 12).Value = "Sage Gross Amount Excluding Freight"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 13).Value = "Sage Exempt/Nontaxable Amount Excluding Freight"
    .Cells(.Cells.Find("Trimmed Invoice No", After:=.Range("A1")).row, Vertex_NonTax_Compare_LC + 14).Value = "Sage Taxable Amount Excluding Freight"
    .Columns.AutoFit
End With

For i = Vertex_NonTax_Compare.Cells.Find("Trimmed Invoice No", After:=Vertex_NonTax_Compare.Range("A1")).row + 1 To Vertex_NonTax_Compare_LR - 1
Sage_Freight = InStr(WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Sage.Range("EI:EI"), Sage.Range("J:J"), 0, 0), "MKTPLCFAC")

'Lookup Document Customer No.
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 1).Value = "=XLOOKUP(A" & i & "," & Sage.Range("EI:EI").Address(external:=True) & "," & Sage.Range("H:H").Address(external:=True) & ",0,0)"
    
'Calculate Sage Gross nontaxable Sales Sales (Vertex exempt/nontaxable Freight + Sage Nontaxable Sales
        Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 2).Value = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Taxable.Address(external:=True) & ",0,0)+XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Nontaxable.Address(external:=True) & ",0,0)" & "+XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Freight.Address(external:=True) & ",0,0)"
    
'Calculate Gross sales variance
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 3).Value = "=D" & i & "-H" & i
    
'Calculate Sage Taxable Sales (Include Freight)
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 4).Value = WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Lookup_Sage_InvoiceNo, Lookup_Sage_Taxable, 0, 0) + WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Lookup_Vertex_Freight_Sales_InvoiceNo, Lookup_Vertex_Freight_Taxable, 0, 0)
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 4).Value = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Taxable.Address(external:=True) & ",0,0)" & "+XLOOKUP(A" & i & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Taxable.Address(external:=True) & ",0,0)"
'Calculate Taxable Sales Variance
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 5).Value = "=E" & i & "-J" & i
    
' Calculate Nontaxable Sales
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 6).Value = WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Lookup_Sage_InvoiceNo, Lookup_Sage_Nontaxable, 0, 0) + WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Lookup_Vertex_Freight_Sales_InvoiceNo, Lookup_Vertex_Freight_Exempt, 0, 0)
    
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 6).Value = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Nontaxable.Address(external:=True) & ",0,0)" & "+XLOOKUP(A" & i & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Exempt.Address(external:=True) & ",0,0)"
    
'Calculate NonTaxable Sales Variance
        Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 7).Value = "=F" & i & "-L" & i
    
'Determine if ether sage or vertex has more than 0 for nontaxable
    If Vertex_NonTax_Compare.Range("F" & i).Value <> 0 Or Vertex_NonTax_Compare.Range("L" & i).Value <> 0 Then
        'Check to see if the transaction was a marketplace transaction, if its not then...
        If InStr(WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Sage.Range("EI:EI"), Sage.Range("J:J"), 0, 0), "MKTPLCFAC") <> 1 Then
        'Look up the exemption reason from the Exemption Report Sheet based on customer number looked up from Sage AR data
        Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 8).Value = "=XLOOKUP(TEXT(" & Chr(34) & "00" & Chr(34) & Chr(38) & "Xlookup(A" & i & ",'Sage AR Data'!$EI:$EI,'Sage AR Data'!$H:$H,0,0)," & Chr(34) & "000000000" & Chr(34) & "),'Exemption Report'!$C:$C,'Exemption Report'!$F:$F," & Chr(34) & "Nontaxable Item" & Chr(34) & ",0)"
         
         'Otherwise Check to see if its a marketplace transactions
         ElseIf InStr(WorksheetFunction.XLookup(Vertex_NonTax_Compare.Range("A" & i).Value, Sage.Range("EI:EI"), Sage.Range("J:J"), 0, 0), "MKTPLCFAC") = 1 Then
         'if so fill in "Marketplace Sale" into the exemption reason column
        Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 8).Value = "Marketplace Sale"
        Else
        End If
        Else
        'If there are taxable sales put in taxable sale
    Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 8).Value = "Taxable Sale"
    End If
    
'Populate the Vertex Gross Amount Exclusive of Freight
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 9).Formula = "=XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Gross.Address(external:=True) & ",0,0)"

'Populate the Vertex Exempt Amount Exclusive of Freight
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 10).Formula = "=XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Exempt.Address(external:=True) & ",0,0)"

'Populate the Vertex Taxable Amount Excluding Freight
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 11).Formula = "=XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Taxable.Address(external:=True) & ",0,0)"

'Populate the Sage Gross Amount Exclusive of Freight
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 12).Formula = "=XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Nontaxable.Address(external:=True) & ",0,0)+XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Taxable.Address(external:=True) & ",0,0)"

'Populate the Sage Nontaxable Amount Excluding Freight
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 13).Formula = "=XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Nontaxable.Address(external:=True) & ",0,0)"

'Populate the Sage Taxable Amount Excluding Freight
Vertex_NonTax_Compare.Cells(i, Vertex_NonTax_Compare_LC + 14).Formula = "=XLOOKUP(" & Vertex_NonTax_Compare.Range("A" & i).Address(external:=True) & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Taxable.Address(external:=True) & ",0,0)"

Next i



End Sub

Sub Step_15_Secondary_Vertex_Compare_Gross()
'Vertex Gross Sales Pivot Compare

'Get the last column in row 6 of the Vertex_Gross_Sheet
Vertex_Gross_LC = WorksheetFunction.CountA(Vertex_Gross_Sheet.Range("6:6"))

'Populate the specific columns in the Vertex_Gross_Sheet
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 1).Value = "Vertex Freight"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 2).Value = "Vertex Tax"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 3).Value = "Sage Gross Sales"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 4).Value = "Gross Sales Variance"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 5).Value = "Sage Taxable Sales"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 6).Value = "Taxable Sales Variance"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 7).Value = "Sage NonTaxable Sales"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 8).Value = "NonTaxable Sales Variance"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 9).Value = "Sage Freight Amount"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 10).Value = "Freight Variance"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 11).Value = "Sage Tax Amount"
Vertex_Gross_Sheet.Cells(6, Vertex_Gross_LC + 12).Value = "Tax Variance"

'Resize the columns to fit the data
Vertex_Gross_Sheet.Columns.AutoFit

'Find the row number of the cell containing "Grand Total" in the Vertex_Gross_Sheet
Vertex_Pivot_LR = Vertex_Gross_Sheet.Cells.Find("Grand Total", Vertex_Gross_Sheet.Range("A1")).row


Dim ColumnNum As Long
Dim ColumnLetter As String


' Loop through the rows in the Vertex_Gross_Sheet
With Vertex_Gross_Sheet
    For i = 7 To Vertex_Pivot_LR - 1
        
        ' Calculate the Gross Sales for the current row using XLOOKUP formulas
        .Cells(i, Vertex_Gross_LC + 1).Formula = "=XLOOKUP(A" & i & ", " & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Gross.Address(external:=True) & ",0,0)"
        .Cells(i, Vertex_Gross_LC + 2).Formula = "=XLOOKUP(A" & i & "," & Lookup_Vertex_Tax_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Tax_Tax.Address(external:=True) & ",0,0)"
        
        ' Calculate the Gross Sales for the current row in Sage using XLOOKUP formulas
        .Cells(i, Vertex_Gross_LC + 3).Formula = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Gross.Address(external:=True) & ",0,0)"
        
        ' Calculate the Gross Sales Variance for the current row using a formula
        .Cells(i, .Cells.Find("Gross Sales Variance", After:=.Range("A1")).Column).Formula = "=B" & i & "-H" & i

        ' Calculate the Taxable Sales for the current row using XLOOKUP formulas
        .Cells(i, Vertex_Gross_LC + 5).Formula = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Taxable.Address(external:=True) & ",0,0) + XLOOKUP(A" & i & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Taxable.Address(external:=True) & ",0,0)"
        
        ' Calculate the Taxable Sales Variance for the current row using a formula
        .Cells(i, .Cells.Find("Taxable Sales Variance", After:=.Range("A1")).Column).Formula = "=C" & i & "-J" & i

        ' Calculate the Non-Taxable Sales for the current row using XLOOKUP formulas
        .Cells(i, .Cells.Find("Sage NonTaxable Sales", After:=.Range("A1")).Column).Formula = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Nontaxable.Address(external:=True) & ",0,0) + XLOOKUP(A" & i & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Exempt.Address(external:=True) & ",0,0)"
        
        ' Calculate the Non-Taxable Sales Variance for the current row using a formula
        .Cells(i, .Cells.Find("NonTaxable Sales Variance", After:=.Range("A1")).Column).Formula = "=D" & i & "-L" & i

        ' Calculate the Freight Amount for the current row using XLOOKUP formulas
        .Cells(i, .Cells.Find("Sage Freight Amount", After:=.Range("A1")).Column).Formula = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_Sage_Freight.Address(external:=True) & ",0,0)"

        ' Calculate the Freight Variance for the current row using a formula
        .Cells(i, .Cells.Find("Freight Variance", After:=.Range("A1")).Column).Formula = "=F" & i & "-N" & i
    
        ' Calculate the Tax Amount for the current row using XLOOKUP formulas
        .Cells(i, .Cells.Find("Sage Tax Amount", After:=.Range("A1")).Column).Formula = "=XLOOKUP(A" & i & "," & Lookup_Sage_InvoiceNo.Address(external:=True) & "," & Lookup_SageTax.Address(external:=True) & ",0,0)"
        
        ' Calculate the Tax Variance for the current row using a formula
        .Cells(i, .Cells.Find("Tax Variance", After:=.Range("A1")).Column).Formula = "=G" & i & "-P" & i
    Next i
End With

End Sub


Sub Step_15_Secondary_Sage_Compare_NonFreight()

'Populate Data Comparison for Sage Pivot sheet (excludes Marketplace facilitated sales)

'Find the last Row of the Sage Non-Marketplace Faciliated Sales
Sage_Pivot_LC = WorksheetFunction.CountA(Sage_Pivot.Range("6:6"))

'Create Column Headers for different data on sheet "Sage Pivot". This sheet containes a pivot of sage sources sales data excluding marketplace sales

' Create a column to pull in the vertex gross sales for a given record number
Sage_Pivot.Cells(6, Sage_Pivot_LC + 1).Value = "Vertex Gross Sales"

'create a column to calculate any variance between the Vertex Gross Amount, and the Sage Gross Amount (Inclusive of Freight?)
Sage_Pivot.Cells(6, Sage_Pivot_LC + 2).Value = "Gross Sales Variance"

'Create a column to pull vertex taxable sales for a given record number
Sage_Pivot.Cells(6, Sage_Pivot_LC + 3).Value = "Vertex Taxable Sales"

'create a column to calculate any variance between the Vertex taxable sales amount, and the Sage Gross Amount (Inclusive of Freight?)
Sage_Pivot.Cells(6, Sage_Pivot_LC + 4).Value = "Taxable Sales Variance"

'Create a column to pull in looked up Vertex Exempt Sales Amount
Sage_Pivot.Cells(6, Sage_Pivot_LC + 5).Value = "Vertex NonTaxable Sales"

'Create a column to calculate any variance in Nontaxable Sales
Sage_Pivot.Cells(6, Sage_Pivot_LC + 6).Value = "Nontaxable Sales Variance"

'Create a column to hold the amount of sales related to freight
Sage_Pivot.Cells(6, Sage_Pivot_LC + 7).Value = "Vertex Freight Sales"

'Create a column to calculate any variance in freight sales between sage and vertex
Sage_Pivot.Cells(6, Sage_Pivot_LC + 8).Value = "Freight Variance"

'Create a column to look up the tax per record number
Sage_Pivot.Cells(6, Sage_Pivot_LC + 9).Value = "Vertex Tax"

'Create a column to store tax variance
Sage_Pivot.Cells(6, Sage_Pivot_LC + 10).Value = "Tax Variance"

Sage_Pivot.Columns.AutoFit


' Loop Through the pivot table document numbers and write lookup formula. E.G Populate Comparison Fields in Sheet "Sage Pivot".

For i = 7 To WorksheetFunction.CountA(Sage_Pivot.Range("A:A")) + 3
    
    ' Get the cell address of the current invoice number to use in the XLOOKUP formula
    LookupInvoiceNo = Sage_Pivot.Range("A" & i).Address(external:=True)
    
    ' Populate the Gross Vertex Sales Amount Column with an XLOOKUP formula
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 1).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_GrossComp_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_GrossComp_Gross.Address(external:=True) & ",0,0)"
    
    ' Populate the Gross Sales Variance Column
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 2).Value = "=B" & i & "-G" & i
    
    ' Populate the Vertex Taxable Sales Column Excluding Freight with an XLOOKUP formula
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 3).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Taxable.Address(external:=True) & ",0,0)"
    
    ' Populate the Taxable Sales Variance Column
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 4).Value = "=C" & i & "-I" & i
    
    ' Populate the Vertex Exempt Sales Column Excluding Freight with an XLOOKUP formula
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 5).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Exempt.Address(external:=True) & ",0,0)"
    
    ' Populate the Exempt Sales Variance Column
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 6).Value = "=D" & i & "-K" & i
    
    ' Populate the Vertex Freight Sales Column Excluding Marketplace Sales with an XLOOKUP formula
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 7).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Gross.Address(external:=True) & ",0,0)"
    
    ' Populate the Freight Sales Variance Column Excluding Marketplace Sales
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 8).Value = "=E" & i & "-M" & i
    
    ' Populate the Vertex Tax Amount Column with an XLOOKUP formula
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 9).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_Tax_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Tax_Tax.Address(external:=True) & ",0,0)"
    
    ' Populate the Tax Variance Column
    Sage_Pivot.Cells(i, Sage_Pivot_LC + 10).Value = "=F" & i & "-O" & i
Next i


End Sub

Sub Step_15_Secondary_Sage_Compare_Gross()
'Sage Gross Sales (Including Marketplace) Compare to Vertex (Sheet "Sage_Gross Compare" sheet)

' Get the last column of the Sage_Gross worksheet
Sage_Gross_LC = WorksheetFunction.CountA(Sage_Gross.Range("6:6"))

' Add column headers for the Vertex Gross Sales Report
Sage_Gross.Cells(6, Sage_Gross_LC + 1).Value = "Vertex Gross Sales"
Sage_Gross.Cells(6, Sage_Gross_LC + 2).Value = "Gross Sales Variance"
Sage_Gross.Cells(6, Sage_Gross_LC + 3).Value = "Vertex Taxable Sales"
Sage_Gross.Cells(6, Sage_Gross_LC + 4).Value = "Taxable Sales Variance"
Sage_Gross.Cells(6, Sage_Gross_LC + 5).Value = "Vertex Non-Taxable Sales"
Sage_Gross.Cells(6, Sage_Gross_LC + 6).Value = "Non-Taxable Sales Variance"
Sage_Gross.Cells(6, Sage_Gross_LC + 7).Value = "Vertex Freight"
Sage_Gross.Cells(6, Sage_Gross_LC + 8).Value = "Freight Variance"
Sage_Gross.Cells(6, Sage_Gross_LC + 9).Value = "Vertex Tax"
Sage_Gross.Cells(6, Sage_Gross_LC + 10).Value = "Tax Variance"

' Auto-fit the columns in the Sage_Gross worksheet
Sage_Gross.Columns.AutoFit

' Find the last row in the Sage_Gross worksheet where "Grand Total" is located
Sage_Pivot_LR = Sage_Gross.Cells.Find("Grand Total", Sage_Gross.Range("A1")).row


For i = 7 To Sage_Pivot_LR - 1
    ' Get the cell address of the current invoice number to use in the XLOOKUP formula
    LookupInvoiceNo = Sage_Gross.Range("A" & i).Address(external:=True)
    
    ' Populate the Gross Vertex Sales Amount Column with XLOOKUP formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 1).Formula = "=XLOOKUP(A" & i & "," & Lookup_Vertex_Gross_InvoiceNo.Address & "," & Lookup_Vertex_Gross_Gross.Address & ",""ERROR"",0)"
    
    ' Populate the Gross Sales Variance Column with formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 2).Formula = "=B" & i & "-G" & i
    
    ' Populate Column Vertex Taxable Sales Excludes freight based on source with XLOOKUP formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 3).Formula = "=XLOOKUP(A" & i & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address & "," & Lookup_Vertex_NontaxNoFreight_Taxable.Address & ",""ERROR"",0)"
    
    ' Populate the Taxable Sales Variance Column with formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 4).Formula = "=C" & i & "-I" & i
    
    ' Populate the Vertex Exempt Sales Column (Excludes Freight) with XLOOKUP formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 5).Formula = "=XLOOKUP(A" & i & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address & "," & Lookup_Vertex_NontaxNoFreight_Exempt.Address & ",0,0)"
    
    ' Populate the Exempt Sales Variance Column with formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 6).Formula = "=D" & i & "-K" & i
    
    ' Populate the Freight Sales Column (excludes marketplace sales) with XLOOKUP formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 7).Formula = "=XLOOKUP(A" & i & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address & "," & Lookup_Vertex_Freight_Gross.Address & ",0,0)"
    
    ' Populate the Freight Sales Variance Column (excludes marketplace sales) with formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 8).Formula = "=E" & i & "-M" & i
    
    ' Populate the Vertex Tax Amount Column with XLOOKUP formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 9).Formula = "=XLOOKUP(A" & i & "," & Lookup_Vertex_Tax_InvoiceNo.Address & "," & Lookup_Vertex_Tax_Tax.Address & ",0,0)"
    
    ' Populate the Tax Variance Column with formula
    Sage_Gross.Cells(i, Sage_Gross_LC + 10).Formula = "=F" & i & "-O" & i
Next i



' Resize columns in the Sage_Gross worksheet to fit data
Sage_Gross.Columns.AutoFit
End Sub

Sub Step_15_Secondary_Sage_Compare_Marketplace()
'Compare Data Exclusive of Marketplace Sales
' Initialize Mktplc_LC variable to the number of non-empty cells in the 6th row of the Mktplc worksheet
Mktplc_LC = WorksheetFunction.CountA(Mktplc.Range("6:6"))

' Populate the header row of the Mktplc worksheet with column names
Mktplc.Cells(6, Mktplc_LC + 1).Value = "Vertex Gross Sales"
Mktplc.Cells(6, Mktplc_LC + 2).Value = "Gross Sales Variance"
Mktplc.Cells(6, Mktplc_LC + 3).Value = "Vertex Taxable Sales"
Mktplc.Cells(6, Mktplc_LC + 4).Value = "Taxable Sales Variance"
Mktplc.Cells(6, Mktplc_LC + 5).Value = "Vertex NonTaxable Sales"
Mktplc.Cells(6, Mktplc_LC + 6).Value = "Nontaxable Sales Variance"
Mktplc.Cells(6, Mktplc_LC + 7).Value = "Vertex Freight Sales"
Mktplc.Cells(6, Mktplc_LC + 8).Value = "Freight Variance"
Mktplc.Cells(6, Mktplc_LC + 9).Value = "Vertex Tax"
Mktplc.Cells(6, Mktplc_LC + 10).Value = "Tax Variance"

For i = 7 To WorksheetFunction.CountA(Mktplc.Range("A:A")) + 3
    
    ' Get the address of the current invoice number for use in formulas
    LookupInvoiceNo = Mktplc.Range("A" & i).Address(external:=True)
    
    ' Populate the Gross Marketplace Sales Amount Column
    Mktplc.Cells(i, Mktplc_LC + 1).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_GrossComp_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_GrossComp_Gross.Address(external:=True) & ",0,0)"
    
    ' Populate the Gross Sales Variance Column
    Mktplc.Cells(i, Mktplc_LC + 2).Value = "=B" & i & "-G" & i
    
    ' Populate the Taxable Marketplace Sales Column (Excludes Freight)
    Mktplc.Cells(i, Mktplc_LC + 3).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Taxable.Address(external:=True) & ",0,0)"
    
    ' Populate the Taxable Sales Variance Column
    Mktplc.Cells(i, Mktplc_LC + 4).Value = "=C" & i & "-I" & i
    
    ' Populate the Marketplace Exempt Sales Column (Excludes Freight)
    Mktplc.Cells(i, Mktplc_LC + 5).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_NontaxNoFreight_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_NontaxNoFreight_Exempt.Address(external:=True) & ",0,0)"
    
    ' Populate the Exempt Sales Variance Column
    Mktplc.Cells(i, Mktplc_LC + 6).Value = "=D" & i & "-K" & i
    
    ' Populate the Freight Sales Column (excludes non-marketplace sales)
    Mktplc.Cells(i, Mktplc_LC + 7).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_Freight_Sales_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Freight_Gross.Address(external:=True) & ",0,0)"
    
    ' Populate the Freight Sales Variance Column (excludes non-marketplace sales)
    Mktplc.Cells(i, Mktplc_LC + 8).Value = "=E" & i & "-M" & i
    
    ' Populate the Marketplace Tax Amount Column
    Mktplc.Cells(i, Mktplc_LC + 9).Value = "=XLOOKUP(" & LookupInvoiceNo & "," & Lookup_Vertex_Tax_InvoiceNo.Address(external:=True) & "," & Lookup_Vertex_Tax_Tax.Address(external:=True) & ",0,0)"
    
    ' Populate the Tax Variance Column
    Mktplc.Cells(i, Mktplc_LC + 10).Value = "=F" & i & "-O" & i
    
Next i
End Sub



Sub ZD_Create_Index()

Dim Index As Worksheet
Dim Total_Worksheets As Integer


Sheets("Exemption Pivot Data").Visible = False
Sheets("Return Report").Visible = False


Set Index = Sheets.Add(After:=Sheets("Summary"))
Index.Name = "Index"
Index.Range("A1").Value = "Sheet Name"
Index.Range("A1").Font.Bold = True
Index.Range("A1").Interior.Color = RGB(180, 198, 231)
Index.Range("B1").Value = "Sheet Description"
Index.Range("B1").Font.Bold = True
Index.Range("B1").Interior.Color = RGB(180, 198, 231)
Counter = WorksheetFunction.CountA(Index.Range("A:A")) + 1
Total_Worksheets = Worksheets.Count
For i = 1 To Total_Worksheets

If Sheets(i).Visible = False Then
ElseIf Sheets(i).Name = "Index" Then
Else

Index.Range("A" & Counter).Value = Sheets(i).Name


Counter = Counter + 1

End If

If Sheets(i).Name = "Summary" Then
Index.Range("B" & Counter - 1).Value = "Summary of Data between Vertex, Sage, and Tax Return Data. Return Data has to be manually input for comparison"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Vertex ET Non Freight" Then
Index.Range("B" & Counter - 1).Value = "Pivot of Vertex Data by Invoice Number Excluding Freight Sales"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Exemption Report" Then
Index.Range("B" & Counter - 1).Value = "Database of Customer Exemptions Generated from Vertex Config Report"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Error Report" Then
Index.Range("B" & Counter - 1).Value = "Summary of Variances found between the Vertex and Sage Databases."
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "MKTPLCFAC" Then
Index.Range("B" & Counter - 1).Value = "Pivot Table of Sage Marketplace Sales"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Vertex_Nontaxable" Then
Index.Range("B" & Counter - 1).Value = "Working Sheet with Pivot of Vertex data compared to related sage data by document number including freight. Data is used to also determine exemption reason."
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Vertex_Gross Compare" Then
Index.Range("B" & Counter - 1).Value = "Working Sheet with Pivot of Vertex data compared to related sage data by document number including gross freight comparisons. Data is used in generation of Error Report sheet."
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Vertex Pivot" Then
Index.Range("B" & Counter - 1).Value = "Pivots of Vertex Data isolating both gross sales, taxable and exempt (Non taxable + exempt), total tax, and taxable and non-taxable freight for reverse freight taxability lookups. Each pivot is by document number"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Sage_Gross Compare" Then
Index.Range("B" & Counter - 1).Value = "Working Sheet  with Pivots of Sage Data combined with freight data and compared with vertex data used in generation of Error Report"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Sage Pivot" Then
Index.Range("B" & Counter - 1).Value = "Pivot of Sage Sales Excluding Marketplace Sales"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Sage AR Data" Then
Index.Range("B" & Counter - 1).Value = "Raw Sage Data downloaded directly from the back end"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Vertex" Then
Index.Range("B" & Counter - 1).Value = "Raw Vertex Data downloaded directly from the vertex portal and filtered for the reconciliation state"
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit
ElseIf Sheets(i).Name = "Credit Report" Then
Index.Range("B" & Counter - 1).Value = "Downloaded Vertex Credit Report stored on Dropbox. This report is downloaded and saved monthly "
Index.Range("B" & Counter - 1).Select
ActiveCell.Hyperlinks.Add Anchor:=Index.Range("B" & Counter - 1), Address:="", SubAddress:=Sheets(i).Range("A1").Address(external:=True), TextToDisplay:=Index.Range("B" & Counter - 1).Value
Index.Columns.AutoFit

End If

Next i
Index.Range("A1").Activate

Summary.Activate
End Sub

Sub ZD_Create_Sage_Correction_Upload()

Dim SageCorrect As Worksheet
Dim Sage_AR_InvoiceNumber As Range
Dim Sage_AR_LR As Long
Dim Sage_AR_LC As Long
Dim CounterH As Long
Dim CounterL As Long

Set SageCorrect = Sheets.Add
SageCorrect.Name = "SageUpload"
'SageCorrect.Range("A1").Value = ""
'SageCorrect.Range("B1").Value = "H.InvoiceNo"
'SageCorrect.Range("C1").Value = "H.InvoiceDate"
'SageCorrect.Range("D1").Value = "H.InvoiceType"
'SageCorrect.Range("E1").Value = "H.CustomerNo"
'SageCorrect.Range("F1").Value = "H.BillToName"
'SageCorrect.Range("G1").Value = "H.BillToAddress1"
'SageCorrect.Range("H1").Value = "H.BillToAddress2"
'SageCorrect.Range("I1").Value = "H.BillToAddress3"
'SageCorrect.Range("J1").Value = "H.BillToCity"
'SageCorrect.Range("K1").Value = "H.BillToState"
'SageCorrect.Range("L1").Value = "H.BillToZipCode"
'SageCorrect.Range("M1").Value = "H.BillToCountryCode"
'SageCorrect.Range("N1").Value = "H.ShipToCode"
'SageCorrect.Range("O1").Value = "H.ShipToName"
'SageCorrect.Range("P1").Value = "H.ShipToAddress1"
'SageCorrect.Range("Q1").Value = "H.ShipToAddress2"
'SageCorrect.Range("R1").Value = "H.ShipToAddress3"
'SageCorrect.Range("S1").Value = "H.ShipToCity"
'SageCorrect.Range("T1").Value = "H.ShipToState"
'SageCorrect.Range("U1").Value = "H.ShipToZipcode"
'SageCorrect.Range("V1").Value = "H.ShipToCountryCode"
'SageCorrect.Range("W1").Value = "H.TaxSchedule"
'SageCorrect.Range("X1").Value = "H.ApplyToInvoiceNo"
'SageCorrect.Range("Y1").Value = "H.UDF_ORDER_MANAGER"
'SageCorrect.Range("Z1").Value = "H.PaymentType"

'SageCorrect.Range("A2").Value = ""
'SageCorrect.Range("B2").Value = "L.ItemCode"
'SageCorrect.Range("C2").Value = "L.ItemType"
'SageCorrect.Range("D2").Value = "L.QuanityOrdered"
'SageCorrect.Range("E2").Value = "L.Quantity Shipped"
'SageCorrect.Range("F2").Value = "L.UnitPrice"
'SageCorrect.Range("G2").Value = "L.UnitCost"
'SageCorrect.Range("H2").Value = "L.TaxAmt"

SageCorrect.Columns.AutoFit


For i = 3 To Sheets("Sage AR Data").Cells(Rows.Count, 1).End(xlUp).row

CounterH = WorksheetFunction.CountA(SageCorrect.Range("A:A")) + 1
CounterL = WorksheetFunction.CountA(SageCorrect.Range("A:A")) + 2



' If the Tax Schedule is not Vertex and
' If the customer number is in the vertex exemption listing
' And the Sage Tax Amount is Greater than 0
' And the Vertex Tax Amount is Greater Then 0 Then Create a credit memo

If _
    ActiveWorkbook.Sheets("Error Report").Range("G" & i).Value = "VERTEX" And _
    ActiveWorkbook.Sheets("Error Report").Range("M" & i).Value <> "ERROR" And _
    ActiveWorkbook.Sheets("Error Report").Range("R" & i).Value > 0 And _
    ActiveWorkbook.Sheets("Error Report").Range("AA" & i).Value = 0 Then
    
'Header Designator
SageCorrect.Range("A" & CounterH).Value = "H"
'Document Number to be Created
SageCorrect.Range("B" & CounterH).Value = "C" & Sheets("Error Report").Range("A" & i).Value

'Document Date
SageCorrect.Range("C" & CounterH).Value = Sheets("Error Report").Range("B" & i).Value
'Invoice Type
SageCorrect.Range("D" & CounterH).Value = "CM"

'Customer No
SageCorrect.Range("E" & CounterH).Value = "=TEXT(" & Sheets("Error Report").Range("C" & i).Address(external:=True) & ", " & Chr(34) & "0000000" & Chr(34) & ")"


'Bill To Name
SageCorrect.Range("F" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AC:AC," & Chr(34) & "ERROR" & Chr(34) & ",0)"


'Bill To Address 1
SageCorrect.Range("G" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AD:AD," & Chr(34) & "ERROR" & Chr(34) & ",0)"


'Bill To Address 2
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("H" & CounterH).Value = ""
Else
SageCorrect.Range("H" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Bill To Address 3
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("I" & CounterH).Value = ""
Else
SageCorrect.Range("I" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Bill to City

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("J" & CounterH).Value = ""
Else
SageCorrect.Range("J" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToState
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("K" & CounterH).Value = ""
Else
SageCorrect.Range("K" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("L" & CounterH).Value = ""
Else
SageCorrect.Range("L" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToCountryCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("M" & CounterH).Value = ""
Else
SageCorrect.Range("M" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToCode
'If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
'SageCorrect.Range("N" & CounterH).Value = ""
'Else
'SageCorrect.Range("N" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)"
'End If

'ShipToName
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("O" & CounterH).Value = ""
Else
SageCorrect.Range("O" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress1
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("P" & CounterH).Value = ""
Else
SageCorrect.Range("P" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress2
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("Q" & CounterH).Value = ""
Else
SageCorrect.Range("Q" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress3
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("R" & CounterH).Value = ""
Else
SageCorrect.Range("R" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShiptoCity
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("S" & CounterH).Value = ""
Else
SageCorrect.Range("S" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToState
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("T" & CounterH).Value = ""
Else
SageCorrect.Range("T" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("U" & CounterH).Value = ""
Else
SageCorrect.Range("U" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("V" & CounterH).Value = ""
Else
SageCorrect.Range("V" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Tax Schedule

SageCorrect.Range("W" & CounterH).Value = "DEFAULT"

'ApptoInvoice

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("X" & CounterH).Value = ""
Else
SageCorrect.Range("X" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)"

End If

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("Y" & CounterH).Value = ""
Else
SageCorrect.Range("Y" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"

End If

SageCorrect.Range("Z" & CounterH).Value = ""

'Begin Lines

'Designator
SageCorrect.Range("A" & CounterL).Value = "L"
SageCorrect.Range("A" & CounterL + 1).Value = "L"

'LineItem Code
SageCorrect.Range("B" & CounterL).Value = "/TAX-VERTEX"
SageCorrect.Range("B" & CounterL + 1).Value = "/NONTAX-VERTEX"

'LineItem Type
SageCorrect.Range("C" & CounterL).Value = "5"
SageCorrect.Range("C" & CounterL + 1).Value = "5"

'LineItem Ordered
SageCorrect.Range("D" & CounterL).Value = "1"
SageCorrect.Range("D" & CounterL + 1).Value = "1"

'LineItem Shipped
SageCorrect.Range("E" & CounterL).Value = "1"
SageCorrect.Range("E" & CounterL + 1).Value = "1"

'LineItem UnitPrice
SageCorrect.Range("F" & CounterL).Value = "=" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL).NumberFormat = "General"
SageCorrect.Range("F" & CounterL + 1).Value = "=-" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL + 1).NumberFormat = "General"

'LineItem Cost
SageCorrect.Range("G" & CounterL).Value = "0"
SageCorrect.Range("G" & CounterL + 1).Value = "0"


'LineItem Comment
SageCorrect.Range("H" & CounterL).Value = Left(Sheets("Error Report").Range("AC" & i).Value, InStr(Sheets("Error Report").Range("AC" & i).Value, "."))
SageCorrect.Range("H" & CounterL + 1).Value = Left(Sheets("Error Report").Range("AC" & i).Value, InStr(Sheets("Error Report").Range("AC" & i).Value, "."))

'LineItem TaxAmt
'SageCorrect.Range("H" & CounterL).Value = "=-" & Sheets("Error Report").Range("R" & I).Address(External:=True)'
'SageCorrect.Range("H" & CounterL).NumberFormat = "General"
'SageCorrect.Range("H" & CounterL + 1).Value = "0"
'SageCorrect.Range("H" & CounterL + 1).NumberFormat = "General"


ElseIf _
    ActiveWorkbook.Sheets("Error Report").Range("G" & i).Value = "VERTEX" And _
    ActiveWorkbook.Sheets("Error Report").Range("M" & i).Value = "ERROR" And _
    ActiveWorkbook.Sheets("Error Report").Range("R" & i).Value < ActiveWorkbook.Sheets("Error Report").Range("AA" & i).Value Then
    
    'Header Designator
SageCorrect.Range("A" & CounterH).Value = "H"

'Document Number to be Created
SageCorrect.Range("B" & CounterH).Value = "C" & Sheets("Error Report").Range("A" & i).Value


SageCorrect.Range("C" & CounterH).Value = Sheets("Error Report").Range("B" & i).Value

'Invoice Type
SageCorrect.Range("D" & CounterH).Value = "CM"

'InvoiceType
SageCorrect.Range("E" & CounterH).Value = Sheets("Error Report").Range("C" & i).Value


'Bill To Name
SageCorrect.Range("F" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AC:AC," & Chr(34) & "ERROR" & Chr(34) & ",0)"


'Bill To Address 1
SageCorrect.Range("G" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AD:AD," & Chr(34) & "ERROR" & Chr(34) & ",0)"


'Bill To Address 2
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("H" & CounterH).Value = ""
Else
SageCorrect.Range("H" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Bill To Address 3
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("I" & CounterH).Value = ""
Else
SageCorrect.Range("I" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Bill to City

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("J" & CounterH).Value = ""
Else
SageCorrect.Range("J" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToState
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("K" & CounterH).Value = ""
Else
SageCorrect.Range("K" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("L" & CounterH).Value = ""
Else
SageCorrect.Range("L" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToCountryCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("M" & CounterH).Value = ""
Else
SageCorrect.Range("M" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToCode
'If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
'SageCorrect.Range("N" & CounterH).Value = ""
'Else
'SageCorrect.Range("N" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)"
'End If

'ShipToName
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("O" & CounterH).Value = ""
Else
SageCorrect.Range("O" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress1
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("P" & CounterH).Value = ""
Else
SageCorrect.Range("P" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress2
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("Q" & CounterH).Value = ""
Else
SageCorrect.Range("Q" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress3
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("R" & CounterH).Value = ""
Else
SageCorrect.Range("R" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShiptoCity
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("S" & CounterH).Value = ""
Else
SageCorrect.Range("S" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToState
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("T" & CounterH).Value = ""
Else
SageCorrect.Range("T" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("U" & CounterH).Value = ""
Else
SageCorrect.Range("U" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("V" & CounterH).Value = ""
Else
SageCorrect.Range("V" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Tax Schedule

SageCorrect.Range("W" & CounterH).Value = "DEFAULT"

'ApptoInvoice

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("X" & CounterH).Value = ""
Else
SageCorrect.Range("X" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)"

End If

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("Y" & CounterH).Value = ""
Else
SageCorrect.Range("Y" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"

End If

SageCorrect.Range("Z" & CounterH).Value = ""

'Begin Lines

'Designator
SageCorrect.Range("A" & CounterL).Value = "L"
SageCorrect.Range("A" & CounterL + 1).Value = "L"

'LineItem Code
SageCorrect.Range("B" & CounterL).Value = "/TAX-VERTEX"
SageCorrect.Range("B" & CounterL + 1).Value = "/NONTAX-VERTEX"

'LineItem Type
SageCorrect.Range("C" & CounterL).Value = "5"
SageCorrect.Range("C" & CounterL + 1).Value = "5"

'LineItem Ordered
SageCorrect.Range("D" & CounterL).Value = "1"
SageCorrect.Range("D" & CounterL + 1).Value = "1"

'LineItem Shipped
SageCorrect.Range("E" & CounterL).Value = "1"
SageCorrect.Range("E" & CounterL + 1).Value = "1"

'LineItem UnitPrice
SageCorrect.Range("F" & CounterL).Value = "=" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL).NumberFormat = "General"
SageCorrect.Range("F" & CounterL + 1).Value = "=-" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL + 1).NumberFormat = "General"

'LineItem Cost
SageCorrect.Range("G" & CounterL).Value = "0"
SageCorrect.Range("G" & CounterL + 1).Value = "0"



End If
    
Next i

'----

For i = 3 To Sheets("Sage AR Data").Cells(Rows.Count, 1).End(xlUp).row

CounterH = WorksheetFunction.CountA(SageCorrect.Range("A:A")) + 1
CounterL = WorksheetFunction.CountA(SageCorrect.Range("A:A")) + 2


If _
    ActiveWorkbook.Sheets("Error Report").Range("G" & i).Value = "VERTEX" And _
    ActiveWorkbook.Sheets("Error Report").Range("M" & i).Value = "ERROR-ERROR" And _
    ActiveWorkbook.Sheets("Error Report").Range("R" & i).Value > 0 And _
    ActiveWorkbook.Sheets("Error Report").Range("AA" & i).Value > 0 Then
    
'Header Designator
SageCorrect.Range("A" & CounterH).Value = "H"

'Document Number to be Created
SageCorrect.Range("B" & CounterH).Value = "C" & Sheets("Error Report").Range("A" & i).Value
SageCorrect.Range("C" & CounterH).Value = Sheets("Error Report").Range("B" & i).Value

'Invoice Type

    If ActiveWorkbook.Sheets("Error Report").Range("R" & i).Value < ActiveWorkbook.Sheets("Error Report").Range("AA" & i).Value Then ' If Sage Sales Tax is Less than Vertex Sales Tax Then type is Debit Memo Otherwise Type is Creditmemo
        SageCorrect.Range("D" & CounterH).Value = "DM"
        Else
            SageCorrect.Range("D" & CounterH).Value = "CM"
    End If

    'InvoiceType
    SageCorrect.Range("E" & CounterH).Value = Sheets("Error Report").Range("C" & i).Value


    'Bill To Name
    SageCorrect.Range("F" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AC:AC," & Chr(34) & "ERROR" & Chr(34) & ",0)"


    'Bill To Address 1
    SageCorrect.Range("G" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AD:AD," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    
    
    'Bill To Address 2
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("H" & CounterH).Value = ""
    Else
    SageCorrect.Range("H" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If
    
    'Bill To Address 3
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("I" & CounterH).Value = ""
    Else
    SageCorrect.Range("I" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If
    
    'Bill to City
    
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("J" & CounterH).Value = ""
    Else
    SageCorrect.Range("J" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If
    
    'BillToState
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("K" & CounterH).Value = ""
    Else
    SageCorrect.Range("K" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If
    
    'BillToZipCode
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("L" & CounterH).Value = ""
    Else
    SageCorrect.Range("L" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If
    
    'BillToCountryCode
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("M" & CounterH).Value = ""
    Else
    SageCorrect.Range("M" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToCode
    'If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    'SageCorrect.Range("N" & CounterH).Value = ""
    'Else
    'SageCorrect.Range("N" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    'End If

    'ShipToName
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("O" & CounterH).Value = ""
    Else
    SageCorrect.Range("O" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToAddress1
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("P" & CounterH).Value = ""
    Else
    SageCorrect.Range("P" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToAddress2
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("Q" & CounterH).Value = ""
    Else
    SageCorrect.Range("Q" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToAddress3
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("R" & CounterH).Value = ""
    Else
    SageCorrect.Range("R" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShiptoCity
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("S" & CounterH).Value = ""
    Else
    SageCorrect.Range("S" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToState
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("T" & CounterH).Value = ""
    Else
    SageCorrect.Range("T" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToZipCode
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("U" & CounterH).Value = ""
    Else
    SageCorrect.Range("U" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'ShipToZipCode
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
        SageCorrect.Range("V" & CounterH).Value = ""
        Else
        SageCorrect.Range("V" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    'Tax Schedule
        SageCorrect.Range("W" & CounterH).Value = "DEFAULT"

    'ApptoInvoice
    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
    SageCorrect.Range("X" & CounterH).Value = ""
    Else
    SageCorrect.Range("X" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
        SageCorrect.Range("Y" & CounterH).Value = ""
            Else
                SageCorrect.Range("Y" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"
    End If

    SageCorrect.Range("Z" & CounterH).Value = ""

'Begin Lines

'Designator
SageCorrect.Range("A" & CounterL).Value = "L"
SageCorrect.Range("A" & CounterL + 1).Value = "L"

'LineItem Code
SageCorrect.Range("B" & CounterL).Value = "/TAX-VERTEX"
SageCorrect.Range("B" & CounterL + 1).Value = "/TAX-VERTEX"

'LineItem Type
SageCorrect.Range("C" & CounterL).Value = "5"
SageCorrect.Range("C" & CounterL + 1).Value = "5"

'LineItem Ordered
SageCorrect.Range("D" & CounterL).Value = "1"
SageCorrect.Range("D" & CounterL + 1).Value = "1"

'LineItem Shipped
SageCorrect.Range("E" & CounterL).Value = "1"
SageCorrect.Range("E" & CounterL + 1).Value = "1"

'LineItem UnitPrice
SageCorrect.Range("F" & CounterL).Value = "=" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL).NumberFormat = "General"
SageCorrect.Range("F" & CounterL + 1).Value = "=-" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL + 1).NumberFormat = "General"

'LineItem Cost
SageCorrect.Range("G" & CounterL).Value = "0"
SageCorrect.Range("G" & CounterL + 1).Value = "0"

'LineItem Comment
SageCorrect.Range("H" & CounterL).Value = Left(Sheets("Error Report").Range("AC" & i).Value, InStr(Sheets("Error Report").Range("AC" & i).Value, "."))
SageCorrect.Range("H" & CounterL + 1).Value = Left(Sheets("Error Report").Range("AC" & i).Value, InStr(Sheets("Error Report").Range("AC" & i).Value, "."))

'LineItem TaxAmt
'SageCorrect.Range("H" & CounterL).Value = "=-" & Sheets("Error Report").Range("R" & I).Address(External:=True)'
'SageCorrect.Range("H" & CounterL).NumberFormat = "General"
'SageCorrect.Range("H" & CounterL + 1).Value = "0"
'SageCorrect.Range("H" & CounterL + 1).NumberFormat = "General"


ElseIf _
    ActiveWorkbook.Sheets("Error Report").Range("G" & i).Value = "VERTEX" And _
    ActiveWorkbook.Sheets("Error Report").Range("M" & i).Value = "ERROR" And _
    ActiveWorkbook.Sheets("Error Report").Range("R" & i).Value < ActiveWorkbook.Sheets("Error Report").Range("AA" & i).Value Then
    
    'Header Designator
SageCorrect.Range("A" & CounterH).Value = "H"

'Document Number to be Created
SageCorrect.Range("B" & CounterH).Value = "C" & Sheets("Error Report").Range("A" & i).Value


SageCorrect.Range("C" & CounterH).Value = Sheets("Error Report").Range("B" & i).Value

'Invoice Type
SageCorrect.Range("D" & CounterH).Value = "CM"

'InvoiceType
SageCorrect.Range("E" & CounterH).Value = Sheets("Error Report").Range("C" & i).Value


'Bill To Name
SageCorrect.Range("F" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AC:AC," & Chr(34) & "ERROR" & Chr(34) & ",0)"


'Bill To Address 1
SageCorrect.Range("G" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AD:AD," & Chr(34) & "ERROR" & Chr(34) & ",0)"


'Bill To Address 2
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("H" & CounterH).Value = ""
Else
SageCorrect.Range("H" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AE:AE," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Bill To Address 3
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("I" & CounterH).Value = ""
Else
SageCorrect.Range("I" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AF:AF," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Bill to City

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("J" & CounterH).Value = ""
Else
SageCorrect.Range("J" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AG:AG," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToState
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("K" & CounterH).Value = ""
Else
SageCorrect.Range("K" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AH:AH," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("L" & CounterH).Value = ""
Else
SageCorrect.Range("L" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AI:AI," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'BillToCountryCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("M" & CounterH).Value = ""
Else
SageCorrect.Range("M" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AJ:AJ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToCode
'If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
'SageCorrect.Range("N" & CounterH).Value = ""
'Else
'SageCorrect.Range("N" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & I).Address(External:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AK:AK," & Chr(34) & "ERROR" & Chr(34) & ",0)"
'End If

'ShipToName
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("O" & CounterH).Value = ""
Else
SageCorrect.Range("O" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AL:AL," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress1
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("P" & CounterH).Value = ""
Else
SageCorrect.Range("P" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AM:AM," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress2
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("Q" & CounterH).Value = ""
Else
SageCorrect.Range("Q" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AN:AN," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToAddress3
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("R" & CounterH).Value = ""
Else
SageCorrect.Range("R" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AO:AO," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShiptoCity
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("S" & CounterH).Value = ""
Else
SageCorrect.Range("S" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AP:AP," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToState
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("T" & CounterH).Value = ""
Else
SageCorrect.Range("T" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AQ:AQ," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("U" & CounterH).Value = ""
Else
SageCorrect.Range("U" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AR:AR," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'ShipToZipCode
If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("V" & CounterH).Value = ""
Else
SageCorrect.Range("V" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!AS:AS," & Chr(34) & "ERROR" & Chr(34) & ",0)"
End If

'Tax Schedule

SageCorrect.Range("W" & CounterH).Value = "DEFAULT"

'ApptoInvoice

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("X" & CounterH).Value = ""
Else
SageCorrect.Range("X" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!A:A," & Chr(34) & "ERROR" & Chr(34) & ",0)"

End If

If Evaluate("=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)") = 0 Then
SageCorrect.Range("Y" & CounterH).Value = ""
Else
SageCorrect.Range("Y" & CounterH).Value = "=XLOOKUP(" & Sheets("Error Report").Range("A" & i).Address(external:=True) & ",'Sage AR Data'!EH:EH,'Sage AR Data'!CY:CY," & Chr(34) & "ERROR" & Chr(34) & ",0)"

End If

SageCorrect.Range("Z" & CounterH).Value = ""

'Begin Lines

'Designator
SageCorrect.Range("A" & CounterL).Value = "L"
SageCorrect.Range("A" & CounterL + 1).Value = "L"

'LineItem Code
SageCorrect.Range("B" & CounterL).Value = "/TAX-VERTEX"
SageCorrect.Range("B" & CounterL + 1).Value = "/NONTAX-VERTEX"

'LineItem Type
SageCorrect.Range("C" & CounterL).Value = "5"
SageCorrect.Range("C" & CounterL + 1).Value = "5"

'LineItem Ordered
SageCorrect.Range("D" & CounterL).Value = "1"
SageCorrect.Range("D" & CounterL + 1).Value = "1"

'LineItem Shipped
SageCorrect.Range("E" & CounterL).Value = "1"
SageCorrect.Range("E" & CounterL + 1).Value = "1"

'LineItem UnitPrice
SageCorrect.Range("F" & CounterL).Value = "=" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL).NumberFormat = "General"
SageCorrect.Range("F" & CounterL + 1).Value = "=-" & Sheets("Error Report").Range("O" & i).Address(external:=True)
SageCorrect.Range("F" & CounterL + 1).NumberFormat = "General"

'LineItem Cost
SageCorrect.Range("G" & CounterL).Value = "0"
SageCorrect.Range("G" & CounterL + 1).Value = "0"

'LineItem Comment
SageCorrect.Range("H" & CounterL).Value = Left(Sheets("Error Report").Range("AC" & i).Value, InStr(Sheets("Error Report").Range("AC" & i).Value, "."))
SageCorrect.Range("H" & CounterL + 1).Value = Left(Sheets("Error Report").Range("AC" & i).Value, InStr(Sheets("Error Report").Range("AC" & i).Value, "."))


End If
    
Next i
ActiveWorkbook.SaveAs FileName:=Format(EndDate, "yyyy-mm-dd") & " " & ReconState & " Sage Adjustment Upload.csv", FileFormat:=xlCSV
On Error Resume Next
ActiveWorkbook.SaveAs FileName:=SaveDir & "\" & FileName, FileFormat:=xlWorkbookDefault
End Sub

Sub ZE_Compare_State()

Set Sage = Sheets("Sage AR Data")
Set Vertex = Sheets("Vertex")

Sage.Cells(1, WorksheetFunction.CountA(Sage.Range("1:1")) + 1).Value = "Vertex State"
For i = 2 To WorksheetFunction.CountA(Sage.Range("A:A"))
Sage.Cells(i, WorksheetFunction.CountA(Sage.Range("1:1"))).Value = WorksheetFunction.XLookup(Sage.Range("EH" & i), Vertex.Range("EX:EX"), Vertex.Range("CE:CE"), "Did not find record", 0)

Next i


End Sub

Sub PromptUser()

    Dim prompt As String
    Dim title As String
    Dim response As VbMsgBoxResult
    Dim timeout As Long
    
    prompt = "Is return data available?"
    title = "Data Availability"
    timeout = Now + TimeValue("00:00:15")
    
    response = MsgBox(prompt, vbQuestion + vbYesNo, title)
    
    Do While response = vbRetry And Now < timeout
        response = MsgBox(prompt, vbQuestion + vbYesNo, title)
    Loop
    
    If response = vbYes Then
        'Code to retrieve and display the available data
        TaxReturn_Details.Show
    Else
        'Code to handle the case when the user answers No or when the timeout is reached
        
    End If
    
End Sub

Sub CleanErrorReport()
    Dim ErrorReport As Worksheet
    Dim ShipToState As Range
    Dim ShipToStateRange As Range
    Dim ShipToStateCol As Long
    Dim Starting As Long
    Dim Ending As Long
    Dim i As Long
    Application.ScreenUpdating = False
    Set ErrorReport = ActiveWorkbook.Sheets("Error Report")
    Starting = 3 + 1
    Ending = LastRow(ErrorReport)
    ShipToStateCol = Application.Match("Ship To State", ErrorReport.Range("3:3"), 0)
    Set ShipToStateRange = ErrorReport.Range(ErrorReport.Cells(Starting, ShipToStateCol), ErrorReport.Cells(Ending, ShipToStateCol))
    
    For i = ShipToStateRange.Rows.Count To 4 Step -1
        Set ShipToState = ErrorReport.Cells(i, ShipToStateCol)
        
        If ShipToState.Value = "" Then
            ' Skip blank cells
        ElseIf CheckState(ShipToState.Value) = False Then
            ' Delete row if the ShipToState is not in the state list
            ShipToState.EntireRow.Delete
        ElseIf ShipToState.Value = "FL" And ErrorReport.Cells(ShipToState.row, Application.Match("Tax Schedule", ErrorReport.Range("3:3"), 0)).Value = "MKTPLCFAC" Then
            ' Delete row if the ShipToState is FL and TaxSchedule is Mktplcfac
            ShipToState.EntireRow.Delete
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub


Function LastRow(ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
End Function
Function CheckState(str As String) As Boolean
    Dim stateList() As Variant
    Dim i As Integer
    Dim matchFound As Boolean
    
    'Initialize array with list of states
    stateList = Array("AL", "AR", "AZ", "CA", "CO", "CT", "DC", "FL", "GA", "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MS", "NC", "ND", "NE", "NJ", "NM", "NV", "NY", "OH", "OK", "PA", "RI", "SC", "TN", "TX", "UT", "VA", "WA", "WI", "WV")
    
    ' Loop through array and check for match
    For i = LBound(stateList) To UBound(stateList)
        If str = stateList(i) Then
            matchFound = True
            Exit For
        End If
    Next i
    
    CheckState = matchFound
End Function


