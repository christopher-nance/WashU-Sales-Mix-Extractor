# Built by Christopher Nance for WashU Car Wash
# Version 4.1
# Sales Mix Report Generator

# Dependencies:
# > Python 3.10
# > Pandas
# > Power Automate
# > Dropbox
# > SiteWatch Cloud (SiteWatch Version 26+)
# > Microsoft Azure Cloud Functions


import pandas as pd
import openpyxl
from openpyxl.chart.plotarea import DataTable
from openpyxl.styles import NamedStyle, Font, Alignment, PatternFill
from openpyxl.chart import BarChart, Reference, PieChart, ProjectedPieChart
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.chart.label import DataLabelList
from copy import copy
from datetime import datetime, date
import re



def createSalesMixSheet(gsrFilePath, templateFilePath, fileNameForParser, trendsFilePath=None, excludedLocations=None):
    #--> Create Dictionaries for looking up item names.
    # TODO: This needs to be attached to a database and allow insertions through the manager portal or a form.
    Corporation_Totals_Name = 'Corporation Totals'
    global monthsInReport
    monthsInReport = 1
    desplainesOverride = False

    WashPkgPrices = {
        "express wash": 6,
        "clean wash": 10,
        "protect wash": 15,
        "ushine wash": 20
    }

    ARM_Sold_Names = {
        "Express Wash": {
            "New Mnthly Express",
        },
        "Clean Wash": {
            "Mnth Cln 9.95 (BOGO)",
            "Monthly Cln '21 9.95",
            "Monthly Cln '21 Sld",
        },
        "Protect Wash": {
            "Mnthly Protect Promo",
            "New Mnthly Protect",
        },
        "UShine Wash": {
            "Mnth USh 9.95 (BOGO)",
            "Monthly UShine Promo",
            "Monthly UShine Sld",
        },
    }

    ARM_Recharge_Names = {
        "Express Wash": {
            "New Mnthly Exp Rchg",
            "New Monthly Expr Rcg",
        },
        "Clean Wash": {
            "Monthly Cln '21 Rchg",
        },
        "Protect Wash": {
            "New Mnthly Prot Rchg",
            "Mnthly Prot 6mo Rchg"
        },
        "UShine Wash": {
            "Monthly UShine Rchg"
        },
    }

    ARM_Termination_Names = {
        "Express Wash": {
            "New Mnthly Exp NoRfn",
            "New Mnthly Exp Rfnd",
            "New Monthly Exp Rfnd",
            "New Monthly ExpNoRfn",
        },
        "Clean Wash": {
            "Monthly Cln '21 Rfnd",
            "MonthlyCln'21 NoRfnd",
        },
        "Protect Wash": {
            "Mnthly Prot 6m NRfnd",
            "Mnthly Prot 6mo Rfnd",
            "New Mnthly Pro NoRfn",
            "New Mnthly Prot Rfnd",
        },
        "UShine Wash": {
            "Monthly UShine NRfnd",
            "Monthly UShine Rfnd",
        },
    }

    ARM_PKG_Names = {
        "Express Wash": {
            "New Mnthly Exp Rdmd",
            "New Monthly Expr Rdm",
            "3 Mo Express Rdmd",
            "CLUB-express Rdmd",
            "COMP-CLUB-xprs Rdmd",
        },
        "Clean Wash": {
            "3 Mo Clean Rdmd",
            "CLUB-clean Rdmd",
            "Monthly Cln '21 Rdmd",
            "COMP-CLUB-clean Rdmd",
        },
        "Protect Wash": {
            "3 Mo Protect Rdmd",
            "City Fire Club Rdmd",
            "CLUB-protect Rdmd",
            "COMP-CLUB-prot Rdmd",
            "Mnthly Prot 6mo Rdmd",
            "New Mnthy Prot Rdmd",
            "New Mnthly Prot Rdmd",
        },
        "UShine Wash": {
            "Monthly UShine Rdmd",
            "3 Mo UShine Rdmd",
            "CLUB-ushine Rdmd",
            "COMP-CLUB-ushin Rdmd",
            "Free Week UShine Rdm",
        },
    }
    Website_PKG_Names = {
        # The packages sold on the website differ in name and category than the other standard location packages. 
        # For instance, all website items (including ARM Sellers & Giftcard Sellers) are lumped into one "Website Sold" category
        "Retail": {
            "Express Wash": {
                
            },
            "Clean Wash": {
                "W-1-clean wash",
            },
            "Protect Wash": {
                "W-1-protect wash",
            },
            "UShine Wash": {
                "W-1-UShine wash",
            },
        },

        "Monthly": {
            "Express Wash": {
                "W-New MonthlyExprSld",
            },
            "Clean Wash": {
                "W-Mnthly Cln '21 Sld",
            },
            "Protect Wash": {
                "W-Unl. Prot 9.95 Sld",
                "W-Unl. protect Sld",
            },
            "UShine Wash": {
                "W-MonthlyUShine Sld",
                "W-UShine 9.95 Sld",
            },
        }
    }
    blacklisted_items = [
        # Copied directly from the GSR CSV, hence the duplicate values.
        "Deposits",
        "XPT Cash Add/Remove",
        "XPT Cash Add/Remove",
        "XPT Cash Add/Remove",
        "XPT Cash Add/Remove",
        "XPT Cash Add/Remove",
        "House Acct Related",
        "House Acct Related",
        "House Acct Related",
        "House Acct Related",
        "House Acct Related",
        "House Acct Related",
        "Employee Accounts",
        "Employee Accounts",
        "Cash",
        "XPT Cash Over/Short",
        "XPT Chg Over/Short",
        "Credit Card",
        "Credit Card",
        "Credit Card",
        "Other Tenders",
        "Other Tenders",
        "House Accounts",
        "XPT Balancing",
        "Picture Mismatch",
        "Employees"
        #"Website Sold"
    ]
    blacklisted_sites = [
        # Add sites to this list to EXCLUDE them from the report.
        "Hub Office",
        #"wash*u - Centennial",
        "Zwash*u - Centennial Old",
        #"wash*u - Berwyn",
        #"wash*u - Joliet",
        #"wash*u - Des Plaines",
        #"wash*u - Des Plaine"
    ]
    
    Monthly_Total_Categories = [
        # Member revenue is calculated by adding amounts from these categories.
        "ARM Plans Sold", 
        "ARM Plans Recharged", 
        "Club Plans Sold", 
        "ARM Plans Terminated"
    ]
    ARM_Member_Categories = [
        "ARM Plans Sold", 
        "ARM Plans Recharged", 
        "ARM Plans Terminated"
    ]
    Discount_Categories = [
        "Website Discounts", 
        "Wash Discounts", 
        "Wash LPM Discounts",
        "Prepaid Redeemed"
    ]
    MONTHLY_STATS = {}
    COMBINED_STATS = {}
    COMBINED_SALES = {}
    MONTHLY_SALES = {}

    #--> Utility Functions
    def find_parent(json_obj, target_str, start_parent=None, current_parent=None):
        """
        Searches for a string in a nested JSON object (Python dict) starting from a specified parent,
        and returns the parent key under which the string resides.
        
        Parameters:
        - json_obj (dict): The JSON object to search through.
        - target_str (str): The string to look for.
        - start_parent (str): The parent key from where to start the search.
        - current_parent: The current parent key, used for recursion.
        
        Returns:
        - The parent key under which the string resides, or None if the string is not found.
        """
        if start_parent and current_parent is None:
            json_obj = json_obj.get(start_parent, {})
            
        for key, value in json_obj.items():
            if key == target_str:
                return current_parent
            if isinstance(value, dict):
                result = find_parent(value, target_str, None, key)
                if result:
                    return result
            elif isinstance(value, (list, set)):
                if target_str in value:
                    return key
        return None
    
    #--> Utility Functions (Graphs)
    def createMiniBarGraph(worksheet, START_COL, START_ROW):
        ws = worksheet

        # Create a bar chart
        chart = BarChart()

        # Set the data range for the chart
        # Assuming data for the chart is in cells A4:B7
        values = Reference(ws, min_col=START_COL+1, min_row=START_ROW+3, max_col=START_COL+1, max_row=START_ROW+6)
        labels = Reference(ws, min_col=START_COL, min_row=START_ROW+3, max_row=START_ROW+6)
        chart.add_data(values, titles_from_data=False)
        chart.set_categories(labels)

        # Set the dimensions of the chart
        chart.height = 4.05  # height
        chart.width = 9.5  # width

        # Remove the legend
        chart.legend = None

        # Add data labels
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True

        # Change bar colors (optional)
        colors = ['4682B4']  # Blue
        for i, s in enumerate(chart.series):
            s.graphicalProperties.solidFill = colors[i % len(colors)]
        
        CHART_PLACEMENT_COL = START_COL + 4
        CHART_PLACEMENT_ROW = START_ROW + 2

        # Add the chart to the worksheet
        ws.add_chart(chart, f"{get_column_letter(CHART_PLACEMENT_COL)}{CHART_PLACEMENT_ROW}")
    
    def createPieChart(worksheet, placementCell, dataReference, categoryReference, chartTitle, chartHeight=None, chartWidth=None, valueStyle=None):
        ws = worksheet

        # Create a bar chart
        chart = PieChart()

        # Add the data range for the chart
        chart.add_data(dataReference, titles_from_data=False)
        chart.set_categories(categoryReference)

        # Set the dimensions of the chart
        chart.height = chartHeight if chartHeight != None else 9.3
        chart.width = chartWidth if chartWidth != None else 9.5 


        # Add data labels
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
        chart.dLbls.showSerName = False
        chart.dLbls.showLeaderLines = True


        # Add the chart to the worksheet
        ws.add_chart(chart, placementCell)

        chart.title = chartTitle
    
    def createCorporatePieChart(worksheet, placementCell, dataReference, categoryReference, chartTitle, chartHeight=None, chartWidth=None, valueStyle=None, showPercents=False):
        ws = worksheet

        # Create a bar chart
        chart = PieChart()

        # Add the data range for the chart
        chart.add_data(dataReference, titles_from_data=False)
        chart.set_categories(categoryReference)

        # Set the dimensions of the chart
        chart.height = chartHeight if chartHeight != None else 20
        chart.width = chartWidth if chartWidth != None else 20 


        # Add data labels
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
        chart.dLbls.showSerName = False
        chart.dLbls.showLeaderLines = True
        if showPercents == True:
            chart.dLbls.showPercent = True


        # Add the chart to the worksheet
        ws.add_chart(chart, placementCell)

        chart.title = chartTitle

    def createBarGraph(worksheet, placementCell, dataReference, categoryReference, chartTitle, chartHeight=None, chartWidth=None, valueStyle=None, y_axisTitle=None):
        ws = worksheet

        # Create a bar chart
        chart = BarChart()

        # Add the data range for the chart
        chart.add_data(dataReference, titles_from_data=False)
        chart.set_categories(categoryReference)

        # Set the dimensions of the chart
        chart.height = chartHeight if chartHeight != None else 9.3
        chart.width = chartWidth if chartWidth != None else 19 


        # Add data labels
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
        chart.dLbls.showLeaderLines = True
        chart.legend = None
        chart.y_axis.title = y_axisTitle if y_axisTitle != None else ''
        

        # Add the chart to the worksheet
        ws.add_chart(chart, placementCell)

        chart.title = chartTitle

    def createCorporateBarGraph(worksheet, placementCell, dataReference, categoryReference, chartTitle, chartHeight=None, chartWidth=None, valueStyle=None, y_axisTitle=None, useDataTitle=True):
        ws = worksheet

        # Create a bar chart
        chart = BarChart()

        # Add the data range for the chart
        chart.add_data(dataReference, titles_from_data=useDataTitle)
        chart.set_categories(categoryReference)

        # Set the dimensions of the chart
        chart.height = chartHeight if chartHeight != None else 20
        chart.width = chartWidth if chartWidth != None else 20


        # Add data labels
        chart.dLbls = DataLabelList()
        #chart.dLbls.showVal = True
        #chart.dLbls.showLeaderLines = True
        #chart.legend = None
        chart.y_axis.title = y_axisTitle if y_axisTitle != None else ''

        chart.plot_area.dTable = DataTable()
        chart.plot_area.dTable.showHorzBorder = True
        chart.plot_area.dTable.showVertBorder = True
        chart.plot_area.dTable.showOutline = True
        chart.plot_area.dTable.showKeys = True

        chart.type = "col"
        chart.grouping = "stacked"
        chart.overlap = 100
        

        # Add the chart to the worksheet
        ws.add_chart(chart, placementCell)

        chart.title = chartTitle
    
    def createPieInPieChart(worksheet, placementCell, dataReference, categoryReference, chartTitle, chartHeight=None, chartWidth=None, valueStyle=None, y_axisTitle=None):
        ws = worksheet

        # Create a bar chart
        chart = ProjectedPieChart()

        # Add the data range for the chart
        chart.add_data(dataReference, titles_from_data=False)
        chart.set_categories(categoryReference)

        # Set the dimensions of the chart
        chart.height = chartHeight if chartHeight != None else 9.3
        chart.width = chartWidth if chartWidth != None else 19

        chart.splitPos=4


        # Add data labels
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
        chart.dLbls.showLeaderLines = True
        

        # Add the chart to the worksheet
        ws.add_chart(chart, placementCell)

        chart.title = chartTitle
    
    




    #########################################################################################################################################################################
    ############################################################################## GSR PARSING ##############################################################################
    #########################################################################################################################################################################
    #--> Gather Start/End Dates for the report

    # Read the CSV file to extract the start date and end date
    df = pd.read_csv(gsrFilePath, nrows=2)  # Reading only the first 2 rows

    def extract_dates_from_filepath():
        # Extract the file name from the file path
        filename = fileNameForParser
        global monthsInReport
        # Use regular expression to find dates in the filename
        match = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{4}-\d{2}-\d{2})", filename)
        
        if match:
            start_date_str, end_date_str = match.groups()
            
            # Convert the extracted date strings to datetime objects
            start_date_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date_dt = datetime.strptime(end_date_str, '%Y-%m-%d')

            start_date = start_date_dt.date()
            end_date = end_date_dt.date()

            # Define the date range bounds for 8/14 - 8/23
            lower_bound = date(2023, 8, 14)
            upper_bound = date(2023, 8, 23)

            # Check if the date range 8/14 - 8/23 is completely covered by start_date and end_date
            is_range_covered = start_date <= lower_bound and end_date >= upper_bound
            if is_range_covered: desplainesOverride = True

            # Convert the start and end dates from strings to datetime objects
            # Calculate the difference in days between the two dates
            delta = (end_date_dt- start_date_dt).days

            # Calculate the number of months
            monthsInReport = delta / 30.44  # On average, a month is about 30.44 days

            print('Months in report:', monthsInReport)

            return start_date, end_date
        else:
            return None, None

    # Example usage
    start_date, end_date = extract_dates_from_filepath()
    print(start_date, end_date)
    # Format the date strings
    short_date_str, complete_date_str = str(start_date) + ' - ' + str(end_date), str(start_date) + ' - ' + str(end_date)
        
    #--> Prepare dictionaries for data entry
    df = pd.read_csv(gsrFilePath)
    mask = ~df['Site'].isin(blacklisted_sites)
    df = df[mask]
    for _, row in df.iterrows():
        site = row['Site']
        print("Calculating for", site)
        category = row['Report Category']
        item_name = row['Item Name']
        amount = row['Amount']

        if site not in blacklisted_sites:
            if site not in MONTHLY_STATS:
                MONTHLY_STATS[site] = {"Express Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}, "Clean Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}, "Protect Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}, "UShine Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}}
            if site not in COMBINED_STATS:
                COMBINED_STATS[site] = {"Express Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}, "Clean Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}, "Protect Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}, "UShine Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}}
            if Corporation_Totals_Name not in MONTHLY_STATS:
                MONTHLY_STATS[Corporation_Totals_Name] = {"Express Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}, "Clean Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}, "Protect Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}, "UShine Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0, 'Estimated Member Count': 0}}
            if Corporation_Totals_Name not in COMBINED_STATS:
                COMBINED_STATS[Corporation_Totals_Name] = {"Express Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}, "Clean Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}, "Protect Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}, "UShine Wash":{'Count': 0, 'Price': 0, 'Quantity': 0, 'Amount': 0}}
            
            if site not in MONTHLY_SALES:
                MONTHLY_SALES[site] = {"NET Sales": 0}
            if site not in COMBINED_SALES:
                COMBINED_SALES[site] = {"Discounts": 0, "NET Sales": 0}
            if Corporation_Totals_Name not in MONTHLY_SALES:
                MONTHLY_SALES[Corporation_Totals_Name] = {"NET Sales": 0}
            if Corporation_Totals_Name not in COMBINED_SALES:
                COMBINED_SALES[Corporation_Totals_Name] = {"Discounts": 0, "NET Sales": 0}

    #--> Gather TOTAL sales & discounts details
    #NOTE: RETAIL statistics are determined based off of Combined & Monthly. Spreadsheet formulas will determine the Quantity and Totals as well.
    df = pd.read_csv(gsrFilePath)
    mask = ~df['Site'].isin(blacklisted_sites)
    df = df[mask]
    for _, row in df.iterrows():
        site = row['Site']
        category = row['Report Category']
        item_name = row['Item Name']
        amount = row['Amount']
        
        if site not in blacklisted_sites and item_name is not None and category not in blacklisted_items:

            COMBINED_SALES[site]['NET Sales'] += amount if amount != None else 0
            COMBINED_SALES[Corporation_Totals_Name]['NET Sales'] += amount if amount != None else 0
            if category in Discount_Categories:
                COMBINED_SALES[site]['Discounts'] += amount if amount != None else 0
                COMBINED_SALES[Corporation_Totals_Name]['Discounts'] += amount if amount != None else 0

            # Handle the Query Server
            if site == 'Query Server':
                if find_parent(Website_PKG_Names, item_name, start_parent="Monthly") != None or category in Monthly_Total_Categories:
                    MONTHLY_SALES['Query Server']['NET Sales'] += amount if amount != None else 0
                    MONTHLY_SALES[Corporation_Totals_Name]['NET Sales'] += amount if amount != None else 0
            # Any other locations
            else: 
                if category in Monthly_Total_Categories:
                    MONTHLY_SALES[site]['NET Sales'] += amount if amount != None else 0
                    MONTHLY_SALES[Corporation_Totals_Name]['NET Sales'] += amount if amount != None else 0

            

    #--> Gather Sales Mix
    """
    Calculating Combined, Monthly, and Retail Stats:

    Components:
    - ARM Redeemer: A package leader item triggered when a FastPass Plan customer pays.
    - ARM Plan Packages: Package leader items contained within the ARM Redeemer.
    - Wash Package: Actual wash service (e.g., 'Express', 'Protect').
    - ARM Redeemer Discount: Makes the Wash Package cost $0 for FastPass Plan customers.

    Transaction Flow:
    1. FastPass Plan customer arrives, triggering the ARM Redeemer.
    2. The ARM Redeemer contains ARM Plan Packages, each of which includes:
        1. A Wash Package (e.g., 'Express', 'Protect').
        2. An ARM Redeemer Discount, setting the Wash Package cost to $0.

    Stat Calculations:
    - Combined Stats (Retail + Monthly): Both retail and monthly transactions are recorded under 'Basic Washes' because they both trigger a Wash Package.
    - Monthly Stats: Tracked when an ARM Redeemer is activated (exclusive to monthly FastPass Plan customers).
    - Retail Stats: Obtained by subtracting Monthly Stats from Combined Stats.

    Formula:
    Retail Stats = Combined Stats - Monthly Stats

    Template:
    The template is Excel based and therefore can be set up early to calculate the retail data based on the input from this script for combined and monthly.
    """
    #NOTE: RETAIL statistics are determined based off of Combined & Monthly. Spreadsheet formulas will determine the Quantity and Totals as well.
    df = pd.read_csv(gsrFilePath)
    mask = ~df['Site'].isin(blacklisted_sites)
    df = df[mask]

    MonthlyAmounts = {}
    CombinedAmounts = {}
    AccurateMemberCountAdjuster = 0
    price_sum_dict = {}  # This will hold the sum of prices for each item for each site
    price_count_dict = {}  # This will hold the count of entries for each item for each site


    for _, row in df.iterrows():
        site = row['Site']
        category = row['Report Category']
        item_name = row['Item Name']
        amount = row['Amount']
        count = row['Count']
        price = row['Price']
        quantity = row['Quantity']
        
        if site not in blacklisted_sites and item_name is not None and category not in blacklisted_items:
            
            ## TODO: Add period multiplier 

            if item_name in ['WEB Discontinue ARM', 'Discontinue ARM Plan']:
                AccurateMemberCountAdjuster += quantity if pd.isna(quantity) != True else 0 # Qty is already negative in the GSR, need to add it.

            washPkg = find_parent(ARM_Recharge_Names, item_name) # Member Count = ARM Recharges
            if washPkg != None:
                MONTHLY_STATS[site][washPkg]['Estimated Member Count'] += round(quantity/monthsInReport) if pd.isna(quantity) != True else 0

                MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Estimated Member Count'] += round(quantity/monthsInReport) if pd.isna(quantity) != True else 0
            washPkg = find_parent(Website_PKG_Names, item_name, "Monthly") # Member Count = ARM Recharges + Website Sld
            if washPkg != None:
                MONTHLY_STATS[site][washPkg]['Estimated Member Count'] += round(quantity/monthsInReport) if pd.isna(quantity) != True else 0

                MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Estimated Member Count'] += round(quantity/monthsInReport) if pd.isna(quantity) != True else 0
            washPkg = find_parent(ARM_Sold_Names, item_name) # Member Count = ARM Recharges + Website Sld + ARM Sld
            if washPkg != None:
                MONTHLY_STATS[site][washPkg]['Estimated Member Count'] += round(quantity/monthsInReport) if pd.isna(quantity) != True else 0

                MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Estimated Member Count'] += round(quantity/monthsInReport) if pd.isna(quantity) != True else 0
            washPkg = find_parent(ARM_Termination_Names, item_name) # Member Count = ARM Recharges + Website Sld + ARM Sld - ARM Terminations
            if washPkg != None:
                MONTHLY_STATS[site][washPkg]['Estimated Member Count'] -= round(quantity/monthsInReport) if pd.isna(quantity) != True else 0
                MONTHLY_STATS[site][washPkg]['Estimated Member Count'] = abs(MONTHLY_STATS[site][washPkg]['Estimated Member Count'])
                MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Estimated Member Count'] -= round(quantity/monthsInReport) if pd.isna(quantity) != True else 0
                MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Estimated Member Count'] = abs(MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Estimated Member Count'])

            # Handle the Query Server
            if site == 'Query Server':
                # We're only going to show the amount of passes sold rather than the redemption for the
                # Query Server because you cannot redeem passes online.
                washPkg = find_parent(Website_PKG_Names, item_name, "Monthly")
                if washPkg != None:
                    MONTHLY_STATS['Query Server'][washPkg]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    MONTHLY_STATS['Query Server'][washPkg]['Count'] += count if pd.isna(count) != True else 0
                    MONTHLY_STATS['Query Server'][washPkg]['Amount'] += amount if pd.isna(amount) != True else 0
                    MONTHLY_STATS['Query Server'][washPkg]['Price'] += price if pd.isna(price) != True else 0

                    MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Count'] += count if pd.isna(count) != True else 0
                    MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Amount'] += amount if pd.isna(amount) != True else 0
                    if not washPkg in MonthlyAmounts:
                        MonthlyAmounts[washPkg] = []
                    MonthlyAmounts[washPkg].append(price if pd.isna(price) != True else 0)
                
                washPkg = find_parent(Website_PKG_Names, item_name, "Retail")
                if washPkg != None:
                    COMBINED_STATS['Query Server'][washPkg]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    COMBINED_STATS['Query Server'][washPkg]['Count'] += count if pd.isna(count) != True else 0
                    COMBINED_STATS['Query Server'][washPkg]['Amount'] += amount if pd.isna(amount) != True else 0
                    COMBINED_STATS['Query Server'][washPkg]['Price'] += price if pd.isna(price) != True else 0

                    COMBINED_STATS[Corporation_Totals_Name][washPkg]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    COMBINED_STATS[Corporation_Totals_Name][washPkg]['Count'] += count if pd.isna(count) != True else 0
                    COMBINED_STATS[Corporation_Totals_Name][washPkg]['Amount'] += amount if pd.isna(amount) != True else 0
                    if not washPkg in MonthlyAmounts:
                        CombinedAmounts[washPkg] = []
                    CombinedAmounts[washPkg].append(price if pd.isna(price) != True else 0)

            # Any other locations
            else: 
                washPkg = find_parent(ARM_PKG_Names, item_name)
                # Calculate MONTHLY Stats using the redemption items
                if category in ["ARM Plans Redeemed", "Club Plans Redeemed"] and washPkg is not None:
                    MONTHLY_STATS[site][washPkg]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    MONTHLY_STATS[site][washPkg]['Count'] += count if pd.isna(count) != True else 0
                    MONTHLY_STATS[site][washPkg]['Amount'] += amount if pd.isna(amount) != True else 0
                    MONTHLY_STATS[site][washPkg]['Price'] += abs(price) if pd.isna(price) != True else 0

                    MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Count'] += count if pd.isna(count) != True else 0
                    MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Amount'] += amount if pd.isna(amount) != True else 0
                    if not washPkg in MonthlyAmounts:
                        MonthlyAmounts[washPkg] = []
                    MonthlyAmounts[washPkg].append(price if pd.isna(price) != True else 0) 
                elif category in ['Basic Washes'] and item_name is not None and item_name in COMBINED_STATS[site]:
                    COMBINED_STATS[site][item_name]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    COMBINED_STATS[site][item_name]['Count'] += count if pd.isna(count) != True else 0
                    COMBINED_STATS[site][item_name]['Amount'] += amount if pd.isna(amount) != True else 0
                    COMBINED_STATS[site][item_name]['Price'] += amount/quantity if pd.isna(quantity) != True else 0

                    COMBINED_STATS[Corporation_Totals_Name][item_name]['Quantity'] += quantity if pd.isna(quantity) != True else 0
                    COMBINED_STATS[Corporation_Totals_Name][item_name]['Count'] += count if pd.isna(count) != True else 0
                    COMBINED_STATS[Corporation_Totals_Name][item_name]['Amount'] += amount if pd.isna(amount) != True else 0
                    if not washPkg in CombinedAmounts:
                        CombinedAmounts[item_name] = []
                    CombinedAmounts[item_name].append(price if pd.isna(price) != True else 0)
                    '''if not site in price_sum_dict:
                        price_sum_dict[site] = {}
                    if not item_name in price_sum_dict[site]:
                        price_sum_dict[site][item_name] = []
                    price_sum_dict[site][item_name].append(price if pd.isna(price) != True else 0)'''

    #print(json.dumps(price_sum_dict, indent=4))
    for washPkg in ARM_Sold_Names:
        COMBINED_STATS[Corporation_Totals_Name][washPkg]['Price'] = COMBINED_STATS[Corporation_Totals_Name][washPkg]['Amount']/COMBINED_STATS[Corporation_Totals_Name][washPkg]['Quantity']
        MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Price'] = MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Amount']/MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Quantity']
        #COMBINED_STATS[Corporation_Totals_Name][washPkg]['Price'] = sum(CombinedAmounts[washPkg])/len(CombinedAmounts[washPkg])
        #MONTHLY_STATS[Corporation_Totals_Name][washPkg]['Price'] = sum(MonthlyAmounts[washPkg])/len(MonthlyAmounts[washPkg])
    
    #########################################################################################################################################################################
    ############################################################################## T&C PARSING ##############################################################################
    #########################################################################################################################################################################
    '''
    ChurnRates = {}
    df = pd.read_csv(trendsFilePath)
    for _, row in df.iterrows():
        date = row['Date']
        plan = row['Plan']
        site = row['Site']
        washes_per_day = row['Washes per day']
        upsells_dollars = row['Upsells (Dollars)']
        upsells = row['Upsells']
        plans_sold_dollars = row['Plans Sold (Dollars)']
        plans_sold = row['Plans Sold']
        recharges_dollars = row['Recharges (Dollars)']
        recharges = row['Recharges']
        plans_transferred_in = row['Plans Transferred In']
        plans_transferred_out = row['Plans Transferred Out']
        plans_transferred_in_dollars = row['Plans Transferred In (Dollars)']
        plans_transferred_out_dollars = row['Plans Transferred Out (Dollars)']
        suspended = row['Suspended']
        suspended_dollars = row['Suspended (Dollars)']
        resumed = row['Resumed']
        resumed_dollars = row['Resumed (Dollars)']
        plans_terminated = row['Plans Terminated']
        plans_terminated_dollars = row['Plans Terminated (Dollars)']
        plans_discontinued = row['Plans Discontinued']
        cc_declining_expired = row['CC Declining + CC Expired']
        expired_members = row['Expired Members']
        members_per_day = row['Members per day']
        revenue = row['Revenue']
        pass_wash_percentage = row['Pass Wash Percentage']
        churn_rate = row['Churn Rate']

    Need a formula to determine the churn rate for the corporatiion based off of the churn rate from the individual stores. 

    '''
    #########################################################################################################################################################################
    ############################################################################## XCL LOADING ##############################################################################
    #########################################################################################################################################################################
    #--> Utility Functions to copy a range of cells from one location to another, including their styles and formulas
    def adjust_cell_reference(match, row_shift, col_shift):
        cell_ref = match.group(0)
        col_ref, row_ref = re.match(r"([A-Z]+)([0-9]+)", cell_ref).groups()
        new_col = get_column_letter(column_index_from_string(col_ref) + col_shift)
        new_row = str(int(row_ref) + row_shift)
        return new_col + new_row

    def adjust_formula(formula, row_shift, col_shift):
        pattern = r"(\$?[A-Z]+\$?[0-9]+)"
        return re.sub(pattern, lambda match: adjust_cell_reference(match, row_shift, col_shift), formula)

    def copy_cells(ws, start_row, end_row, start_col, end_col, target_start_row, target_start_col, target_ws=None):
        row_shift = target_start_row - start_row
        col_shift = target_start_col - start_col
        
        # If no target worksheet is given, default to the current worksheet
        if target_ws is None:
            target_ws = ws

        for i, row in enumerate(ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col)):
            for j, cell in enumerate(row):
                target_cell = target_ws.cell(row=target_start_row + i, column=target_start_col + j)
                
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    # Adjust the formula
                    adjusted_formula = adjust_formula(cell.value[1:], row_shift, col_shift)
                    target_cell.value = f"={adjusted_formula}"
                else:
                    target_cell.value = cell.value

                # Copy cell style if it has one
                if cell.has_style:
                    target_cell._style = copy(cell._style)
                
                # Copy font, border, fill, number format, etc.
                target_cell.font = copy(cell.font)
                target_cell.border = copy(cell.border)
                target_cell.fill = copy(cell.fill)
                target_cell.number_format = copy(cell.number_format)
                target_cell.protection = copy(cell.protection)
                target_cell.alignment = copy(cell.alignment)


    #--> Populate Worksheet Function to fill in each location block
    def populate_salesmix_worksheet(worksheet, start_row, start_col, location):
        global monthsInReport

        #--> Assign location name as the new block name
        if location != 'Query Server':
            worksheet.cell(row=start_row, column=start_col, value=location)
        else:
            worksheet.cell(row=start_row, column=start_col, value="E-Commerce Website")
        
        worksheet.cell(row=start_row, column=start_col+5, value=monthsInReport)
        
        #--> BLOCK: Retail Stats
        #### Need to fill in the prices (average price) because the Quantity will be determined by the excel formulas.
        for i, (washPkg, pkgProps) in enumerate(COMBINED_STATS[location].items()):
            worksheet.cell(row=start_row+3+i, column=start_col+3, value=pkgProps['Price'])
        
        #--> BLOCK: Monthly Stats
        #### Need to fill in the prices (average price) because the Quantity will be determined by the excel formulas.
        for i, (washPkg, pkgProps) in enumerate(MONTHLY_STATS[location].items()):
            worksheet.cell(row=start_row+12+i, column=start_col+1, value=pkgProps['Quantity'])
            if location != Corporation_Totals_Name:
                worksheet.cell(row=start_row+12+i, column=start_col+5, value=pkgProps['Estimated Member Count'])#-AccurateMemberCountAdjuster/len(WashPkgPrices)/monthsInReport) # Not sure if maybe adding some of the discontinuations makes it more accurate? 
            else:
                worksheet.cell(row=start_row+12+i, column=start_col+5, value=pkgProps['Estimated Member Count'])#-AccurateMemberCountAdjuster/len(WashPkgPrices)/monthsInReport)
        worksheet.cell(row=start_row+17, column=start_col+1, value=MONTHLY_SALES[location]['NET Sales'])

        #--> BLOCK: Combined Stats
        #### Need to fill in the prices (average price) because the Quantity will be determined by the excel formulas.
        for i, (washPkg, pkgProps) in enumerate(COMBINED_STATS[location].items()):
            if location != 'Query Server':
                worksheet.cell(row=start_row+21+i, column=start_col+1, value=pkgProps['Quantity'])
                worksheet.cell(row=start_row+21+i, column=start_col+3, value=pkgProps['Price'])
            else:
                worksheet.cell(row=start_row+3+i, column=start_col+1, value=pkgProps['Quantity'])
                worksheet.cell(row=start_row+3+i, column=start_col+3, value=pkgProps['Price'])
                worksheet.cell(row=start_row+21+i, column=start_col+1, value=pkgProps['Quantity']+MONTHLY_STATS['Query Server'][washPkg]['Quantity'])
                worksheet.cell(row=start_row+21+i, column=start_col+3, value=pkgProps['Price'])
        
        worksheet.cell(row=start_row+26, column=start_col+1, value=COMBINED_SALES[location]['NET Sales'])
        worksheet.cell(row=start_row+26, column=start_col+3, value=COMBINED_SALES[location]['Discounts'])
        
        #--> Query Server Modifications:
        if location == "Query Server":
            worksheet.cell(row=start_row+11, column=start_col+1, value="New Passes Sold")
            worksheet.cell(row=start_row+16, column=start_col, value="Total ARM Plans Sold")
            worksheet.cell(row=start_row+20, column=start_col+1, value="Items Sold")
            worksheet.cell(row=start_row+25, column=start_col, value="Total Items Sold for Period")
        
        
        createMiniBarGraph(worksheet, start_col, start_row)
        createMiniBarGraph(worksheet, start_col, start_row+18)

        ## GRAPH CREATION!
        #--> Set up the sheet
        if location not in [Corporation_Totals_Name, 'Query Server']:
            ws = wb['Sales Mix by Location (Visual)']
            visualWorksheet = wb.copy_worksheet(ws)
            visualWorksheet.title = f'Sales Mix Charts ({location.strip("wash*u - ")})'

            copy_cells(wb['Sales Mix by Location'], start_row, start_row+28, start_col, start_col+5, 1, 1, visualWorksheet)

            # Generate Graphs
            # Retail Sales Mix (Qty)
            createPieChart(visualWorksheet, "H2", Reference(visualWorksheet, min_col=3, min_row=4, max_col=3, max_row=7), Reference(visualWorksheet, min_col=1, min_row=4, max_col=1, max_row=7), "Retail Sales Mix (Qty)")
            # Monthly Membership Mix (Qty)
            createPieChart(visualWorksheet, "J2", Reference(visualWorksheet, min_col=3, min_row=13, max_col=3, max_row=16), Reference(visualWorksheet, min_col=1, min_row=13, max_col=1, max_row=16), "Monthly Membership Mix (Qty)")
            # Combined Wash Mix (Qty)
            createPieChart(visualWorksheet, "L2", Reference(visualWorksheet, min_col=3, min_row=22, max_col=3, max_row=25), Reference(visualWorksheet, min_col=1, min_row=22, max_col=1, max_row=25), "Combined Wash Mix (Qty)")
            
            # Net Revenue Breakdown (%)
            createPieChart(visualWorksheet, "H19", Reference(visualWorksheet, min_col=3, min_row=41, max_col=3, max_row=46), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=46), "Net Revenue Breakdown (%)", 13.95, 14.25)
            # Net Revenue Breakdown ($)
            createPieChart(visualWorksheet, "K19", Reference(visualWorksheet, min_col=2, min_row=41, max_col=2, max_row=46), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=46), "Net Revenue Breakdown ($)", 13.95, 14.25)
            # Retail Revenue Breakdown (%)
            createPieChart(visualWorksheet, "H43", Reference(visualWorksheet, min_col=4, min_row=41, max_col=4, max_row=45), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=45), "Retail Revenue Breakdown (%)", 13.95, 14.25)
            # Retail Revenue Breakdown ($)
            createPieChart(visualWorksheet, "K43", Reference(visualWorksheet, min_col=2, min_row=41, max_col=2, max_row=45), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=45), "Retail Revenue Breakdown ($)", 13.95, 14.25)
            
            # Monthly Membership Utilization
            createBarGraph(visualWorksheet, "O2", Reference(visualWorksheet, min_col=5, min_row=13, max_col=5, max_row=16), Reference(visualWorksheet, min_col=1, min_row=13, max_col=1, max_row=16), "Avg. Monthly Pass Utilization", y_axisTitle="Average # of Washes over Period")
            # Monthly Membership Count
            createBarGraph(visualWorksheet, "O19", Reference(visualWorksheet, min_col=6, min_row=13, max_col=6, max_row=16), Reference(visualWorksheet, min_col=1, min_row=13, max_col=1, max_row=16), "Total Number of Monthly Members", y_axisTitle="Number of Members")

            # Traffic Pie in pie Chart (Retail + Monthly, then monthly broken into its own pie chart)
            createPieInPieChart(visualWorksheet, "O36", Reference(visualWorksheet, min_col=2, min_row=52, max_col=2, max_row=59), Reference(visualWorksheet, min_col=1, min_row=52, max_col=1, max_row=59), "Retail & Monthly Wash Mix", y_axisTitle="Number of Members")

            createMiniBarGraph(visualWorksheet, 1, 3-2)
            createMiniBarGraph(visualWorksheet, 1, 21-2)
        elif location == Corporation_Totals_Name:
            # Due to the nature of the corporate tab, data is built in this function and then churned into graphs, rather than allowing the spreadsheet to handle it.
            # This is because the number of locations is dynamic and changing therefore we cannot hard-code the corporate graph data functions like we do with the others.

            # Copy the template to the workbook
            ws = wb['Sales Mix by Location (Visual)']
            visualWorksheet = wb.copy_worksheet(ws)
            visualWorksheet.title = 'Sales Mix Charts (Corporation)'

            copy_cells(wb['Sales Mix by Location'], start_row, start_row+28, start_col, start_col+5, 1, 1, visualWorksheet)

            # Generate Graphs (Same graphs as designed from the locations)
            # Retail Sales Mix (Qty)
            createPieChart(visualWorksheet, "H2", Reference(visualWorksheet, min_col=3, min_row=4, max_col=3, max_row=7), Reference(visualWorksheet, min_col=1, min_row=4, max_col=1, max_row=7), "Retail Sales Mix (Qty)")
            # Monthly Membership Mix (Qty)
            createPieChart(visualWorksheet, "J2", Reference(visualWorksheet, min_col=3, min_row=13, max_col=3, max_row=16), Reference(visualWorksheet, min_col=1, min_row=13, max_col=1, max_row=16), "Monthly Membership Mix (Qty)")
            # Combined Wash Mix (Qty)
            createPieChart(visualWorksheet, "L2", Reference(visualWorksheet, min_col=3, min_row=22, max_col=3, max_row=25), Reference(visualWorksheet, min_col=1, min_row=22, max_col=1, max_row=25), "Combined Wash Mix (Qty)")
            
            # Net Revenue Breakdown (%)
            createPieChart(visualWorksheet, "H19", Reference(visualWorksheet, min_col=3, min_row=41, max_col=3, max_row=46), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=46), "Net Revenue Breakdown (%)", 13.95, 14.25)
            # Net Revenue Breakdown ($)
            createPieChart(visualWorksheet, "K19", Reference(visualWorksheet, min_col=2, min_row=41, max_col=2, max_row=46), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=46), "Net Revenue Breakdown ($)", 13.95, 14.25)
            # Retail Revenue Breakdown (%)
            createPieChart(visualWorksheet, "H43", Reference(visualWorksheet, min_col=4, min_row=41, max_col=4, max_row=45), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=45), "Retail Revenue Breakdown (%)", 13.95, 14.25)
            # Retail Revenue Breakdown ($)
            createPieChart(visualWorksheet, "K43", Reference(visualWorksheet, min_col=2, min_row=41, max_col=2, max_row=45), Reference(visualWorksheet, min_col=1, min_row=41, max_col=1, max_row=45), "Retail Revenue Breakdown ($)", 13.95, 14.25)
            
            # Monthly Membership Utilization
            createBarGraph(visualWorksheet, "O36", Reference(visualWorksheet, min_col=5, min_row=13, max_col=5, max_row=16), Reference(visualWorksheet, min_col=1, min_row=13, max_col=1, max_row=16), "Avg. Monthly Pass Utilization", y_axisTitle="Average # of Washes")
            # Monthly Membership Count
            createBarGraph(visualWorksheet, "O19", Reference(visualWorksheet, min_col=6, min_row=13, max_col=6, max_row=16), Reference(visualWorksheet, min_col=1, min_row=13, max_col=1, max_row=16), "Total Num. of Monthly Members", y_axisTitle="Number of Members")

            # Traffic Pie in pie Chart (Retail + Monthly, then monthly broken into its own pie chart)
            createPieInPieChart(visualWorksheet, "O2", Reference(visualWorksheet, min_col=2, min_row=52, max_col=2, max_row=59), Reference(visualWorksheet, min_col=1, min_row=52, max_col=1, max_row=59), "Retail & Monthly Wash Mix")

            createMiniBarGraph(visualWorksheet, 1, 3-2)
            createMiniBarGraph(visualWorksheet, 1, 21-2)

            # Generate Competitive Graphs (These will compare sites against other sites on basic revenue, traffic, discounts, etc.)
            # Create Data Tables (will be used to input the data into the visuals sheet for the corporation, tracking the cell inputs with python we can dynamically generate & gather graph data.)
            retail_quantity_mix_data = {}

            for location in COMBINED_STATS:
                if location not in [Corporation_Totals_Name, 'Query Server']:
                    for washPkg in COMBINED_STATS[location]:
                        if location not in retail_quantity_mix_data: 
                            retail_quantity_mix_data[location] = {}
                        retail_quantity_mix_data[location].update({washPkg: COMBINED_STATS[location][washPkg]['Quantity'] - MONTHLY_STATS[location][washPkg]['Quantity']})
            

            retail_revenue_mix_data = {}
            total_retail_revenue_mix_data = {}

            for location in COMBINED_STATS:
                if location not in [Corporation_Totals_Name, 'Query Server']:
                    total_retail_revenue_mix_data[location] = COMBINED_SALES[location]['NET Sales'] - MONTHLY_SALES[location]['NET Sales']
                    for washPkg in COMBINED_STATS[location]:
                        if location not in retail_revenue_mix_data: retail_revenue_mix_data[location] = {}
                        retail_revenue_mix_data[location].update({washPkg: (COMBINED_STATS[location][washPkg]['Quantity'] - MONTHLY_STATS[location][washPkg]['Quantity'])*COMBINED_STATS[location][washPkg]['Price']})#WashPkgPrices[washPkg.lower()]})
                    retail_revenue_mix_data[location].update({'Discounts': COMBINED_SALES[location]['Discounts']})

            monthly_quantity_mix_data = {}

            for location in MONTHLY_STATS:
                if location not in [Corporation_Totals_Name, 'Query Server']:
                    for washPkg in COMBINED_STATS[location]:
                        if location not in monthly_quantity_mix_data: 
                            monthly_quantity_mix_data[location] = {}
                        monthly_quantity_mix_data[location].update({washPkg: MONTHLY_STATS[location][washPkg]['Quantity']})
            

            monthly_revenue_mix_data = {}
            total_monthly_revenue_mix_data = {}

            for location in MONTHLY_STATS:
                if location not in [Corporation_Totals_Name, 'Query Server']:
                    total_monthly_revenue_mix_data[location] = MONTHLY_SALES[location]['NET Sales']
                    for washPkg in COMBINED_STATS[location]:
                        if location not in monthly_revenue_mix_data: monthly_revenue_mix_data[location] = {}
                        monthly_revenue_mix_data[location].update({washPkg: MONTHLY_STATS[location][washPkg]['Amount']})
            

            combined_quantity_mix_data = {}

            for location in COMBINED_STATS:
                if location not in [Corporation_Totals_Name, 'Query Server']:
                    for washPkg in COMBINED_STATS[location]:
                        if location not in combined_quantity_mix_data: 
                            combined_quantity_mix_data[location] = {}
                        combined_quantity_mix_data[location].update({washPkg: COMBINED_STATS[location][washPkg]['Quantity']})
            

            combined_revenue_mix_data = {}
            total_combined_revenue_mix_data = {}

            for location in COMBINED_STATS:
                if location not in [Corporation_Totals_Name, 'Query Server']:
                    total_combined_revenue_mix_data[location] = COMBINED_SALES[location]['NET Sales'] - COMBINED_SALES[location]['Discounts']
                    for washPkg in COMBINED_STATS[location]:
                        if location not in combined_revenue_mix_data: combined_revenue_mix_data[location] = {}
                        combined_revenue_mix_data[location].update({washPkg: COMBINED_STATS[location][washPkg]['Amount']})
            
            start_row = 75
            start_col = 1
            worksheet = wb[f'Sales Mix Charts (Corporation)']

            accounting_style_2 = NamedStyle(
                name='accounting_style_2', 
                number_format='$#,##0.00',  # or you could use 'Accounting'
                alignment=Alignment(horizontal='right')
            )

            '''# Loop through each location and its corresponding wash packages
            for location, washPackages in retail_quantity_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row, column=start_col+1, value=location.strip("wash*u - "))
                
                # Increment the row to start adding wash packages for this location
                current_row = start_row + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    if start_col == 1:
                        worksheet.cell(row=current_row, column=start_col, value=washPackage)
                    worksheet.cell(row=current_row, column=start_col + 1, value=value)
                    current_row += 1  # Move down a row for the next wash package
                
                # Move to the next column to start adding the next location
                start_col += 1  # Shift 2 columns to the right for the next location'''
            # Loop through each location and its corresponding wash packages
            for location, washPackages in retail_quantity_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    worksheet.cell(row=75, column=current_col, value=washPackage)
                    worksheet.cell(row=start_row + 1, column=current_col, value=value)
                    current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location

            createCorporateBarGraph(worksheet, "AA2", Reference(worksheet, min_col=2, min_row=75, max_col=5, max_row=start_row), Reference(worksheet, min_col=1, min_row=76, max_col=1, max_row=start_row), "Retail Wash Sales by Site (QTY)", y_axisTitle="# Washes Sold")

            
            start_row += 1  # Shift 2 rows down for the next location
            washLabelRow = start_row

            for location, washPackages in retail_revenue_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    worksheet.cell(row=washLabelRow, column=current_col, value=washPackage)
                    worksheet.cell(row=start_row + 1, column=current_col, value=value).style = accounting_style_2
                    current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            
            #createCorporateBarGraph(visualWorksheet, "AM2", Reference(visualWorksheet, min_col=2, min_row=washLabelRow, max_col=6, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Retail Revenue Breakdown ($)", useDataTitle=True)

            start_row += 1  # Shift 2 rows down for the next location
            washLabelRow = start_row
            
            for location in total_retail_revenue_mix_data:
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                worksheet.cell(row=start_row + 1, column=current_col, value=total_retail_revenue_mix_data[location]).style = accounting_style_2
                current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            
            createCorporatePieChart(visualWorksheet, "AM2", Reference(visualWorksheet, min_col=2, min_row=washLabelRow+1, max_col=5, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Retail Revenue Breakdown (%)", showPercents=True)


            start_row += 3  # Shift 2 rows down for the next location
            washLabelRow = start_row
# Loop through each location and its corresponding wash packages
            for location, washPackages in monthly_quantity_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    worksheet.cell(row=washLabelRow, column=current_col, value=washPackage)
                    worksheet.cell(row=start_row + 1, column=current_col, value=value)
                    current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            createCorporateBarGraph(worksheet, "AA37", Reference(visualWorksheet, min_col=2, min_row=washLabelRow, max_col=5, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Monthly Wash Rdmds by Site (QTY)", y_axisTitle="# Washes Rdmd")

            
            start_row += 1  # Shift 2 rows down for the next location
            washLabelRow = start_row

            for location, washPackages in monthly_revenue_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    worksheet.cell(row=washLabelRow, column=current_col, value=washPackage)
                    worksheet.cell(row=start_row + 1, column=current_col, value=value).style = accounting_style_2
                    current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            
            #createCorporateBarGraph(visualWorksheet, "AM37", Reference(visualWorksheet, min_col=2, min_row=washLabelRow, max_col=6, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Monthly Revenue Breakdown by Site ($)", useDataTitle=True)

            start_row += 1  # Shift 2 rows down for the next location
            washLabelRow = start_row
            
            for location in total_monthly_revenue_mix_data:
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                worksheet.cell(row=start_row + 1, column=current_col, value=total_monthly_revenue_mix_data[location]).style = accounting_style_2
                current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            
            createCorporatePieChart(visualWorksheet, "AM37", Reference(visualWorksheet, min_col=2, min_row=washLabelRow+1, max_col=5, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Monthly Revenue Breakdown (%)", showPercents=True)




# Loop through each location and its corresponding wash packages
            start_row += 3  # Shift 2 rows down for the next location
            washLabelRow = start_row
            for location, washPackages in combined_quantity_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    worksheet.cell(row=washLabelRow, column=current_col, value=washPackage)
                    worksheet.cell(row=start_row + 1, column=current_col, value=value)
                    current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location

            createCorporateBarGraph(worksheet, "AA74", Reference(visualWorksheet, min_col=2, min_row=washLabelRow, max_col=5, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Combined Wash Sales by Site (QTY)", y_axisTitle="# Washes Sold")

            
            start_row += 1  # Shift 2 rows down for the next location
            washLabelRow = start_row

            for location, washPackages in combined_revenue_mix_data.items():
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                # Loop through each wash package and add it to the sheet
                for washPackage, value in washPackages.items():
                    worksheet.cell(row=washLabelRow, column=current_col, value=washPackage)
                    worksheet.cell(row=start_row + 1, column=current_col, value=value).style = accounting_style_2
                    current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            
            #createCorporateBarGraph(visualWorksheet, "AM2", Reference(visualWorksheet, min_col=2, min_row=washLabelRow, max_col=6, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Retail Revenue Breakdown ($)", useDataTitle=True)

            start_row += 1  # Shift 2 rows down for the next location
            washLabelRow = start_row
            
            for location in total_combined_revenue_mix_data:
                # Write location name to the sheet
                worksheet.cell(row=start_row + 1, column=start_col, value=location.strip("wash*u - "))
                
                
                # Increment the column to start adding wash packages for this location
                current_col = start_col + 1
                
                worksheet.cell(row=start_row + 1, column=current_col, value=total_combined_revenue_mix_data[location]).style = accounting_style_2
                current_col += 1  # Move to the next column for the next wash package
                
                # Move to the next row to start adding the next location
                start_row += 1  # Shift 2 rows down for the next location
            
            createCorporatePieChart(visualWorksheet, "AM74", Reference(visualWorksheet, min_col=2, min_row=washLabelRow+1, max_col=5, max_row=start_row), Reference(visualWorksheet, min_col=1, min_row=washLabelRow+1, max_col=1, max_row=start_row), "Combined Revenue Breakdown ($)", showPercents=True)




    #--> Modify the template and populate all blocks with site data
    wb = openpyxl.load_workbook(templateFilePath)
    ws = wb['Sales Mix by Location']

    start_col = 1
    #--> Assign a report date to the report.
    ws.cell(row=1, column=2, value=complete_date_str)
    locations = list(MONTHLY_STATS.keys())
    locations.remove("Corporation Totals")
    if "Query Server" in locations:
        locations.remove("Query Server")
    for location in locations:
        print(f"Creating {location} Stats Block...")
        copy_cells(ws, 3, 30, 1, 6, 3, start_col) #copy_cells(ws, 1, 28, 1, 6, 1, start_col)
        populate_salesmix_worksheet(ws, 3, start_col, location)
        start_col += 7
    
    copy_cells(ws, 3, 30, 1, 7, 32, 1)
    populate_salesmix_worksheet(ws, 32, 1, Corporation_Totals_Name)

    if not "Query Server" in blacklisted_sites:
        print("Creating Query Server Stats Block...")
        copy_cells(ws, 3, 30, 1, 7, 32, 8)
        populate_salesmix_worksheet(ws, 32, 8, "Query Server")

    
        
    #--> Discounts Worksheet Stuff
    # Define styles
    int_style = NamedStyle(name='integer_style', number_format='#,##0')
    accounting_style = NamedStyle(
        name='accounting_style', 
        number_format='#,##0.00',  # or you could use 'Accounting'
        alignment=Alignment(horizontal='right')
    )
    percent_style = NamedStyle(name='percent_style', number_format='0.00%')

    # Modify the populate_discountmix_worksheet function to include formatting
    def populate_discountmix_worksheet(worksheet, start_row, start_col, location, discount_data):
        worksheet.cell(row=start_row, column=start_col, value=location)
        
        # Adding headers
        headers = ["Discount Name", "Quantity Used", "Price*", "Total Discount Dollars", "Discount Distribution (Qty)", "Discount Distribution ($)"]
        for j, header in enumerate(headers):
            worksheet.cell(row=start_row + 2, column=start_col + j, value=header)
            
        total_quantity = 0
        total_amount = 0

        for i, (discount_name, discount_info) in enumerate(discount_data.items()):
            total_quantity += discount_info['Quantity'] if discount_info['Quantity'] is not None else 0
            total_amount += discount_info['Amount'] if discount_info['Amount'] is not None else 0
        
        for i, (discount_name, discount_info) in enumerate(discount_data.items()):
            if discount_name is not None:

                # Add discount name
                worksheet.cell(row=start_row+3+i, column=start_col, value=discount_name)
                
                # Add and format Quantity
                cell_quantity = worksheet.cell(row=start_row+3+i, column=start_col+1, value=discount_info['Quantity'])
                cell_quantity.style = int_style
                
                # Add and format Price
                cell_price = worksheet.cell(row=start_row+3+i, column=start_col+2, value=discount_info['Price'])
                cell_price.style = accounting_style
                
                # Add and format Amount
                cell_amount = worksheet.cell(row=start_row+3+i, column=start_col+3, value=discount_info['Amount'])
                cell_amount.style = accounting_style

                # Add and format Distribution by Quantity
                distribution_value = discount_info['Quantity'] / total_quantity if total_quantity != 0 else 0
                cell_amount = worksheet.cell(row=start_row+3+i, column=start_col+4, value=distribution_value)
                cell_amount.style = percent_style  # Apply the percentage style

                # Add and format Distribution by Amount
                distribution_value = discount_info['Amount'] / total_amount if total_amount != 0 else 0
                cell_amount = worksheet.cell(row=start_row+3+i, column=start_col+5, value=distribution_value)
                cell_amount.style = percent_style  # Apply the percentage style
                
                
        '''
        # Add and format "Total"
        total_label_cell = worksheet.cell(row=start_row+4+i, column=start_col, value="Total")
        total_label_cell.font = Font(bold=True)

        # Add and format Total Quantity
        cell_total_quantity = worksheet.cell(row=start_row+4+i, column=start_col+1, value=total_quantity)
        cell_total_quantity.style = int_style  # If you have defined int_style elsewhere
        cell_total_quantity.font = Font(bold=True)

        # Add and format Total Amount
        cell_total_amount = worksheet.cell(row=start_row+4+i, column=start_col+3, value="${:.2f}".format(total_amount))
        cell_total_amount.style = accounting_style  # If you have defined accounting_style elsewhere
        cell_total_amount.font = Font(bold=True)
        '''
        
        # Add a Price disclaimer explaining that some prices are averages
        if location == Corporation_Totals_Name:
            disclaim1 = worksheet.cell(row=start_row+4+i+2, column=start_col, value="NOTES")
            disclaim2 = worksheet.cell(row=start_row+4+i+3, column=start_col, value="Price*: Some 'Open Price' items such as 'Misc. Wash Discount' show the average amount that was paid out over the period under that discount item.")
            disclaim3 = worksheet.cell(row=start_row+4+i+4, column=start_col, value="Discounts that are not applied (are missing on the GSR for the location while parsing; thus quantity is 0) will not be shown on the report for that location.")
            disclaim4 = worksheet.cell(row=start_row+4+i+5, column=start_col, value="Due to API limitations with Excel, a proper Total Row for the table cannot be 'enabled' by default. Therefore, it is recommended that you (the end user) enable the total row by clicking anywhere in the table, click table design in the ribbon bar and then check 'Total Row' in the Table Style Options section.")
            disclaim1.font = Font(bold=True)
            disclaim2.font = Font(bold=True)
            disclaim3.font = Font(bold=True)
            disclaim4.font = Font(bold=True)

        tableName = re.sub(r'\W+', '', location.replace(" ", ""))
        # Adding table and table banding
        tab = Table(displayName=tableName, ref=f"{worksheet.cell(row=start_row + 2, column=start_col).coordinate}:{worksheet.cell(row=start_row + 4 + i, column=start_col + 5).coordinate}")
        style = TableStyleInfo(name="TableStyleMedium23", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        worksheet.add_table(tab)

    # Read the input CSV to a DataFrame
    df = pd.read_csv(gsrFilePath)
    mask = ~df['Site'].isin(blacklisted_sites)
    df = df[mask]

    # Aggregate the data
    discount_data = {}
    miscDisctAverages = {}
    discountTotals = {}
    discount_data[Corporation_Totals_Name] = {}
    for _, row in df.iterrows():
        site = row['Site']
        category = row['Report Category']
        discount_name = row['Item Name']
        quantity = row['Quantity']
        price = row['Price']
        amount = row['Amount']

        if (category in Discount_Categories and 
            site not in blacklisted_sites and 
            pd.notna(discount_name) and 
            pd.notna(quantity) and 
            pd.notna(price) and 
            pd.notna(amount)):
            
            if site not in discount_data:
                discount_data[site] = {}
                
            if discount_name not in discount_data[site]:
                discount_data[site][discount_name] = {'Quantity': 0, 'Price': 0, 'Amount': 0}
            
            if discount_name not in discount_data[Corporation_Totals_Name]:
                discount_data[Corporation_Totals_Name][discount_name] = {'Quantity': 0, 'Price': 0, 'Amount': 0}
            
            discount_data[site][discount_name]['Quantity'] += quantity
            discount_data[Corporation_Totals_Name][discount_name]['Quantity'] += quantity
            if discount_name not in miscDisctAverages:
                miscDisctAverages[discount_name] = []
            miscDisctAverages[discount_name].append(price if pd.notna(price) else 0)
            discount_data[site][discount_name]['Amount'] += amount
            discount_data[Corporation_Totals_Name][discount_name]['Amount'] += amount
    
    # Average the prices
    # Used for Misc. discounts to see the average amount of money paid out to customers
    locations = list(discount_data.keys())

    for location in locations:
        for discount in discount_data[location]:
            try:
                discount_data[location][discount]['Price'] = sum(miscDisctAverages[discount])/len(miscDisctAverages[discount])
                discount_data[Corporation_Totals_Name][discount]['Price'] = sum(miscDisctAverages[discount])/len(miscDisctAverages[discount])
            except:
                discount_data[location][discount]['Price'] = 0

    # Load the template workbook
    ws = wb['Discount Mix by Location']
    # Initial starting column for the first location
    start_col = 1

    # Populate the worksheet
    for location, data in discount_data.items():
        copy_cells(ws, 1, 3, 1, 6, 1, start_col)
        populate_discountmix_worksheet(ws, 1, start_col, location, data)
        start_col += 7
    
    #########################################################################################################################################################################
    ############################################################################ GRAPHS CREATION ############################################################################
    #########################################################################################################################################################################
        


    if 'Sales Mix Charts (Des Plaine)' in wb.sheetnames and desplainesOverride == True:
        #wb.remove('Sales Mix by Location (Visual)')
        wb['Sales Mix Charts (Des Plaine)'].cell(row=29, column=1, value='NOTICE: Due to the free wash weekend, Des Plaines will show an abnormal amount of UShine wash sales which subsequently distorts some charts.')
        # Define the fill color (in this case, yellow)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Apply the fill color to cells A29 to F29
        for column in range(1, 7):  # Columns A to F
            cell = wb['Sales Mix Charts (Des Plaine)'].cell(row=29, column=column)
            cell.fill = yellow_fill
            cell.font = Font(size=12, bold=True)


    wbName = f"Sales Mix - {short_date_str}"
    return wb, wbName

wb, wbName = createSalesMixSheet("input.csv", "input_tac.csv", "Sales_Mix_Template.xlsx")
wb.save(wbName)
