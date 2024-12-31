import pandas as pd
import os

def analyze_service_and_role_performance(data, output_path):
    """
    Analyze service performance by calculating metrics and their MoM comparisons
    for Billing Rate %, Adjustments, and Hourly Rate, grouped by Service, Year, and Month,
    and summarize all roles into a single cell.

    Parameters:
        data (DataFrame): The input dataset.
        output_path (str): The path to save the Excel file.
    """
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Calculate key metrics grouped by Service, Year, Month, and Role
    grouped_data = data.groupby(['Service Areas Shortname', 'Year', 'Month', 'Role']).agg({
        'Billable Hours': 'sum',
        'Total Hours': 'sum',
        'Adjustments': 'sum',
        'Total Revenue': 'sum'
    }).reset_index()

    # Calculate metrics
    grouped_data['Billing_Rate_%'] = (grouped_data['Billable Hours'] / grouped_data['Total Hours']) * 100
    grouped_data['Hourly_Rate'] = grouped_data['Total Revenue'] / grouped_data['Total Hours']

    # Add previous month data for MoM calculations
    grouped_data['Prev_Billing_Rate_%'] = grouped_data.groupby(['Service Areas Shortname', 'Role', 'Year'])['Billing_Rate_%'].shift(1)
    grouped_data['Prev_Adjustments'] = grouped_data.groupby(['Service Areas Shortname', 'Role', 'Year'])['Adjustments'].shift(1)
    grouped_data['Prev_Hourly_Rate'] = grouped_data.groupby(['Service Areas Shortname', 'Role', 'Year'])['Hourly_Rate'].shift(1)

    # Calculate Month-over-Month (MoM) changes
    grouped_data['MoM_Billing_Rate_%'] = ((grouped_data['Billing_Rate_%'] - grouped_data['Prev_Billing_Rate_%']) / grouped_data['Prev_Billing_Rate_%']) * 100
    grouped_data['MoM_Adjustments'] = ((grouped_data['Adjustments'] - grouped_data['Prev_Adjustments']) / grouped_data['Prev_Adjustments']) * 100
    grouped_data['MoM_Hourly_Rate'] = ((grouped_data['Hourly_Rate'] - grouped_data['Prev_Hourly_Rate']) / grouped_data['Prev_Hourly_Rate']) * 100

    # Fill missing MoM values (e.g., for the first month of each year)
    grouped_data[['MoM_Billing_Rate_%', 'MoM_Adjustments', 'MoM_Hourly_Rate']] = grouped_data[
        ['MoM_Billing_Rate_%', 'MoM_Adjustments', 'MoM_Hourly_Rate']
    ].fillna(0)

    # Combine all metrics into a single cell for each role
    grouped_data['Role_Performance'] = grouped_data.apply(
        lambda row: (
            f"Role: {row['Role']}, Billing Rate %: {row['Billing_Rate_%']:.2f}, MoM: {row['MoM_Billing_Rate_%']:.2f}%\n"
            f"Adjustments: {row['Adjustments']:.2f}, MoM: {row['MoM_Adjustments']:.2f}%\n"
            f"Hourly Rate: {row['Hourly_Rate']:.2f}, MoM: {row['MoM_Hourly_Rate']:.2f}%"
        ),
        axis=1
    )

    # Summarize all roles into a single cell for each Service, Year, and Month
    summary_data = grouped_data.groupby(['Service Areas Shortname', 'Year', 'Month']).agg({
        'Role_Performance': lambda x: "\n\n".join(x)  # Combine all roles into a single cell
    }).reset_index()

    # Save results to Excel
    output_file = os.path.join(output_path, "Service_Performance_With_Roles.xlsx")
    summary_data.to_excel(output_file, index=False, sheet_name="Service and Role Performance")
    print(f"Analysis successfully saved to {output_file}")


# ================================
# Main Execution
# ================================
# Define input and output paths
input_data_path = " " #Insert Directory Report3
output_dir = " " #Insert Output Directory

try:
    # Load the input data
    input_data = pd.read_excel(input_data_path)
    print("Data loaded successfully!")

    # Analyze service and role performance and save results
    analyze_service_and_role_performance(input_data, output_dir)

except Exception as e:
    print(f"An error occurred: {e}")
