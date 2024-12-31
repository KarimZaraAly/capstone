import pandas as pd

def analyze_below_median_employees(data, cost_center, role, year=None, month=None):
    """
    Analyze employees whose aggregated average hourly rate is below the aggregated median
    for their service line and role for a specific time period.

    Parameters:
        data (DataFrame): The input dataset.
        cost_center (str): The cost center to filter on.
        role (str): The role to filter on.
        year (int or None): The year to filter. If None, include all years.
        month (int or None): The month to filter. If None, include the entire year.

    Returns:
        str: A string containing all low-performing employee details for the specified timeframe.
    """
    # Filter data for the specified cost center, role, year, and month
    filtered_data = data[
        (data['Cost Center'] == cost_center) & (data['Role'] == role)
    ]
    if year:
        filtered_data = filtered_data[filtered_data['Year'] == year]
    if month:
        filtered_data = filtered_data[filtered_data['Month'] == month]
    
    if filtered_data.empty:
        return None

    # Identify the service line for the filtered data
    service_line = filtered_data['Service Areas Shortname'].iloc[0]
    service_line_role_data = data[
        (data['Service Areas Shortname'] == service_line) & 
        (data['Role'] == role)
    ]
    if year:
        service_line_role_data = service_line_role_data[service_line_role_data['Year'] == year]
    if month:
        service_line_role_data = service_line_role_data[service_line_role_data['Month'] == month]

    # Aggregate service line data for median calculation
    service_line_role_aggregated = service_line_role_data.groupby('Employee ID', as_index=False).agg({
        'Total_Revenue': 'sum',
        'Total_Hours': 'sum'
    })
    # Calculate hourly rates after aggregation
    service_line_role_aggregated['Hourly_Rate'] = service_line_role_aggregated.apply(
        lambda row: row['Total_Revenue'] / row['Total_Hours'] if row['Total_Hours'] > 0 else None,
        axis=1
    )
    # Drop rows with no valid hourly rate for median calculation
    valid_hourly_rates = service_line_role_aggregated['Hourly_Rate'].dropna()
    median_hourly_rate = valid_hourly_rates.median()

    # Debugging: Print the median hourly rate
    print(f"Median Hourly Rate: {median_hourly_rate}")

    # Aggregate employee data for the specified time period
    aggregated_data = filtered_data.groupby('Employee ID', as_index=False).agg({
        'Total_Revenue': 'sum',
        'Total_Hours': 'sum'
    })
    # Calculate hourly rates after aggregation
    aggregated_data['Average Hourly Rate'] = aggregated_data.apply(
        lambda row: row['Total_Revenue'] / row['Total_Hours'] if row['Total_Hours'] > 0 else None,
        axis=1
    )

    # Debugging: Print aggregated employee data
    print("Aggregated Employee Data:\n", aggregated_data)

    # Filter employees below the median hourly rate
    below_median_employees = aggregated_data[
        aggregated_data['Average Hourly Rate'] < median_hourly_rate
    ]

    # Debugging: Print below-median employees
    print("Below Median Employees:\n", below_median_employees)

    # Summarize results for each employee
    employee_details = []
    for _, employee in below_median_employees.iterrows():
        employee_id = employee['Employee ID']
        hourly_rate = employee['Average Hourly Rate']
        employee_details.append(f"Employee ID: {employee_id}, Avg Hourly Rate: {hourly_rate:.2f}")

    # Combine all employee details into one string
    combined_details = "\n".join(employee_details)

    return combined_details

def analyze_all_cost_centers_roles(data):
    """
    Perform analysis for all years, months, and the full year, for each cost center and role.

    Parameters:
        data (DataFrame): The input dataset.

    Returns:
        DataFrame: A combined DataFrame containing analyses for all years, months, and cost center-role combinations.
    """
    results = []

    # Loop through each unique year, cost center, and role
    for year in data['Year'].unique():
        for cost_center in data['Cost Center'].unique():
            for role in data['Role'].unique():
                # Analyze by month
                for month in data[data['Year'] == year]['Month'].unique():
                    combined_details = analyze_below_median_employees(
                        data, cost_center, role, year=year, month=month
                    )
                    if combined_details:
                        results.append({
                            'Year': year,
                            'Month': month,
                            'Cost Center': cost_center,
                            'Role': role,
                            'Details': combined_details
                        })

                # Analyze for the full year
                combined_details = analyze_below_median_employees(
                    data, cost_center, role, year=year, month=None
                )
                if combined_details:
                    results.append({
                        'Year': year,
                        'Month': 'Full Year',
                        'Cost Center': cost_center,
                        'Role': role,
                        'Details': combined_details
                    })

    # Convert all results into a DataFrame
    return pd.DataFrame(results)

# Define the file path
file_path = " report3" #Insert Output Report3

try:
    # Load data from Excel file
    data = pd.read_excel(file_path)
    print("File loaded successfully!")
    
    # Perform analysis for all cost centers and roles
    final_results = analyze_all_cost_centers_roles(data)
    
    if not final_results.empty:
        print(f"\nAnalysis Completed for All Cost Centers and Roles.")
        print(final_results.head())  # Display the first few rows for verification
    
        # Save results to a new Excel file
        output_file = " " #Insert Output Directory
        final_results.to_excel(output_file, index=False)
        print(f"\nResults saved to: {output_file}")

except Exception as e:
    print(f"An error occurred: {e}")
