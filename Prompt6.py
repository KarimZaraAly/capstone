import pandas as pd

def identify_outliers(data, threshold_multiplier=2):
    """
    Identify employees who bill significantly more hours for specific tasks
    within each service line and role, at monthly and yearly levels, aggregating rows
    for the same task and employee. Consolidate all outliers for a specific time period in one cell.

    Parameters:
        data (DataFrame): The input dataset containing tasks, employees, roles, and billed hours.
        threshold_multiplier (float): Multiplier for the standard deviation to define outliers.

    Returns:
        DataFrame: A DataFrame highlighting aggregated summaries of outliers for each time period.
    """
    results = []

    # Aggregate data by Task, Employee, and Role for monthly analysis
    data_aggregated_monthly = data.groupby(
        ['Service Areas Shortname', 'Year', 'Month', 'Role', 'Task', 'Employee']
    )['Billable_Hours'].sum().reset_index()

    # Group data by Service, Role, Year, and Month
    grouped_data = data_aggregated_monthly.groupby(['Service Areas Shortname', 'Role', 'Year', 'Month'])

    for (service, role, year, month), group_data in grouped_data:
        # Initialize a list to hold all outlier summaries for this group
        consolidated_summaries = []

        for task, task_data in group_data.groupby('Task'):
            # Calculate statistics for the task
            task_mean = task_data['Billable_Hours'].mean()
            task_std = task_data['Billable_Hours'].std()

            # Define threshold for outliers
            threshold = task_mean + threshold_multiplier * task_std

            # Identify outliers
            task_data['Outlier'] = task_data['Billable_Hours'] > threshold
            outliers = task_data[task_data['Outlier']]

            # Collect summaries for all outliers
            for _, row in outliers.iterrows():
                summary = (
                    f"Employee: {row['Employee']}, Task: {task}, "
                    f"Billable Hours: {row['Billable_Hours']:.2f}, Task Mean: {task_mean:.2f}"
                )
                consolidated_summaries.append(summary)

        if consolidated_summaries:
            # Combine all summaries for this group into one cell
            aggregated_summary = "\n".join(consolidated_summaries)
            results.append({
                'Service Line': service,
                'Role': role,
                'Year': year,
                'Month': month,
                'Outlier Summaries': aggregated_summary
            })

    # Perform yearly analysis
    data_aggregated_yearly = data.groupby(
        ['Service Areas Shortname', 'Role', 'Year', 'Task', 'Employee']
    )['Billable_Hours'].sum().reset_index()

    yearly_grouped_data = data_aggregated_yearly.groupby(['Service Areas Shortname', 'Role', 'Year'])

    for (service, role, year), group_data in yearly_grouped_data:
        # Initialize a list to hold all outlier summaries for this group
        consolidated_summaries = []

        for task, task_data in group_data.groupby('Task'):
            # Calculate statistics for the task
            task_mean = task_data['Billable_Hours'].mean()
            task_std = task_data['Billable_Hours'].std()

            # Define threshold for outliers
            threshold = task_mean + threshold_multiplier * task_std

            # Identify outliers
            task_data['Outlier'] = task_data['Billable_Hours'] > threshold
            outliers = task_data[task_data['Outlier']]

            # Collect summaries for all outliers
            for _, row in outliers.iterrows():
                summary = (
                    f"Employee: {row['Employee']}, Task: {task}, "
                    f"Billable Hours: {row['Billable_Hours']:.2f}, Task Mean: {task_mean:.2f}"
                )
                consolidated_summaries.append(summary)

        if consolidated_summaries:
            # Combine all summaries for this group into one cell
            aggregated_summary = "\n".join(consolidated_summaries)
            results.append({
                'Service Line': service,
                'Role': role,
                'Year': year,
                'Month': "Yearly",
                'Outlier Summaries': aggregated_summary
            })

    # Convert results to DataFrame
    results_df = pd.DataFrame(results)
    
    return results_df

# Define the file path
file_path = "report" #Insert Directory Report2

try:
    # Load data from Excel file
    data = pd.read_excel(file_path)
    print("File loaded successfully!")
    
    # Identify outliers
    outlier_results = identify_outliers(data, threshold_multiplier=2)
    
    if not outlier_results.empty:
        print("\nOutliers Identified:")
        print(outlier_results[['Service Line', 'Role', 'Year', 'Month', 'Outlier Summaries']].head())
    
        # Save results to a new Excel file
        output_file = " " #Insert output Directory
        outlier_results.to_excel(output_file, index=False)
        print(f"\nOutliers saved to: {output_file}")
    else:
        print("No significant outliers found.")
        
except Exception as e:
    print(f"An error occurred: {e}")



## input_file_path = "/Users/annichenlarsen/Documents/Karim/Upload files/Monthly_Project_Task_Level_Hypercube.xlsx



