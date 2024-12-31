import pandas as pd

def identify_high_cost_tasks(data, num_tasks_list):
    """
    Identify the top N tasks contributing the most to cost for all years, months, service lines, and roles.

    Parameters:
        data (DataFrame): The input dataset.
        num_tasks_list (list): A list of task counts to analyze.

    Returns:
        DataFrame: A DataFrame containing the top tasks for all years, months, service lines, and roles.
    """
    results = []

    # Loop through each unique year, service line, and role
    for year in data['Year'].unique():
        yearly_data = data[data['Year'] == year]

        for service_line in yearly_data['Service Areas Shortname'].unique():
            service_line_data = yearly_data[yearly_data['Service Areas Shortname'] == service_line]

            for role in service_line_data['Role'].unique():
                role_data = service_line_data[service_line_data['Role'] == role]

                # Monthly Analysis
                for month in role_data['Month'].unique():
                    monthly_data = role_data[role_data['Month'] == month]

                    # Aggregate costs by task
                    monthly_aggregated = monthly_data.groupby('Task')['Cost'].agg(['sum', 'mean']).reset_index()
                    monthly_aggregated = monthly_aggregated.sort_values(by='sum', ascending=False)

                    for num_tasks in num_tasks_list:
                        # Get top N tasks
                        top_tasks = monthly_aggregated.head(num_tasks)
                        tasks_list = top_tasks['Task'].tolist()
                        tasks_only = "\n".join(tasks_list)
                        top_tasks_details = "\n".join(
                            [f"Task: {row['Task']}, Total Cost: {row['sum']}, Average Cost: {row['mean']:.2f}"
                             for _, row in top_tasks.iterrows()]
                        )
                        structured_response = "\n".join(
                            [f"{row['Task']} - Avg Cost: {row['mean']:.2f}" for _, row in top_tasks.iterrows()]
                        )
                        results.append({
                            'Year': year,
                            'Month': month,
                            'Service Line': service_line,
                            'Role': role,
                            'Analysis Type': f"Top {num_tasks} Tasks",
                            'Details': top_tasks_details,
                            'Tasks Only': tasks_only,
                            'Structured Response': structured_response
                        })

                # Full Year Analysis
                yearly_aggregated = role_data.groupby('Task')['Cost'].agg(['sum', 'mean']).reset_index()
                yearly_aggregated = yearly_aggregated.sort_values(by='sum', ascending=False)

                for num_tasks in num_tasks_list:
                    # Get top N tasks for the full year
                    top_tasks_year = yearly_aggregated.head(num_tasks)
                    tasks_list_year = top_tasks_year['Task'].tolist()
                    tasks_only_year = "\n".join(tasks_list_year)
                    top_tasks_year_details = "\n".join(
                        [f"Task: {row['Task']}, Total Cost: {row['sum']}, Average Cost: {row['mean']:.2f}"
                         for _, row in top_tasks_year.iterrows()]
                    )
                    structured_response_year = "\n".join(
                        [f"{row['Task']} - Avg Cost: {row['mean']:.2f}" for _, row in top_tasks_year.iterrows()]
                    )
                    results.append({
                        'Year': year,
                        'Month': "Full Year",
                        'Service Line': service_line,
                        'Role': role,
                        'Analysis Type': f"Top {num_tasks} Tasks",
                        'Details': top_tasks_year_details,
                        'Tasks Only': tasks_only_year,
                        'Structured Response': structured_response_year
                    })

    # Convert results to DataFrame
    results_df = pd.DataFrame(results)

    return results_df

# Define the file path
file_path = "Report2" #Insert Directory Report2
try:
    # Load data from Excel file
    data = pd.read_excel(file_path)
    print("File loaded successfully!")
    
    # Specify the numbers of tasks to analyze (e.g., 3, 5, 10)
    num_tasks_input = input("Enter the numbers of tasks to analyze, separated by commas (e.g., 3,5,10): 3, 5 ")
    num_tasks_list = [int(x.strip()) for x in num_tasks_input.split(",")]
    
    # Perform analysis
    results = identify_high_cost_tasks(data, num_tasks_list=num_tasks_list)
    
    if results is not None:
        print(f"\nAnalysis Completed for All Years, Months, Service Lines, and Roles.")
        print(results.head())  # Display the first few rows for verification
    
        # Save results to a new Excel file
        output_file = f" " #Insert Output Directory
        results.to_excel(output_file, index=False)
        print(f"\nResults saved to: {output_file}")

except Exception as e:
    print(f"An error occurred: {e}")
