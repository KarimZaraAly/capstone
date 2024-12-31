import pandas as pd

def analyze_projects(data, num_projects_options=[3, 7]):
    """
    Identify the projects with the lowest and highest Contribution Margin After Adjustments
    for all years, months, and service lines in the dataset, ensuring each project is aggregated
    and appears only once.

    Parameters:
        data (DataFrame): The input dataset.
        num_projects_options (list): A list of numbers for Top/Bottom analysis (e.g., [3, 7]).

    Returns:
        DataFrame: A DataFrame containing the top and bottom projects for all years, months, and service lines.
    """
    results = []

    # Aggregate data at the project level within service, year, and month
    aggregated_data = data.groupby(
        ['Year', 'Service Areas Shortname', 'Month', 'Prosjekt-ID']
    ).agg({
        'Contribution_Margin_After': 'sum'  # Aggregate the contribution margin
    }).reset_index()

    # Loop through each unique year and service line
    for year in aggregated_data['Year'].unique():
        yearly_data = aggregated_data[aggregated_data['Year'] == year]
        
        for service_line in yearly_data['Service Areas Shortname'].unique():
            service_line_data = yearly_data[yearly_data['Service Areas Shortname'] == service_line]
            
            # Monthly Analysis
            for month in service_line_data['Month'].unique():
                monthly_data = service_line_data[service_line_data['Month'] == month]
                
                # Perform analysis for each option (e.g., Bottom 3, Top 3, etc.)
                for num_projects in num_projects_options:
                    # Bottom analysis
                    bottom_projects = (
                        monthly_data.sort_values(by='Contribution_Margin_After', ascending=True)
                        .head(num_projects)
                    )
                    bottom_details = "\n".join(bottom_projects['Prosjekt-ID'].astype(str))
                    results.append({
                        'Year': year,
                        'Month': month,
                        'Service Line': service_line,
                        'Analysis Type': f"Bottom {num_projects}",
                        'Details': bottom_details
                    })

                    # Top analysis
                    top_projects = (
                        monthly_data.sort_values(by='Contribution_Margin_After', ascending=False)
                        .head(num_projects)
                    )
                    top_details = "\n".join(top_projects['Prosjekt-ID'].astype(str))
                    results.append({
                        'Year': year,
                        'Month': month,
                        'Service Line': service_line,
                        'Analysis Type': f"Top {num_projects}",
                        'Details': top_details
                    })

            # Full Year Analysis
            # Re-aggregate data for the entire year to prevent duplicates
            yearly_aggregated_data = aggregated_data.groupby(
                ['Year', 'Service Areas Shortname', 'Prosjekt-ID']
            ).agg({
                'Contribution_Margin_After': 'sum'
            }).reset_index()

            yearly_service_data = yearly_aggregated_data[
                (yearly_aggregated_data['Year'] == year) &
                (yearly_aggregated_data['Service Areas Shortname'] == service_line)
            ]

            for num_projects in num_projects_options:
                # Bottom analysis for the year
                bottom_projects_year = (
                    yearly_service_data.sort_values(by='Contribution_Margin_After', ascending=True)
                    .head(num_projects)
                )
                bottom_details_year = "\n".join(bottom_projects_year['Prosjekt-ID'].astype(str))
                results.append({
                    'Year': year,
                    'Month': "Full Year",
                    'Service Line': service_line,
                    'Analysis Type': f"Bottom {num_projects}",
                    'Details': bottom_details_year
                })

                # Top analysis for the year
                top_projects_year = (
                    yearly_service_data.sort_values(by='Contribution_Margin_After', ascending=False)
                    .head(num_projects)
                )
                top_details_year = "\n".join(top_projects_year['Prosjekt-ID'].astype(str))
                results.append({
                    'Year': year,
                    'Month': "Full Year",
                    'Service Line': service_line,
                    'Analysis Type': f"Top {num_projects}",
                    'Details': top_details_year
                })

    # Convert to DataFrame
    results_df = pd.DataFrame(results)
    
    return results_df

# Define the file path
file_path = "Report2" #Insert Directory Report2

try:
    # Load data from Excel file
    data = pd.read_excel(file_path)
    print("File loaded successfully!")
    
    # Perform analysis for all years, months, and service lines
    results = analyze_projects(data, num_projects_options=[3, 7])
    
    if results is not None:
        print(f"\nAnalysis Completed for All Years, Months, and Service Lines.")
        print(results.head())  # Display the first few rows for verification
    
        # Save results to a new Excel file
        output_file = " " #Insert Output Directory
        results.to_excel(output_file, index=False)
        print(f"\nResults saved to: {output_file}")

except Exception as e:
    print(f"An error occurred: {e}")
