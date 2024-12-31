import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

def aggregate_employee_data(data):
    """
    Aggregate data by Employee, Month, Year, Service Line, and Role.
    Summarize metrics to avoid duplicate entries for the same employee.
    """
    return data.groupby(
        ['Year', 'Month', 'Service Areas Shortname', 'Role', 'Employee'], as_index=False
    ).agg({
        'Billable_Hours': 'sum',         # Sum over all tasks
        'Revenue': 'sum',               # Sum over all tasks
        'Cost': 'sum',                  # Sum over all tasks
        'Adjustments': 'sum'            # Sum over all tasks
    })

def detect_outliers_iqr(data, column):
    """
    Detect outliers in a given column using the Interquartile Range (IQR) method
    and label them as 'High' or 'Low'.
    """
    Q1 = data[column].quantile(0.25)
    Q3 = data[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    
    # Identify and label outliers
    data['Outlier_Type'] = np.where(
        data[column] < lower_bound, 'Low',
        np.where(data[column] > upper_bound, 'High', None)
    )
    return data[(data[column] < lower_bound) | (data[column] > upper_bound)]

def analyze_iqr_outliers(data, output_path):
    """
    Perform IQR analysis for each metric and summarize employee outliers for each month, year,
    service line, and role. Label outliers as 'High' or 'Low' and format summaries neatly.
    """
    # Ensure the output directory exists
    os.makedirs(output_path, exist_ok=True)

    # Metrics to analyze
    metrics = ['Billable_Hours', 'Revenue', 'Cost', 'Adjustments']
    
    results = []
    
    # Aggregate data to avoid duplicate entries per employee
    aggregated_data = aggregate_employee_data(data)

    # Loop through each year, month, service line, and role
    for year in aggregated_data['Year'].unique():
        for month in aggregated_data[aggregated_data['Year'] == year]['Month'].unique():
            for service_line in aggregated_data['Service Areas Shortname'].unique():
                for role in aggregated_data['Role'].unique():
                    # Filter data for the specific year, month, service line, and role
                    filtered_data = aggregated_data[
                        (aggregated_data['Year'] == year) &
                        (aggregated_data['Month'] == month) &
                        (aggregated_data['Service Areas Shortname'] == service_line) &
                        (aggregated_data['Role'] == role)
                    ]
                    
                    if filtered_data.empty:
                        continue
                    
                    # Detect outliers for each metric
                    outlier_summary = []
                    for metric in metrics:
                        outliers = detect_outliers_iqr(filtered_data, metric)
                        if not outliers.empty:
                            for _, outlier in outliers.iterrows():
                                employee = outlier['Employee']
                                value = outlier[metric]
                                outlier_type = outlier['Outlier_Type']
                                outlier_summary.append(
                                    f"- Employee: {employee}, Metric: {metric}, Value: {value:.2f}, Type: {outlier_type}"
                                )
                    
                    # Combine outliers into a structured summary
                    if outlier_summary:
                        combined_outliers = "\n".join(outlier_summary)
                    else:
                        combined_outliers = "No Outliers"
                    
                    # Append the result
                    results.append({
                        'Year': year,
                        'Month': month,
                        'Service Line': service_line,
                        'Role': role,
                        'Outlier Summary': combined_outliers
                    })
                    
                    # Visualization: Split boxplots by metric (Monthly)
                    plt.figure(figsize=(15, 10))  # Adjust figure size as needed
                    rows, cols = 2, 2  # Define the grid layout (2 rows, 2 columns for 4 metrics)

                    for i, metric in enumerate(metrics):
                        plt.subplot(rows, cols, i + 1)  # Create a subplot for each metric
                        filtered_data.boxplot(column=[metric])
                        plt.title(metric)  # Title for each metric
                        # Add red dashed lines for IQR bounds
                        Q1 = filtered_data[metric].quantile(0.25)
                        Q3 = filtered_data[metric].quantile(0.75)
                        IQR = Q3 - Q1
                        lower_bound = Q1 - 1.5 * IQR
                        upper_bound = Q3 + 1.5 * IQR
                        plt.axhline(y=lower_bound, color='red', linestyle='--', label='Lower Bound')
                        plt.axhline(y=upper_bound, color='red', linestyle='--', label='Upper Bound')
                        plt.legend(loc='upper right')  # Optional: Add legend for clarity
                        plt.xticks([])  # Remove x-axis ticks for simplicity

                    plt.tight_layout()
                    # Save the multi-metric plot for the month
                    plot_path = os.path.join(output_path, f"split_boxplot_{service_line}_{role}_{year}_{month}.png")
                    plt.savefig(plot_path)
                    plt.close()

        # Perform Yearly Analysis
        for service_line in aggregated_data['Service Areas Shortname'].unique():
            for role in aggregated_data['Role'].unique():
                # Aggregate data for the specific year, service line, and role
                yearly_data = aggregated_data[
                    (aggregated_data['Year'] == year) &
                    (aggregated_data['Service Areas Shortname'] == service_line) &
                    (aggregated_data['Role'] == role)
                ].groupby(['Year', 'Service Areas Shortname', 'Role', 'Employee'], as_index=False).agg({
                    'Billable_Hours': 'sum',
                    'Revenue': 'sum',
                    'Cost': 'sum',
                    'Adjustments': 'sum'
                })

                if yearly_data.empty:
                    continue

                # Detect outliers for each metric
                outlier_summary = []
                for metric in metrics:
                    outliers = detect_outliers_iqr(yearly_data, metric)
                    if not outliers.empty:
                        for _, outlier in outliers.iterrows():
                            employee = outlier['Employee']
                            value = outlier[metric]
                            outlier_type = outlier['Outlier_Type']
                            outlier_summary.append(
                                f"- Employee: {employee}, Metric: {metric}, Value: {value:.2f}, Type: {outlier_type}"
                            )

                # Combine outliers into a structured summary
                if outlier_summary:
                    combined_outliers = "\n".join(outlier_summary)
                else:
                    combined_outliers = "No Outliers"

                # Append the result
                results.append({
                    'Year': year,
                    'Month': 'Full Year',  # Indicate yearly analysis
                    'Service Line': service_line,
                    'Role': role,
                    'Outlier Summary': combined_outliers
                })

                # Visualization: Split boxplots by metric (Yearly)
                plt.figure(figsize=(15, 10))  # Adjust figure size as needed
                rows, cols = 2, 2  # Define the grid layout (2 rows, 2 columns for 4 metrics)

                for i, metric in enumerate(metrics):
                    plt.subplot(rows, cols, i + 1)  # Create a subplot for each metric
                    yearly_data.boxplot(column=[metric])
                    plt.title(metric)  # Title for each metric
                    # Add red dashed lines for IQR bounds
                    Q1 = yearly_data[metric].quantile(0.25)
                    Q3 = yearly_data[metric].quantile(0.75)
                    IQR = Q3 - Q1
                    lower_bound = Q1 - 1.5 * IQR
                    upper_bound = Q3 + 1.5 * IQR
                    plt.axhline(y=lower_bound, color='red', linestyle='--', label='Lower Bound')
                    plt.axhline(y=upper_bound, color='red', linestyle='--', label='Upper Bound')
                    plt.legend(loc='upper right')  # Optional: Add legend for clarity
                    plt.xticks([])  # Remove x-axis ticks for simplicity

                plt.tight_layout()
                # Save the yearly plot
                plot_path = os.path.join(output_path, f"yearly_boxplot_{service_line}_{role}_{year}.png")
                plt.savefig(plot_path)
                plt.close()

    # Convert results to a DataFrame
    results_df = pd.DataFrame(results)
    
    # Save the results to an Excel file
    output_file = os.path.join(output_path, "IQR_Outlier_Analysis_By_Role.xlsx")
    results_df.to_excel(output_file, index=False, sheet_name="Outlier Analysis")
    print(f"Analysis and visualizations saved to: {output_file}")

## Main Execution
input_file_path = " " #Insert Directory Report1
output_dir = " " #Insert Output Directory

try:
    # Load data from Excel file
    data = pd.read_excel(input_file_path)
    print("Data loaded successfully!")

    # Perform IQR analysis
    analyze_iqr_outliers(data, output_dir)

except Exception as e:
    print(f"An error occurred: {e}")

