import pandas as pd
import numpy as np
import os
import calendar  

# Load data
file_path = r" " #Insert Filepath of uploadfile 1

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Extract the directory of the input file
output_directory = os.path.dirname(file_path)

# File paths for outputs
output_file1 = os.path.join(output_directory, 'service_performance_summary.xlsx')
output_file2 = os.path.join(output_directory, 'utilization_rate_summary.xlsx')

# Define functions for quarterly and half-yearly labeling
def get_quarter(month):
    if month in [1, 2, 3]:
        return 'Q1'
    elif month in [4, 5, 6]:
        return 'Q2'
    elif month in [7, 8, 9]:
        return 'Q3'
    elif month in [10, 11, 12]:
        return 'Q4'

def get_half_year(month):
    return 'H1' if month <= 6 else 'H2'

# Set up dictionaries to store results
results_performance = {}
results_utilization = {}

# Filtering only the required columns and metrics
for year in df['Year'].unique():
    yearly_df = df[df['Year'] == year]
    year_summary_performance = []
    year_summary_utilization = []

    # Iterate through each service area
    for service in yearly_df['Service Areas Shortname'].unique():
        service_df = yearly_df[yearly_df['Service Areas Shortname'] == service]
        
        # Yearly summary using aggregated values
        total_billable_hours = service_df['Billable Hours'].sum(skipna=True)
        total_hours = service_df['Total Hours'].sum(skipna=True)
        utilized_hours = service_df['Utilized Hours'].sum(skipna=True)

        # Calculations
        billing_rate_avg = (total_billable_hours / total_hours * 100) if total_hours > 0 else 0
        utilization_rate_avg = (utilized_hours / total_hours * 100) if total_hours > 0 else 0

        # Debugging logs for verification
        print(f"Yearly Debug | Service: {service}, Billing Rate: {billing_rate_avg:+.2f}, Utilization Rate: {utilization_rate_avg:+.2f}")

        # Store yearly summary for utilization
        year_summary_performance.append(
            f"{service}:\n- Billing Rate%: {billing_rate_avg:+.2f}"
        )
        year_summary_utilization.append(
            f"{service}:\n- Billing Rate%: {billing_rate_avg:+.2f}\n"
            f"- Utilization Rate%: {utilization_rate_avg:+.2f}"
        )

        # Quarterly and half-yearly summaries
        for period_func, period_name in [(get_quarter, 'Quarter'), (get_half_year, 'Half-Year')]:
            for period in service_df['Month'].apply(period_func).unique():
                period_df = service_df[service_df['Month'].apply(period_func) == period]

                total_billable_hours_period = period_df['Billable Hours'].sum(skipna=True)
                total_hours_period = period_df['Total Hours'].sum(skipna=True)
                utilized_hours_period = period_df['Utilized Hours'].sum(skipna=True)

                # Calculate metrics
                billing_rate_period = (total_billable_hours_period / total_hours_period * 100) if total_hours_period > 0 else 0
                utilization_rate_period = (utilized_hours_period / total_hours_period * 100) if total_hours_period > 0 else 0

                # Debugging logs for verification
                print(f"Period Debug | Service: {service}, Period: {period}, Utilization Rate: {utilization_rate_period:+.2f}")

                period_summary_utilization = (
                    f"{service}:\n- Billing Rate%: {billing_rate_period:+.2f}\n"
                    f"- Utilization Rate%: {utilization_rate_period:+.2f}"
                )

                # Format the Timeline key
                timeline_key = f'{period}, {year}'  # e.g., 'Q1, 2023' or 'H1, 2023'

                results_utilization[timeline_key] = results_utilization.get(timeline_key, "") + period_summary_utilization + "\n\n"

        # Monthly summaries
        for month in service_df['Month'].unique():
            month_name = calendar.month_name[month]
            month_df = service_df[service_df['Month'] == month]

            total_billable_hours_month = month_df['Billable Hours'].sum(skipna=True)
            total_hours_month = month_df['Total Hours'].sum(skipna=True)
            utilized_hours_month = month_df['Utilized Hours'].sum(skipna=True)

            # Calculate metrics
            billing_rate_month = (total_billable_hours_month / total_hours_month * 100) if total_hours_month > 0 else 0
            utilization_rate_month = (utilized_hours_month / total_hours_month * 100) if total_hours_month > 0 else 0

            # Debugging logs for verification
            print(f"Monthly Debug | Service: {service}, Month: {month_name}, Utilization Rate: {utilization_rate_month:+.2f}")

            month_summary_utilization = (
                f"{service}:\n- Billing Rate%: {billing_rate_month:+.2f}\n"
                f"- Utilization Rate%: {utilization_rate_month:+.2f}"
            )

            timeline_key = f'{month_name}, {year}'  # e.g., 'January, 2023'

            results_utilization[timeline_key] = results_utilization.get(timeline_key, "") + month_summary_utilization + "\n\n"

    # Store yearly summaries
    results_performance[f'{year}'] = "\n\n".join(year_summary_performance)
    results_utilization[f'{year}'] = "\n\n".join(year_summary_utilization)

# Create structured DataFrames for output
output_data_performance = [{'Timeline': period, 'Summary': summary} for period, summary in results_performance.items()]
output_df_performance = pd.DataFrame(output_data_performance)

output_data_utilization = [{'Timeline': period, 'Summary': summary} for period, summary in results_utilization.items()]
output_df_utilization = pd.DataFrame(output_data_utilization)

# Export results to Excel
output_df_performance.to_excel(output_file1, index=False)
output_df_utilization.to_excel(output_file2, index=False)

print(f"The performance summary has been successfully saved to {output_file1}")
print(f"The utilization summary has been successfully saved to {output_file2}")
