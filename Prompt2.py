import pandas as pd
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image

def calculate_contribution_margin(row):
    """
    Helper function to calculate Contribution Margin % with special rules.
    """
    total_revenue = row['Total Revenue (k NOK)']
    production_costs = row.get('Production Costs (k NOK)', 0)  # Use 0 if Production Costs column is missing

    if total_revenue == 0:
        if production_costs == 0:
            return 0  # Rule: Total Revenue = 0 and Production Costs = 0
        elif production_costs > 0:
            return -100  # Rule: Total Revenue = 0 and Production Costs > 0
    elif total_revenue < 0:
        # Rule: Flip the sign for negative Total Revenue
        return -(row['Contribution Margin (k NOK)'] / total_revenue) * 100

    return (row['Contribution Margin (k NOK)'] / total_revenue) * 100

def calculate_budget_contribution_margin(row):
    """
    Helper function to calculate Budget Contribution Margin % with special rules.
    """
    budget_total_revenue = row['Budget Total Revenue (k NOK)']
    if budget_total_revenue < 0:
        # Rule: Flip the sign for negative Budget Total Revenue
        return -(row['Budget Contribution Margin (k NOK)'] / budget_total_revenue) * 100
    elif budget_total_revenue != 0:
        return (row['Budget Contribution Margin (k NOK)'] / budget_total_revenue) * 100
    else:
        return 0

def generate_bar_chart(data, year, month=None, prev_data=None, output_dir="./charts"):
    """
    Generate and save a bar chart comparing Contribution Margin % for each service line or months,
    highlighting underperforming service lines and including explanations.

    Parameters:
        data (DataFrame): The input dataset filtered for the timeframe.
        year (int): The year for which the chart is generated.
        month (int or None): The numeric month (1-12) for which the chart is generated. None indicates full year.
        prev_data (DataFrame or None): Data for the previous timeframe (month or year) for comparison.
        output_dir (str): Directory to save the chart images.

    Returns:
        str: The file path of the saved chart.
    """
    os.makedirs(output_dir, exist_ok=True)
    if month:
        title = f"Contribution Margin % - Month {month} {year}"
        filename = f"contribution_margin_{year}_{month}.png"
    else:
        title = f"Contribution Margin % by Service Line - Full Year {year}"
        filename = f"contribution_margin_{year}_full_year.png"

    # Aggregate data
    if month:
        aggregated_data = data.groupby('Service Areas Shortname').agg({
            'Contribution Margin (k NOK)': 'sum',
            'Total Revenue (k NOK)': 'sum',
            'Production Costs (k NOK)': 'sum',
            'Budget Contribution Margin (k NOK)': 'sum',
            'Budget Total Revenue (k NOK)': 'sum'
        }).reset_index()
    else:
        aggregated_data = data.groupby('Service Areas Shortname').agg({
            'Contribution Margin (k NOK)': 'sum',
            'Total Revenue (k NOK)': 'sum',
            'Production Costs (k NOK)': 'sum',
            'Budget Contribution Margin (k NOK)': 'sum',
            'Budget Total Revenue (k NOK)': 'sum'
        }).reset_index()

    # Calculate Contribution Margin % and Budget Contribution Margin %
    aggregated_data['Contribution Margin %'] = aggregated_data.apply(calculate_contribution_margin, axis=1)
    aggregated_data['Budget Contribution Margin %'] = aggregated_data.apply(calculate_budget_contribution_margin, axis=1)

    # Add comparison with previous data (YoY or MoM)
    if prev_data is not None:
        prev_data['Service Areas Shortname'] = prev_data['Service Areas Shortname'].str.strip().str.lower()
        aggregated_data['Service Areas Shortname'] = aggregated_data['Service Areas Shortname'].str.strip().str.lower()

        prev_aggregated = prev_data.groupby('Service Areas Shortname').agg({
            'Contribution Margin (k NOK)': 'sum',
            'Total Revenue (k NOK)': 'sum',
            'Production Costs (k NOK)': 'sum'
        }).reset_index()

        prev_aggregated['Prev Contribution Margin %'] = prev_aggregated.apply(calculate_contribution_margin, axis=1)

        aggregated_data = aggregated_data.merge(
            prev_aggregated[['Service Areas Shortname', 'Prev Contribution Margin %']],
            on='Service Areas Shortname', how='left'
        )

        # Calculate percentage change (YoY or MoM)
        aggregated_data['Change %'] = aggregated_data.apply(
            lambda row: (
                (row['Contribution Margin %'] - row['Prev Contribution Margin %'])
                / abs(row['Prev Contribution Margin %']) * 100
            ) if pd.notnull(row['Prev Contribution Margin %']) and row['Prev Contribution Margin %'] != 0 else 0,
            axis=1
        )
    else:
        aggregated_data['Change %'] = None

    aggregated_data.fillna(0, inplace=True)
    aggregated_data['Underperforming'] = aggregated_data['Contribution Margin %'] < aggregated_data['Budget Contribution Margin %']

    # Plotting
    plt.figure(figsize=(10, 6))
    bars = plt.bar(
        aggregated_data['Service Areas Shortname'],
        aggregated_data['Contribution Margin %'],
        color=['red' if underperform else 'green' for underperform in aggregated_data['Underperforming']]
    )

    for bar, cm, change in zip(
        bars,
        aggregated_data['Contribution Margin %'],
        aggregated_data['Change %']
    ):
        cm_str = f"{cm:.2f}%" if pd.notnull(cm) else "N/A"
        change_str = f"({change:+.2f}%)" if pd.notnull(change) else ""
        label = f"{cm_str}\n{change_str}".strip()
        plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5, label, ha='center', va='bottom', fontsize=10)

    # Add explanation box (legend)
    legend_text = (
        "Color Legend:\n\n"
        "Green: On or above budget\n"
        "Red: Below budget\n\n"
        "Bar Values: Contribution Margin %\n"
        f"({'MoM' if month else 'YoY'} % Change in parentheses)"
    )
    plt.text(
        1.02, 0.5, legend_text,
        transform=plt.gca().transAxes,
        fontsize=10,
        bbox=dict(facecolor='white', alpha=0.7)
    )

    plt.axhline(y=0, color='black', linewidth=0.8, linestyle='--', label="0% Benchmark")
    plt.xlabel('Service Line' if not month else 'Service Line')
    plt.ylabel('Contribution Margin %')
    plt.title(title)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    file_path = os.path.join(output_dir, filename)
    plt.savefig(file_path)
    plt.close()

    return file_path

def analyze_and_generate_charts(data, output_dir="./charts"):
    # Ensure numeric Month column
    if 'Month' not in data.columns:
        data['Month'] = pd.to_datetime(data['Date Column']).dt.month

    results = []
    for year in data['Year'].unique():
        yearly_data = data[data['Year'] == year]
        prev_year_data = data[data['Year'] == year - 1] if (year - 1) in data['Year'].unique() else None

        for month in range(1, 13):  # Iterate through numeric months 1 to 12
            monthly_data = yearly_data[yearly_data['Month'] == month]

            # Handle January: Use December of the previous year
            if month == 1 and prev_year_data is not None:
                prev_month_data = prev_year_data[prev_year_data['Month'] == 12]
            else:
                prev_month_data = yearly_data[yearly_data['Month'] == month - 1]

            chart_path = generate_bar_chart(
                monthly_data, year=year, month=month, prev_data=prev_month_data, output_dir=output_dir
            )
            results.append({
                'Year': year,
                'Month': month,
                'Service Line': 'All',
                'Chart Path': chart_path
            })

    return pd.DataFrame(results)

def save_results_with_images(results_df, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Contribution Margin Analysis"
    headers = list(results_df.columns)
    ws.append(headers)

    for _, row in results_df.iterrows():
        row_data = [str(row[col]) if pd.notnull(row[col]) else '' for col in headers if col != 'Chart Path']
        ws.append(row_data)
        if isinstance(row['Chart Path'], str) and os.path.exists(row['Chart Path']):
            img = Image(row['Chart Path'])
            img.width = 120
            img.height = 80
            ws.add_image(img, f"E{ws.max_row}")

    wb.save(output_file)

# Define the file path and output directory
file_path = "report1" #Insert Directory Report1
output_dir = " " #Insert Output Directory
output_excel = " " #Insert Output Directory

try:
    data = pd.read_excel(file_path)
    print("File loaded successfully!")

    final_results = analyze_and_generate_charts(data, output_dir=output_dir)

    if not final_results.empty:
        print(f"\nCharts Generated for All Timeframes.")
        print(final_results.head())

        save_results_with_images(final_results, output_excel)
        print(f"\nResults with images saved to: {output_excel}")

except Exception as e:
    print(f"An error occurred: {e}")
