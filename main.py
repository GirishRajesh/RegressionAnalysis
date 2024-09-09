import pandas as pd
import statsmodels.api as sm
import numpy as np

# Define the path to your Excel file
file_path = 'Regression data for analsysis .xlsx'  # Make sure this matches the exact file name

# Load the dataset
df = pd.read_excel(file_path)

# Adjust the column names according to your data structure (make sure they match)
metrics = {
    'Netincome': ('Netincome PRE', 'Netincome POST'),
    'EBITDA': ('EBITDA PRE', 'EBITDA POST'),
    'Margin': ('Margin PRE', 'Margin POST'),
    'ROA': ('ROA PRE ', 'ROA POST '),
    'ROE': ('ROE PRE ', 'ROE POST '),
    'Revenue': ('Revenue PRE ', 'Revenue POST'),
    'Debt to Equity': ('D/E PRE ', 'D/E POST')
}

# Create a High Growth Dummy variable (1 if High Growth, 0 if Low Growth)
df['High Growth Dummy'] = df['Growth sector'].apply(lambda x: 1 if x == 'High' else 0)


# Function to perform regression analysis on each metric
def run_regression(metric, data):
    pre_col, post_col = metrics[metric]
    data[f'{metric} Change'] = data[post_col] - data[pre_col]

    # Independent variable: High Growth Dummy
    # Control variables: ROE PRE, ROA PRE, D/E PRE, Revenue PRE (you can adjust this as needed)
    control_vars = ['ROE PRE ', 'ROA PRE ', 'D/E PRE ', 'Revenue PRE ']
    X = data[['High Growth Dummy'] + control_vars]  # Independent variables
    X = sm.add_constant(X)  # Adds a constant term to the regression

    # Dependent variable: Change in the financial metric
    Y = data[f'{metric} Change']  # Dependent variable

    # Run the regression
    model = sm.OLS(Y, X).fit()

    # Save regression summary as HTML table
    save_summary_as_html_table(model, metric)


# Function to save regression summary as an HTML table
def save_summary_as_html_table(model, metric):
    summary = model.summary()

    # Extract regression results as tables
    coefficients_table = summary.tables[1]  # Coefficients table
    other_stats_table = summary.tables[0]  # Other statistics

    # Prepare HTML output
    html_text = f"<h2>Regression Results for {metric}</h2>\n"

    # Coefficients table
    html_text += "<h3>Coefficients Table:</h3>\n"
    html_text += convert_table_to_html(coefficients_table)

    # Other statistics (like R-squared, F-statistic, etc.)
    html_text += "<h3>Other Statistics:</h3>\n"
    html_text += convert_table_to_html(other_stats_table)

    # Save the HTML output to a file
    with open(f'{metric}_regression_summary.html', 'w') as f:
        f.write(html_text)


# Function to convert statsmodels table into well-formatted HTML
def convert_table_to_html(table):
    html = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse: collapse;'>\n"

    # Split the table into rows
    rows = table.as_html().split('<tr>')[1:]  # Skip the first split item, which is before the first <tr>

    # Process each row
    for row in rows:
        row = row.replace('</td>', '</td>').replace('<th>', '<th style="text-align: left;">').strip()
        html += f"<tr>{row}</tr>\n"

    html += "</table>\n"
    return html


# Run regression for all financial metrics
for metric in metrics.keys():
    run_regression(metric, df)