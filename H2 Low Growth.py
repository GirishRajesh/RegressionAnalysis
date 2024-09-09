import pandas as pd
import statsmodels.api as sm

# Define the path to your Excel file
file_path = 'Regression data for analsysis .xlsx'

# Load the dataset
df = pd.read_excel(file_path)

# Filter the data for low-growth sector
df_high_growth = df[df['Growth sector'] == 'Low'].copy()  # Use .copy() to avoid SettingWithCopyWarning

# Create the interaction term (no need to multiply by anything because it's a single variable)
df_high_growth.loc[:, 'Interaction'] = df_high_growth['GDP GROWTH YOY%']  # Use .loc to avoid SettingWithCopyWarning

# Define the metrics for regression
metrics = {
    'Netincome': ('Netincome PRE', 'Netincome POST'),
    'EBITDA': ('EBITDA PRE', 'EBITDA POST'),
    'Margin': ('Margin PRE', 'Margin POST'),
    'ROA': ('ROA PRE ', 'ROA POST '),
    'ROE': ('ROE PRE ', 'ROE POST '),
    'Revenue': ('Revenue PRE ', 'Revenue POST'),
    'Debt to Equity': ('D/E PRE ', 'D/E POST')
}


# Function to run regression and save results as HTML
def run_regression(metric, data):
    pre_col, post_col = metrics[metric]

    # Safely create the 'Change' column using .loc
    data.loc[:, f'{metric} Change'] = data[post_col] - data[pre_col]

    # Independent variables: GDP Growth and Interaction (basically just GDP GROWTH YOY%)
    X = data[['GDP GROWTH YOY%', 'Interaction']]
    X = sm.add_constant(X)  # Adds a constant term to the regression

    # Dependent variable: Change in the financial metric
    Y = data[f'{metric} Change']

    # Run the regression
    model = sm.OLS(Y, X).fit()

    # Save regression summary as HTML table
    save_summary_as_html_table(model, metric)


# Function to save regression summary as an HTML table
def save_summary_as_html_table(model, metric):
    summary = model.summary()

    # Extract regression results as tables
    coefficients_table = summary.tables[1]  # Coefficients table
    other_stats_table = summary.tables[0]    # Other statistics

    # Prepare HTML output
    html_text = f"<h2>Regression Results for {metric} (Low Growth Sector)</h2>\n"

    # Coefficients table
    html_text += "<h3>Coefficients Table:</h3>\n"
    html_text += convert_table_to_html(coefficients_table)

    # Other statistics (like R-squared, F-statistic, etc.)
    html_text += "<h3>Other Statistics:</h3>\n"
    html_text += convert_table_to_html(other_stats_table)

    # Save the HTML output to a file
    with open(f'{metric}_regression_summary_H2_low.html', 'w') as f:
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


# Run regression for each metric in the low-growth sector
for metric in metrics.keys():
    run_regression(metric, df_high_growth)