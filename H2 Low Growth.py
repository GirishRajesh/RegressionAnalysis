import pandas as pd
import statsmodels.api as sm

# Define the path to your Excel file
file_path = 'Regression data for analsysis .xlsx'

# Load the dataset
df = pd.read_excel(file_path)

# Filter the data for low-growth sector
df_high_growth = df[df['Growth sector'] == 'High'].copy()  # Use .copy() to avoid SettingWithCopyWarning

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


# Function to run regression and save results
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

    # Print the regression summary
    print(f'Regression results for {metric} (high Growth Sector):')
    print(model.summary())


# Run regression for each metric in the low-growth sector
for metric in metrics.keys():
    run_regression(metric, df_high_growth)
