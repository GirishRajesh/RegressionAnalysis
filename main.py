import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

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
    print(f'Regression results for {metric}:')
    print(model.summary())

    # Save regression summary as a PNG
    save_summary_as_png(model, metric)

    return model


# Function to save regression summary as a PNG
def save_summary_as_png(model, metric):
    # Convert summary to string
    summary_str = model.summary().as_text()

    # Create a figure
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.subplots_adjust(left=0.2, right=0.8, top=0.8, bottom=0.2)
    ax.axis('off')  # No axes, text-only figure

    # Display the summary text
    ax.text(0.01, 0.99, summary_str, transform=ax.transAxes, fontsize=10, verticalalignment='top')

    # Save the figure as a PNG file
    plt.savefig(f'{metric}_regression_summary.png', bbox_inches='tight', dpi=300)
    plt.close()


# Run regression for all financial metrics
for metric in metrics.keys():
    run_regression(metric, df)
