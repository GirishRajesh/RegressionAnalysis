import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import numpy as np

# Define the path to your Excel file (just the file name since it's in the same directory)
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

# Create the interaction term between High Growth Dummy and GDP Growth
df['Interaction'] = df['High Growth Dummy'] * df['GDP GROWTH YOY%']


# Function to perform regression analysis on each metric, with interaction term
def run_regression_with_interaction(metric, data):
    pre_col, post_col = metrics[metric]
    data[f'{metric} Change'] = data[post_col] - data[pre_col]

    # Independent variables: High Growth Dummy, GDP Growth, and the interaction term
    control_vars = ['High Growth Dummy', 'GDP GROWTH YOY%', 'Interaction', 'ROE PRE ', 'ROA PRE ', 'D/E PRE ',
                    'Revenue PRE ']  # Adjust according to your actual control variables
    X = data[control_vars]  # Independent variables
    X = sm.add_constant(X)  # Adds a constant term to the regression

    # Dependent variable: Change in the financial metric
    Y = data[f'{metric} Change']  # Dependent variable

    # Run the regression
    model = sm.OLS(Y, X).fit()
    print(f'Regression results for {metric}:')
    print(model.summary())

    # Save the regression summary as a PNG
    save_summary_as_png(model, metric)

    return model


# Function to save regression summary as a PNG
def save_summary_as_png(model, metric):
    summary_str = model.summary().as_text()

    # Create a figure to hold the summary text
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.subplots_adjust(left=0.2, right=0.8, top=0.8, bottom=0.2)
    ax.axis('off')  # Turn off axis

    # Display the summary text
    ax.text(0.01, 0.99, summary_str, transform=ax.transAxes, fontsize=10, verticalalignment='top')

    # Save the figure as a PNG
    plt.savefig(f'{metric}_regression_with_interaction_summary.png', bbox_inches='tight', dpi=300)
    plt.close()
# Run regression for all financial metrics with interaction term
for metric in metrics.keys():
    run_regression_with_interaction(metric, df)

# Optional: Save or plot as needed
