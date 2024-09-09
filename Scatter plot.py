import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import numpy as np

# Define the path to your Excel file
file_path = 'Regression data for analsysis .xlsx'

# Load the dataset
df = pd.read_excel(file_path)

# Remove any leading/trailing spaces from the column names
df.columns = df.columns.str.strip()

# Create a High Growth Dummy variable (1 if High Growth, 0 if Low Growth)
df['High Growth Dummy'] = df['Growth sector'].apply(lambda x: 1 if x == 'High' else 0)

# Create the Revenue Change column
df['Revenue Change'] = df['Revenue POST'] - df['Revenue PRE']

# List of control variables
control_var = 'ROE PRE'

# Remove extreme outliers for a cleaner plot
df_filtered = df[(df['Revenue Change'] < 1e6) & (df['ROE PRE'] < 100)]  # Filter to remove extreme values

# Plot scatter points
plt.figure(figsize=(6, 4))
plt.scatter(df_filtered[control_var], df_filtered['Revenue Change'], color='blue', label='Data points')

# Perform linear regression for the line of best fit
X = sm.add_constant(df_filtered[control_var])  # Add constant term to the regression model
model = sm.OLS(df_filtered['Revenue Change'], X).fit()

# Plot the regression line
x_vals = np.linspace(df_filtered[control_var].min(), df_filtered[control_var].max(), 100)
y_vals = model.params.iloc[0] + model.params.iloc[1] * x_vals  # Calculate the line points
plt.plot(x_vals, y_vals, color='red', linestyle='--', label=f'Fit line: y={model.params[1]:.2f}x+{model.params[0]:.2f}')

# Set a custom y-axis limit to zoom in on the data and make the line steeper
plt.ylim(-1e5, 1e5)

# Add labels and title
plt.title(f'Revenue Change vs {control_var} ')
plt.xlabel(control_var)
plt.ylabel('Revenue Change')
plt.legend()

# Save the plot as a PNG file
plt.savefig(f'Revenue_Change_vs_{control_var}_zoomed.png')

# Show the plot
plt.show()
