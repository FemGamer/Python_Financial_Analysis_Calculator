###Linear Regression model with Python

import pandas as pd
import matplotlib.pyplot as plt
import statsmodels.api as sm ##python library for functioning linear regression and correaltion coefficient

Ratios = "D:/PY_FemGamer/Thesis/04_Results/Regression.xlsx" ##calling combined summary stats dataset for calculating regression line
sheet_activate2= "Ratio" #summary stats of Ratios
sheet_activate1= "Stock"## summary stats of Stock
df= pd.read_excel(Ratios, sheet_name=sheet_activate2, index_col=0)
df2= pd.read_excel(Ratios, sheet_name=sheet_activate1, index_col=0)


Y = df2['Close']  ##Dependent variable Y which is the the standard deviation of closing price 

X = df # Independent variables

X = sm.add_constant(X) # Add a constant to the independent variables

# Fit the regression model, using the OLS model
model = sm.OLS(Y, X).fit() 

# The model summary
model_summary = model.summary()
model_summary_df = pd.read_html(model_summary.tables[1].as_html(), header=0, index_col=0)[0]
model_summary_df.to_excel("D:/PY_FemGamer/Thesis/04_Results/012_regression_results.xlsx")

#Visualize the results
#Ratios v/s Closing Price with regression line scatter grpah with regression line
plt.scatter(Y, model.predict(X), label='Data')
plt.plot(Y, model.predict(X), color='red', label='Regression Line')
plt.xlabel("Ratios")
plt.ylabel("Closing Value Std")
plt.title("Ratios v/s Closing Price with regression line")
plt.savefig("D:/PY_FemGamer/Thesis/04_Results/Regression_Line")  # Save the plot as an image
plt.show()


#####
###corelation coefficient

Ratios = "D:/PY_FemGamer/Thesis/04_Results/Regression.xlsx" ##calling combined summary stats dataset for correlation coefficitne
sheet_activate2= "Ratio"
sheet_activate1= "Stock"
df= pd.read_excel(Ratios, sheet_name=sheet_activate2, index_col=0)
df2= pd.read_excel(Ratios, sheet_name=sheet_activate1, index_col=0)


closing_std = df2['Close'].std() # # Dependent variable Y which is the the standard deviation of closing price 


print(Y)
# Calculate correlation coefficients
correlation_coefficients = df.corrwith(Y)

# Print the correlation coefficients
print("Correlation Coefficients:")
print(correlation_coefficients)

# Save to Excel
correlation_coefficients.to_excel("D:/PY_FemGamer/Thesis/04_Results/02_correlation_coefficients.xlsx")

# Visualize the results (scatter plot)
plt.scatter(Y, X.mean(axis=1))  # Plotting mean of Ratios v/s Closing Price Corelation Coefficient
plt.plot(Y, model.predict(X), color='red', label='Correalation Coefficient Line')
plt.xlabel("Ratios")
plt.ylabel("Std. of Closing Price")
plt.title("Ratios v/s Closing Price Corelation Coefficient")
plt.savefig("D:/PY_FemGamer/Thesis/04_Results/02_Correlation_Coefficient.png")  # Save the plot as an image
plt.show()
