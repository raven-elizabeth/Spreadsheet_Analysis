# Import modules
import matplotlib.pyplot as plt
import pandas as pd
import seaborn

# Kaggle dataset
customer_shopping_trends = pd.read_csv("shopping_trends.csv")

# Seaborn bar chart
seaborn_bar_chart = seaborn.barplot(x="Season", y="Purchase Amount (USD)", data=customer_shopping_trends)
plt.show()