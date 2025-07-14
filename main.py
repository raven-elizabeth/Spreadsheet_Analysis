# Please note: This program/code still needs refinement, particularly for Single Responsibility

# Import modules
import csv
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.patches as m_patches
import pandas as pd
from collections import Counter
from time import sleep

class CSV:
    # Open and read from CSV
    def __init__(self, file_name):
        with open(f"{file_name}", "r") as file:
            # Convert to list so data can be accessed once file has closed
            self.spreadsheet = list(csv.DictReader(file))


class Sales(CSV):
    def __init__(self, csv_file):
        super().__init__(csv_file)
        self.data = {}
        self.years = []
        self.current_year = None
        self.month_order = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
        self.total_sales = 0
        self.total_expense = 0
        self.total_profit = 0
        self.add_data()

    # Retrieve data from spreadsheet and sort into dictionary of dictionaries
    def add_data(self):
        for row in self.spreadsheet:
            year = row["year"]
            month = row["month"]
            sales = int(row["sales"])
            expense = int(row["expenditure"])

            if year not in self.data:
                self.data[year] = {}

            self.data[year][month] = {
                "sales": sales,
                "expense": expense,
                "profit": sales - expense,
                "percentage_change": 0
            }

    def get_months(self, year):
        return sorted(self.data[year].keys(), key=lambda x: self.month_order.index(x))

    def get_average(self, year):
        monthly_sales = [self.data[year][month]["sales"] for month in self.data[year]]
        # manual_avg = round(total_sales / 12, 2)
        numpy_avg = round(np.mean(monthly_sales), 2)
        return numpy_avg # Average using NumPy

    def get_annual_max(self, year):
        sales = self.get_pd_series_by_month(year, "sales")
        return sales.idxmax()

    def get_annual_min(self, year):
        sales = self.get_pd_series_by_month(year, "sales")
        return sales.idxmin()

    def display_written_data(self):
        self.get_user_choice("year", self.data.keys())  # sales, expense, profit, average, percentage changes in sales & profit

    @staticmethod
    def display_options(question, valid_options):
        print(f"Choose a {question}:")
        for option in valid_options:
            print(option)

    def get_user_choice(self, question, valid_options):
        self.display_options(question, valid_options)
        method = getattr(self, f"show_{question}ly")
        ans = input("").lower()

        if ans in valid_options:
            method(ans) #self.show_yearly(ans)
        elif ans == "exit":
            raise SystemExit
        else:
            print(f"Enter a given {question} or exit the program by entering 'exit'")
            self.get_user_choice(question, valid_options)

    def show_yearly(self, year):
        self.current_year = year
        total_sales, total_expense, total_profit = self.get_yearly_data(year)
        self.total_sales += total_sales
        self.total_expense += total_expense
        self.total_profit += total_profit

        average_sales = self.get_average(year)
        print(f"Year: {year}\nSales: {total_sales}\nExpense: {total_expense}\nProfit: {total_profit}\nAverage Sales: {average_sales}\n")

    def show_data_by_month(self):
        year = self.current_year
        self.get_user_choice("month", self.data[year].keys())

    def show_monthly(self, month):
        year = self.current_year
        sales = self.data[year][month]["sales"]
        expense = self.data[year][month]["expense"]
        profit = self.data[year][month]["profit"]
        sale_pct_change, profit_pct_change = self.calc_perc_change(year)


        summary = f"Month: {month.title()}\nSales: {sales}\nExpense: {expense}\nProfit: {profit}\nSale Change Compared to Previous Month (%): {float(sale_pct_change[month])}\nProfit Change Compared to Previous Month (%): {float(profit_pct_change[month])}\n"
        print(summary)

    def get_yearly_data(self, year):
        # Get only sales & expense data for each month using .values()
        monthly_data = self.data[year].values()

        # Convert each month's data into separate counters so they can be added together
        monthly_counters = (Counter(month) for month in monthly_data)

        # Have sum begin with an empty Counter and add to it to avoid errors adding to the default integer 1
        yearly_data = sum(monthly_counters, Counter())

        return yearly_data["sales"], yearly_data["expense"], yearly_data["profit"]

    def get_pd_series_by_month(self, year, data_type):
        data = [self.data[year][month][data_type] for month in self.data[year].keys()]
        return pd.Series(data, index=self.data[year].keys())

    def calc_perc_change(self, year):

        # Series is a one-dimensional array, unlike DataFrame which is a two-dimensional table with rows.columns
        # index parameter allows items to be labelled by month (default starts at 0)
        sales_per_month = self.get_pd_series_by_month(year, "sales")
        profit_per_month = self.get_pd_series_by_month(year, "profit")

        # pct_change() gives difference as a fraction, so must be multiplied to get percentage
        sales_pct_change = round(sales_per_month.pct_change() * 100, 2)
        profit_pct_change = round(profit_per_month.pct_change() * 100, 2)
        return sales_pct_change, profit_pct_change

    def display_annual_graphs(self):
        #TEMPORARY
        year = "2018"
        graph = Graph(self.clean_graph_data(year))
        graph.line_graph()
        graph.scatter_graph()
        graph.bar_charts()
        # Show and label all diagrams
        plt.suptitle(f"{year} Sales")
        plt.show()
        graph.pie_chart()

    def clean_graph_data(self, year):
        sales_change, profit_change = self.calc_perc_change(year)
        clean_data = {
            "year": year,
            "months": self.get_months(year),
            "sales": [self.data[year][month]["sales"] for month in self.get_months(year)],
            "expense": [self.data[year][month]["expense"] for month in self.get_months(year)],
            "profit": [self.data[year][month]["profit"] for month in self.get_months(year)],
            "sales_change": sales_change,
            "profit_change": profit_change,
            "avg_sales": self.get_average(year),
            "max_sales": self.get_annual_max(year),
            "min_sales": self.get_annual_min(year)
        }

        return clean_data

    def create_summary_excel_file(self):
        # TEMPORARY
        year = "2018"
        # Summarise data list results in a table using pandas
        data = self.clean_graph_data(year)

        insights = pd.DataFrame(
            {
                # Use [] to indicate single-row DataFrames with no index needed (will raise error otherwise)
                "Average Sales": [self.get_average(year)],
                "Highest Sale Month": [self.get_annual_max(year)],
                "Lowest Sale Month": [self.get_annual_min(year)]
            }
        )

        overview = pd.DataFrame(
            {
                "Sales": data["sales"],
                "Expenditure": data["expense"],
                "Profit": data["profit"],
                "Sales Change %": data["sales_change"],
                "Profit Change %": data["profit_change"]
             },
            index=data["months"],

        )
        print(overview)

        # Add to Excel file as separate tables
        with pd.ExcelWriter("sales_summary.xlsx") as excel_file:
            insights.to_excel(excel_file, sheet_name="Sales Insights", index=False)
            overview.to_excel(excel_file, sheet_name="Sales Overview")



class Graph:
    def __init__(self, data_dict):
        self.data = data_dict

    def line_graph(self):
        # Line graph
        # Set X-axis values for line graph
        x_axis_month = np.array(self.data["months"])

        # Set Y-axis values for line graph
        y_axis_sales = np.array(self.data["sales"])
        y_axis_expenditure = np.array(self.data["expense"])

        # Subplot enables multiple graphs to be shown at once (row, column, position)
        plt.subplot(2, 2, 1)
        # "s" plots the points as separate squares rather than drawing a line, using marker keeps the line, mfc and mec sets the colour of the points
        plt.plot(x_axis_month, y_axis_sales, label="Sales", color="g", marker="s", mfc="black", mec="black")
        # Change line style with ls
        plt.plot(y_axis_expenditure, label="Expenses", color="r", ls="--")
        # Make a key for multiple lines
        plt.legend(loc='lower right')
        plt.title("Sales per Month (line graph)")
        plt.xlabel("MONTH")
        plt.ylabel("GBP")
        plt.grid()

    def scatter_graph(self):
        plt.subplot(2, 2, 2)
        y_axis_profit = np.array(self.data["profit"])
        plt.scatter(self.data["months"], y_axis_profit, marker="x", color="black")
        plt.title("Profit per Month (scatter graph)")
        plt.xlabel("MONTH")
        plt.ylabel("PROFIT")

    def bar_charts(self):
        # Find highest/lowest sale months
        sales_bar_colours = []
        for sale in self.data["sales"]:
            if sale == self.data["min_sales"]:
                sales_bar_colours.append("purple")
            elif sale == self.data["max_sales"]:
                sales_bar_colours.append("lightblue")
            else:
                sales_bar_colours.append("blue")

        # Bar chart for sales
        plt.subplot(2, 2, 3)
        plt.bar(self.data["months"], self.data["sales"], color=sales_bar_colours)
        plt.title("Sales Per Month (bar chart)")
        plt.xlabel("MONTH")
        plt.ylabel("SALES")
        # Create line on bar chart to represent average
        plt.axhline(y=self.data["avg_sales"], color="black", ls="--", label="Average")

        # Create labels for alternative colours through patch
        highest_month = m_patches.Patch(color="lightblue", label="Highest Sales")
        lowest_month = m_patches.Patch(color="purple", label="Lowest Sales")
        # Access labels automatically added to handles
        handles, labels = plt.gca().get_legend_handles_labels()
        # Add manual labels
        handles.extend([highest_month, lowest_month])
        plt.legend(handles=handles, loc="center right")

        # Bar chart for profit
        # Make months with a loss red and months with a profit green, black for breaking even
        profit_bar_colours = []
        for profit in self.data["profit"]:
            if profit < 0:
                profit_bar_colours.append("red")
            elif profit == 0:
                profit_bar_colours.append("black")
            else:
                profit_bar_colours.append("green")

        plt.subplot(2, 2, 4)
        plt.bar(self.data["months"], self.data["profit"], color=profit_bar_colours)
        plt.title("Profit per Month (bar chart)")
        plt.xlabel("MONTH")
        plt.ylabel("PROFIT")
        # Create line on bar chart to expose months with a loss
        plt.axhline(y=0, color="r", ls="--")

    def pie_chart(self):
        # Pie chart for profit months Vs sales months - find percentages
        gain = 0
        loss = 0
        for profit in self.data["profit"]:
            if profit > 0:
                gain += 1
            elif profit < 0:
                loss += 1

        total = gain + loss
        convert_to_percent = 100 / total
        loss_percent = round(loss * convert_to_percent)
        profit_percent = round(gain * convert_to_percent)

        plt.pie(
            [loss_percent, profit_percent],
            labels=[
                "Months with net loss ({}%)".format(loss_percent),
                "Months with net profit ({}%)".format(profit_percent)
            ],
            startangle=90,
            colors=["r", "g"]
        )
        plt.title(f"{self.data["year"]} Profit/Loss")
        plt.show()


class Program:
    def __init__(self, sales_csv):
        self.report = Sales(sales_csv)

    def ask_user_to(self, result):
        question = result.replace("_", " ")
        ans = input(f"Would you like me to {question}? (y/n) ").lower()
        if ans == "yes" or ans == "y":
            # Use getattr() to find the value of the chosen method (result) from the class (self.report), allowing it to be called in this way for reusability
            method = getattr(self.report, result)
            method()
        elif ans == "exit":
            raise SystemExit
        elif ans == "no" or ans == "n":
            return
        else:
            print("Invalid input. Please enter 'yes', 'no', or 'exit' to end the program.")
            self.ask_user_to(result)

    @staticmethod
    def message(msg):
        print("\n", f" {msg} ".center(100, "*"))

    def app(self):
        self.message("Welcome to the Spreadsheet Analysis tool")
        sleep(2)
        print("\nReading your sales data...\n")
        sleep(3)
        self.ask_user_to("display_written_data")
        self.ask_user_to("show_data_by_month")
        self.ask_user_to("create_summary_excel_file")
        self.ask_user_to("display_annual_graphs")
        self.message("Thank you for using this service")


def run():
    program = Program("sales.csv")
    program.app()


if __name__ == "__main__":
    run()