import matplotlib
import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_excel('tmp.xlsx')

def func1():
    # Print 10 rows
    print(df.head(10))

    rows = len(df.axes[0])
    cols = len(df.axes[1])
    # Print the number of rows and columns
    print("Number of Rows: " + str(rows))
    print("Number of Columns: " + str(cols))

def func2():
    # Print unique cities
    unique_cities = df.City.unique()
    print(unique_cities)

def func3():
    sum_products = df.groupby('Category')['ProductID'].sum()
    print(sum_products)

def func4():
    ar_mean = df.groupby('Category')['UnitsInStock'].mean()
    print(ar_mean)

def func5():
    max_price_row = df.loc[df['UnitPrice (Products)'].idxmax()]
    product_name = max_price_row['Product']
    product_category = max_price_row['Category']
    max_price = max_price_row['UnitPrice (Products)']
    print(f"The most expensive product: {product_name}, Category: {product_category}, Price: {max_price}")

def func6():
    sum_profit = df.groupby('Country')['Profit'].sum()
    print(sum_profit)


def func7():
    country_sales = df.groupby("Country")["Sales"].sum()
    country_sales.plot(kind="bar")
    plt.title("Sales by country")
    plt.ylabel("Sales")
    plt.xlabel("Country")
    plt.show()

    max_sales_country = country_sales.idxmax()
    max_sales_value = country_sales.max()

    print(f"Country with the highest sales: {max_sales_country} {max_sales_value}")

def func8():
    # Print unique city and countries
    unique_rows = df['City and Counry'].unique()
    print(unique_rows)

def func9():
    ar_mean = df.groupby('Category')['UnitCost'].mean()
    print(ar_mean)

def func10():
    largest_3 = df.groupby('Customer Company')['Profit'].sum().nlargest(3)
    print(largest_3)

def func11():
    condition_is_less_30 = df.query('Quantity > 30')
    print( condition_is_less_30)

def func12():
    largest_5 = df.groupby('Product')['UnitsOnOrder'].sum().nlargest(5)
    print(largest_5)

def func13():
    sum_quantity = df.groupby('Customer')['Quantity'].sum()
    print(sum_quantity)

def func14():
    counter = df.groupby('City')['Category'].nunique()
    print(counter)

def func15():
    df['Margin'] = df['UnitPrice'] - df['UnitCost']
    print(df)
    max_margin = df["Margin"].idxmax()
    print("Maximum margin: \n{}".format(df.loc[max_margin]))

def func16():
    for category, group in df.groupby("Category"):
        plt.scatter(group["Quantity"], group["Profit"], label=category)
    plt.title("Profit from the number of goods sold by category")
    plt.xlabel("Number of items sold")
    plt.ylabel("Profit")
    plt.legend()
    plt.show()

def func17():
    df1 = df.groupby("Customer")["UnitPrice (Products)"].agg(
        ["mean", "median"]
    )
    print(df1)
    max_mean = df1['mean'].idxmax()
    print(f"The customer with the highest average cost:\n {df1.loc[max_mean]}")

def func18():
    print("task_18")
    print("\n")

    df["Month"] = df["OrderDate"].dt.to_period("M")
    sales_by_month = df.groupby("Month")["Sales"].sum()
    sales_by_month.plot()
    plt.title("Sales by month")
    plt.xlabel("Month")
    plt.ylabel("Sales")
    plt.show()

def func19():
    unique_rows = (df.groupby('Customer Company')["Customer"].nunique().reset_index().rename(columns={"Customer": "UniqueCustomerCount"}))
    max_customer_company = unique_rows.loc[unique_rows["UniqueCustomerCount"].idxmax()]
    print(
        f"The company with the largest number of different clients: {max_customer_company['Customer Company']}, "
        f"Number of unique clients: {max_customer_company['UniqueCustomerCount']}"
    )
    print(unique_rows)

def func20():
    print(df.loc[df.groupby("Category")['Quantity'].idxmax()])

def func21():
    pivot = df.pivot_table(index=['Category'], values=['Profit'], aggfunc='sum')
    pivot1 = df.pivot_table(index=['City'], values=['Profit'], aggfunc='sum')
    print(pivot,"\n", pivot1)

def func22():
    df["OrderDate"] = pd.to_datetime(
        df["OrderDate"], errors="coerce"
    )
    last_6_months = df[
        df["OrderDate"] >= (df["OrderDate"].max() - pd.DateOffset(months=6))
        ]
    top_cities = last_6_months.groupby("City")["Sales"].sum().nlargest(3)
    print("Top 3 cities with the highest sales over the past 6 months:")
    print(top_cities)

def func23():
    mean_profit = df.groupby("Category")["Profit"].mean()
    df["excess_profit"] = df["Profit"] - df["Category"].map(mean_profit)
    top_5_exceeding = df.nlargest(5, "excess_profit")
    print("Top 5 Posts with the Largest Excess of Average Profit:")
    print(top_5_exceeding)





