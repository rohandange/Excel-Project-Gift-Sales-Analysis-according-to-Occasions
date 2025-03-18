# Excel-Project-Gift-Sales-Analysis-according-to-Occasions

This Excel project focuses on analyzing **Gift Sales Performance** by occasions. The analysis leverages data from multiple tables, including customer, order, and product data, to assess revenue, customer spending, and sales trends across different occasions. The project includes data cleaning, data modeling using Power Query, creation of pivot tables, and a dynamic dashboard to visualize key insights for decision-making.

### **Project Overview**

The project involves the following steps to clean, analyze, and report sales data:

1. **Data Cleaning and Preprocessing**
    - The initial data comes from three tables: **Customer**, **Order**, and **Product**. Basic formatting and data type changes were applied to ensure consistency and accuracy.
    - The **Product Description** column in the Product table was removed for a streamlined analysis.
    - The dataset was converted from CSV to XLSX format to avoid the limitations of CSV files and to preserve all data formatting and formulas.

2. **Data Import and Transformation**
    - A new Excel sheet was created, and data from the **Orders** and **Products** files were imported using **Power Query**.
    - The following transformations were applied to the **Orders** table:
      - **Month_Name**: Extracted the month from the **Order Date** column.
      - **Hour (Order Time)**: Extracted the hour from the **Order Time** column.
      - **Hour (Delivery Time)**: Extracted the hour from the **Delivery Time** column.
      - **Diff_bet_delivery_order_dates**: A custom column created by subtracting **Order Date** from **Delivery Date** to measure the difference between the order and delivery time.
      - **Price**: Merged the **Orders** table with the **Products** table and extracted the price for each order item.
      - **Revenue**: Created a custom column by multiplying **Quantity** and **Price** to calculate revenue.

3. **Building the Data Model**
    - A **Pivot Table** was created with the data model option enabled to analyze the data effectively.
    - To define relationships between the tables, the **Data Model** was used to establish connections between the **Primary Key** and **Foreign Key** columns.

4. **Analysis and Reporting**
    - The following questions were addressed with pivot tables and calculated measures:
      - **Total Revenue**: Calculated the overall revenue generated.
      - **Average Order and Delivery Time**: Analyzed the average time taken for order fulfillment and delivery.
      - **Monthly Sales Performance**: Examined sales performance on a monthly basis.
      - **Top Products by Revenue**: Identified the top-selling products by total revenue.
      - **Customer Spending Analysis**: Analyzed the spending behavior of customers.
      - **Sales Performance by Top 5 Products**: Investigated the sales performance of the top 5 products.
      - **Top 10 Cities by Number of Orders**: Identified cities with the highest order volumes.
      - **Impact of Order Quantities on Delivery Time**: Analyzed if higher order quantities impact delivery times.
      - **Revenue Comparison Between Occasions**: Compared the revenue generated across different occasions.
      - **Product Popularity by Occasion**: Analyzed product popularity for each occasion.

5. **Measures and Calculations**
    - Custom measures were created for deeper insights, including:
      - **Sum of Revenue**
      - **Total Orders Placed**
      - **Average Customer Spending**
      - **Average Difference Between Order and Delivery Dates**

6. **Building the Dashboard**
    - A dynamic **Dashboard** was created using the pivot tables and charts generated above, providing a comprehensive view of the data and insights.
    - **Timelines** were inserted for **Order Date** and **Delivery Date** to allow for time-based filtering.
    - **Slicers** were added for the **Occasions** to filter the data dynamically.
    - The **Timeline** and **Slicers** were connected to all pivot tables and charts via the **Report Connection** option to enable interactive data exploration.

### **Key Features**
- **Data Import & Transformation**: Power Query was used to import, clean, and transform data from the Orders and Products tables.
- **Data Model & Relationships**: Defined relationships between tables to facilitate a more efficient and integrated analysis.
- **Advanced Pivot Tables & Measures**: Created pivot tables to answer key business questions and measures to track revenue, order performance, and customer behavior.
- **Interactive Dashboard**: Developed a visually rich dashboard with pivot charts, timelines, and slicers to provide interactive insights.
- **Dynamic Reporting**: Enabled dynamic reporting with slicers for occasions and timelines for orders and deliveries.

### **How to Use**
1. Open the Excel file in Microsoft Excel.
2. Use the **Timeline** and **Slicer** to filter the data by order date, delivery date, and occasion.
3. View the pivot tables and charts to analyze total revenue, customer spending, and sales performance by occasion, product, and region.
4. Explore the insights by interacting with the dynamic elements in the dashboard, such as the slicers and timelines.

### **Technologies Used**
- Microsoft Excel (Power Query, Pivot Tables, Pivot Charts, Slicers, Timelines)
- Data Modeling and Relationships
- Custom Measures and Calculations

### **Conclusion**
This project offers a comprehensive analysis of gift sales performance, providing actionable insights into customer spending, product popularity, and sales trends for various occasions. The dynamic dashboard allows for quick, interactive analysis, enabling decision-makers to optimize sales strategies and improve performance based on real-time data.
