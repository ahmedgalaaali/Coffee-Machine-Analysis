# Coffee Machine Data Analysis Using Excel
This project was created to showcase some of my Microsoft Excel skills in cleaning the data, show my insights in a dashboard presented in another Excel sheet in the same workbook, and then create a final report using Microsoft Word.
## Overview of the Project
- First, data in the format of a `.CSV` file was downloaded from Kaggle.com, called [Coffee Sales Analysis](https://www.kaggle.com/code/emilcollu/coffee-sales-analysis), data recorded by a coffee machine set up in someplace in Ukrain.
- The data was then loaded and cleaned in Microsoft Excel where multiple changes were made, such as adding aiding columns, using various Excel tools and formulas, creating pivot tables, and pivot charts and even manipulating the style of the main table.
- A covering report was then created to show the detailed insights that could be extracted from the row data. The report contains insights and further recommendations to enhance the service produced by the machine.
- Finally but not final, a simple interactive dashboard was created for quicker insight extraction by the stakeholder.
  - Notice that both the dashboard and the report used the same color pallet called "Espresso Brown" to put the stakeholder in the mode of "Coffee". Here is the pallet: `#4E342E`,`#8D6E63`,`#D7CCC8`, `#FFF3E0`
- Not to forget, I used some help from Python Pandas to create a simple statistical discription for the data as this could be a little bit time consuming if I used excel.
## Cleaning and Preparation
First of all, I had to format the range as a table to make the process faster with no need to manually apply the formulas and changes manually.
### Date and Time
- The first 2 columns were formatted as date and time columns, . had to extract the time in a separate column from the datetime column a,aste it t, andeformat it as a time to show it inbetter
- As long as the the datetime is formatted like `45352.42767`, the formula used to extract the fractional part *which is the time as well* is:
  ```
  =MOD(B2,1)
  ```
  then changed the data type from `General` to `Time`.
- Finally, I had to create a new column that separates the 24 hours of the day into 6 bins:
  - Early morning
  - Mid-morning
  - Afternoon
  - Evening
  - Night
  - Late night
- This step could help me later to show some useful insights that serve the analysis, this step was accomplished by using a nested `=AND()`, `=TIME()` into `=IF()` formula:
  ```
  =IF(AND(B2>=TIME(4,0,0), B2<TIME(9,0,0)), "Early Morning",
  IF(AND(B2>=TIME(9,0,0), B2<TIME(12,0,0)), "Mid-Morning",
  IF(AND(B2>=TIME(12,0,0), B2<TIME(17,0,0)), "Afternoon",
  IF(AND(B2>=TIME(17,0,0), B2<TIME(21,0,0)), "Evening",
  IF(AND(B2>=TIME(21,0,0), B2<TIME(24,0,0)), "Night",
  "Late Night")))))
  ```
  ### Text Manipulation
- For a better vesual of the data, I had to do some changes, first I had to show the payment method to look more professional and to be shown in the visualization in a better look, the formula used:
  ```
  =PROPER(D2)
  ```
- For the same reason of professionality, I had to change names of the columns:
  - `datetime` to `Time`
  - `cash_type` to `Payment Method`
  - `card` to `Card Code`
  - `money` to `Revenue`
  - `coffe_name` to `Coffee Name`
### Data Type changed
- According to the publisher and as mentioned above, the data was collected by a machine that sell coffee set up in someplace in **Ukrain**, so that it was important to change the type of the `revenue` from `General` to `Currency` the **Ukrainian Hryvnia**: `38.70 ₴`
## Data Analysis
- Multiple pivot tables were created to explore, analyze the data and extract the insights, the purpose of creating them was as below:
  - A PivotTable to show the rate of using Cards and cash payment methods. `Count of Payment Method`
  - A PivotTable to show the total revenue collected at the estimated period. `Sum of Revenue`
  - A PivotTable that summarizes the total revenue per month. `Month (date)`, `Sum of Revenue`
  - A PivotTable that shows the frequency of buying the coffee by name. `Coffee Name`, `Count of Coffee Name`
  - A PivotTable that shows the revenue collected by each coffee. `Coffe Name`, `Sum of Revenue`
  - A PivotTable to show the average price paied by the customers when buying coffee. `Coffee name`, `Average of Revenue`
  - A PivotTable that shows the loyalty of the customers of the machine. `Card Code`, `Count of Card Code`
- In another sheet at the same workbook, I created a sub-table extracted from the main table, it has `Coffee Name`, `Revenue`, `Total Revenue` extracted from a pre-created PivotTable using `=VLOOKUP()`
  ```
  =VLOOKUP('Price Revenue Chart'!A2,'Pivot Tables'!$P$4:$Q$11,2,FALSE)
  ```
  The purpose of this step is to see the relationship between the average prices of the coffee and the total revenue by:
  - Creating a **scatter plot** to show the relationship between the values.
  - Calculating the **correlation coefficient** to see whether there's a relationship or not.
- I also tended to get some help from Python especially Pandas to get a quicker statistical discription of the data, code:
  ```python
  import pandas as pd
  df = pd.read_csv("C:\Mine\GitHub Projects\Coffee Machine\coffee_machine_csv.csv")
  df.describe()
  ```
![image](https://github.com/user-attachments/assets/a802a1b9-2ce8-4984-9688-4c35008574a1)

## Data Reporting
**All of the mentioned insights are available in the `.pdf` file provided in the repo**
## Data Visualization
Despite the limited capabilities provided by Excel copared to Power BI, a good looking, well formated dashboard was created by:
- Collecting the **charts** created by the PivotTables in one area. Modified the look to fit the dashboard.
- Created **slicers**, made them connected all together and also modifyed the style to fit the dashboard.
  ![image](https://github.com/user-attachments/assets/9c2edaff-b925-48a5-9a04-46a759762efe)
