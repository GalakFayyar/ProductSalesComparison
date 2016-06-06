# Product Sales comparison
## Presentation
Developed to help stored and vendors learn about stock statistics, this tool is developed to make comparisons of multiple file, as far as files have similar structure. It has been build to work with Microsoft Excel files. The best example is to imagine Excel files representing monthly sales ; each month, we produce file like this detailing the sales for a bunch of products and for each VTA ratio. With the tool, we can easily compare and measure the evolution of each product on a period like a year or a quarter. The operation is fast and precise with no errors possible.
## Procedure
The main source code is developed with Microsoft Visual Studio in Visual Basic with 3.5 .Net Framework. It also include Interop drivers to ensure the operations on Microsoft Excel files. The key of the efficientness for this application is to consider an Excel Workbook as a Database. Each Worksheet becomes a table and each column ... is a column ! Then getting data from a Excel sheet is simple as a 
```sql
USE Workbook1;
SELECT * FROM Sheet1;
```
to automaticaly get the right data (only data, no need to ensure we reach the end of the document).
## Structure
To operate, a simple graphic interface has been developed and contains the minimum requirement elements such as the browsing field to specify the folder containing all Excel files, a multi combobox area in order to display all the Excel files contained in the current Folder and a progress bar with a status label to show and explain the operation status.
