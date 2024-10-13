## Rotary Club Project: Creating a Dynamic Dashboard for their Annual Budget

<p>
In sharing this project, I want to showcase some of my experience in not only visualizing data, but also handling real data through the general cleanup process, restructuring the data to be able to use for pivot tables, and taking feedback from the end-user (the New Rivery Rotary Club) to give them a dynamic dashboard that would accomodate all future values for their budget. I wanted the Rotary Club to have insights into which quarters drove the most income to fund their most expensive events throughout the year. Ultimately, the insights gained in this dashboard will ensure that the Rotary Club is allocating money and resources towards hosting fundraising events that have proven to be successful based on the difference between the projected and the actual, future values.
  
</p>

*NOTE: future values are randomized here, so the refresh button reflects these drastic changes in the visualizations. Scrolling through the table at the bottom also retriggers the RANDBETWEEN function for the actual values*
- Personally, I am not a fan of using this many pie charts, but I wanted to stay true to the client's input on design-choice, so I ordered the slices of the pie to go chronologically clockwise for the quarters of the year for interpretability purposes.

![Rotary Club Excel Dashboard](https://github.com/user-attachments/assets/d01ff6c1-dca6-4a30-82ae-686d66b1795f)

Some highlights from this project were:
1. Using the RANDBETWEEN(#,#) function to ensure that I had future values to build my dashboard around.
2. Incorporating conditional formatting in the Difference column of the budget sheet that was given to me, so the Rotary Club could more easily see patterns in their data.
3. Using VLOOKUP to restructure the data in a new worksheet. This new worksheet served as the source for my dashboard's pivotcharts.
4. Using Excel's Visual Basic Application (VBA), I added a Macro to my dashboard to ensure that the Rotary Club would be able to refresh the data across all visualizations and tables with the click of a button, instead of clicking the Refresh All button in the Data tab for every worksheet after each updated value.
5. I wanted there to be a scroll bar for the table at the bottom of my dashboard, only revealing 10 lines at a time. Since there are only 10 events that were listed as income sources for the year, this choice made sense for now and is why only the table under the "Quarterly Expenses Expected vs. Actual Values" section has a scroll bar.
6. Color is important to me when creating dashboards so I not only wanted the theme to match this Rotary Club's colors (blue and orange), but also make sure that my choice of color in the visualization titles facilitated their interpretation and was not just adding color for the sake of being colorful. Color-choice should be strategic and not detract for the functionality of the dashboard, but of course this should also align with the audience's preference and their level of comfortability in reading graphs.

### This is the spreadsheet that I was given by the Rotary Club:
*I made some changes, to include randomized data and conditional formatting in the Difference column*
![General Rotary Club Spreadsheet](https://github.com/user-attachments/assets/8bd173d5-36de-4394-a0fb-1a68e7a4e4e8)

### New worksheet with restructured data:
*Using VLOOKUP, I wanted to ensure that the event type is correctly split for events that are for-profit (labeled as Income) and all others, like donations, raffles, or scholarships (Expense)*
![Data Cleanup for Pivot Tables with VLOOKUP](https://github.com/user-attachments/assets/22f98894-a81b-4607-b3b7-d4c441846856)

### Next, I needed to create Pivot Charts that would be added to the dashboard:
*I wanted to split these into several worksheets for organization purposes and ensuring enough variety in visualizations*
![Pivot Charts for Dashboard](https://github.com/user-attachments/assets/5ac383e6-61d2-4c7e-8af4-bcc7ee8acade)

### Lastly, this is the Visual Basic Application (VBA) code used to create a Macro within the dashboard:
*The purpose of this macro was to refresh all visualizations with the click of a button to reflect any changes to the actual income and actual expenses columns in the budget spreadsheet*
![Macro RefreshAllPivotTables](https://github.com/user-attachments/assets/e4d4da12-fb0b-446e-a52b-eecf145283a1)
