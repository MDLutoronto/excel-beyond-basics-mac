---
title: Pivot Tables and Pivot Charts
parent: "Excel: Beyond The Basics - Mac"
layout: default
nav_order: 5
---

## Pivot Tables and Pivot Charts

1. Pivot Tables and Pivot Charts are another way to graphically represent and analyze your data.

2. A **Pivot Table** allows you to make sense of a large, detailed data set. It’s a summary of your data that lets you group your data in different ways and explore trends in your information. Pivot tables are particularly useful if you have long rows or columns that hold values you need to track the sums of and easily compare to one another. They can help you draw helpful conclusions more easily, compared with writing a bunch of individual formulas

3. While a standard chart is linked directly to worksheet cells, **Pivot Charts** are based on their associated Pivot Table's data source.

4. The "pivot" part of a pivot table stems from the fact that you can rotate (or pivot) the data in the table to view it from a different perspective. To be clear, you're not adding to, subtracting from, or otherwise changing your data when you make a pivot. Instead, you're simply reorganizing the data so you can reveal useful information from it.

5. For example, we may want to use a pivot table to analyse our **2024_Bus_Delays** data. Let’s say we want to easily understand which routes have the most delays overall, or in a given month or day of the week. We might also want to see how much individual amounts—such as a single incident type — contribute to a total amount— such as the total number of incidents resulting in TTC bus delays.

6. Let’s highlight the cells we’ll need to create our pivot table and chart. **Select all of Columns A through to H** in the **2024_Bus_Delays** worksheet. Often with pivot tables, you'll want to sort your data in some way so it's easier to manage once you turn it into a pivot table. In our case, our data is already sorted by date so there’s no need to resort.

    <img src='{{ '/assets/images/M6_PivotTablesandCharts.png' | relative_url }}' alt='Selection of columns A to H' title='' width='1386' height='582' />

7. From the Insert Ribbon, select **Pivot Chart**. This will automatically create an interactive chart to accompany our table, that will change alongside our table as we reorganize our data.

    <img src='{{ '/assets/images/M7.png' | relative_url }}' alt='Pivot chart buttons are shown' title='' width='2178' height='972' />

8. Note you could also create the table only by selecting **Pivot Chart** or **Recommended Charts** from the **Insert** ribbon. The recommended option will preview some analysis possible with Pivot tables, and pre-arrange your data to present a chosen analysis as a starting point. But today, let’s start by creating a table from scratch, and one that comes with our visualization.

9. You’ll be asked to confirm your data range, which should match your selection. You’ll also be asked to choose where you want the PivotTable to be placed. Select “New Worksheet” and click **OK**.

    <img src='{{ '/assets/images/M9.png' | relative_url }}' alt='Context menu for pivot tables is shown' title='' width='395' height='294' />

    | **Note:** If you're using an earlier version of Excel, "PivotTables" and “PivotCharts” may be under the Tables or Data ribbon, rather than Insert. In Google Sheets, you can create pivot tables from the Data dropdown along the top navigation. |

 

10. Excel has now created a blank Pivot Table in a new sheet and the start of our workbook.

    <img src='{{ '/assets/images/M10_0.png' | relative_url }}' alt='Displaying of the blank pivot chart' title='' width='887' height='505' />

11. At the very right of the window, you should see a section entitled **PivotChart Fields**, which contains all the column titles for the data we had selected from the **2024_Bus_Delays** worksheetto create our Pivot Table and Chart. This is the area where we’ll work with the data to analyze it in our Table and Chart.

    <img src='{{ '/assets/images/M11.png' | relative_url }}' alt='Pivot Chart Fields context menu shown' title='' width='341' height='400' />

12. With your pivot table selected, let’s take a look at some additional areas of our PivotTables Field List. Any fields in your list can be dragged and dropped into any of the 4 areas or sections OR can be selected. If you select a field, Excel will by default add non-numeric fields to the **Rows** area, and numeric fields to the **Columns** Area.

13. Click on the Pivot Chart, to select it. You’ll note that when a Pivot Chart is selected, the area names change — **Rows** becomes **Axis (Categories)** and **Columns** becomes **Legend (Series)**.

    <img src='{{ '/assets/images/M13.png' | relative_url }}' alt='Pivot Chart filters, axis, legend, and values page is shown' title='' width='338' height='333' />

14. We can place a field into the **∑ Values** area to summarize our PivotTable by that field (by default it uses the =sum function).

15. The **Filters** area can be used if you want to provide an option to filter your data by any given field without adding it to your table.

16. You’ll also notice two new tabs appear on the ribbon when your Pivot Table is selected, marked as **Analyze** and **Design**. There is also a **Format** tab when your Pivot Chart is selected (These will only show when you have your table or chart selected; if you click elsewhere in the worksheet they will disappear).

17. Let’s say we want to determine which bus routes have the most delays. Select **Date** from the Field List, and drag it into the Rows section.

    <img src='{{ '/assets/images/M17.png' | relative_url }}' alt='Adding date to the rows of the pivot chart' title='' width='342' height='634' />

18. We now have a quick view of all the months we have in our data. You’ll notice that this data is now represented on both our table and chart.

    <img src='{{ '/assets/images/M18_0.png' | relative_url }}' alt='Display of the row labels as dates' title='' width='111' height='239' />

19. You’ll also see that Excel has created a **Month** Field from our Date Column. This is an automatically calculated summary field that is now available for you to use in the PivotTable.

20. Since we prefer to analyze the data by month instead of date, let’s remove “Date” from our Rows area, leaving only the Month field that Excel has created for us. We can do this by either dragging the field out of the box, or clicking the drop down arrow beside the field name in Rows, and selecting “Remove Field”

    <img src='{{ '/assets/images/image171.png' | relative_url }}' alt='Context menu with Remove Field highlighted' title='' width='273' height='328' />

    | **Group by Dates**    
    If you would like to aggregate date data in other ways for analysis, you can group these by selecting any date on your Pivot Table. Two finger click, and select “Group”. Any aggregate that you select will be automatically added as a new field to your Field list, and to the area your field is currently in (for example, Rows). |

    <img src='{{ '/assets/images/M20_0.png' | relative_url }}' alt='Grouping the data together blocking together certain features, such as hours, days, months' title='' width='260' height='415' />

21. We also want to be able to analyze the day of the week in our report. **Take Day** and drag this into the Rows area. Now you can see the Month, and each of the days under that month.

    <img src='{{ '/assets/images/M21.png' | relative_url }}' alt='Expanding the Dates under Row labels to show each day of the week' title='' width='199' height='328' />

22. You can expand and collapse the different sections by clicking on the + - beside each, or expand or collapse all fields under the **PivotTable Analyze Ribbon.** Note that when you expand sections in your table, the data in your Chart will adjust as well!

    <img src='{{ '/assets/images/image177.png' | relative_url }}' alt='Pivot table analyse, and the text alignment that works with it' title='' width='527' height='142' />

23. **Drag** the **Route field** over to **Columns**. So now we have a blank table and chart containing Months, days and TTC bus routes.

    <img src='{{ '/assets/images/M23.png' | relative_url }}' alt='Display of the blank table and chart' title='' width='1568' height='1084' />

24. We have our x and y values (our axis), but we don’t have any data inside of them ie. we’re not summarizing our data by any value. Say we want to know how many minutes of delays were reported in a given month and on a given day of the week for each route - we want a sum or total of minute delays. Take **Min_Delay** and pull that into the **Values** box.

    <img src='{{ '/assets/images/M24_1.png' | relative_url }}' alt='Resultant chart and table' title='' width='1840' height='1116' />

 

25. What this has done is it’s set our table to sum by that Value, or the total number of minute delays. So, for example, we can quickly see that the #7 Bus had 1640 minutes of delays in January, and looks like Saturdays are not a great time to be relying on that route…

    <img src='{{ '/assets/images/M25_0.png' | relative_url }}' alt='Column shows varying numbers in column 7' title='' width='329' height='232' />

26. Our table and chart also seem to contain some values that aren’t helpful at all. You’ll notice on both our table and our chart that all the fields have downward arrows beside them. This will allow you to further sort or filter those values, exactly as we did earlier in this workshop. Select the downward area beside the **Row Labels** in the Pivot Table, or **Months** on your Pivot Chart. Unselect the values  “<1/1/2024” and “>11/1/2024”.

    <img src='{{ '/assets/images/M26_1.png' | relative_url }}' alt='Selecting the filters, removing certain values in the filter' title='' width='456' height='573' />

    | **Note:** Instead of using filters, you could also remove these values by removing them from the source data ie. our **2024_Bus_Delays** worksheet. After editing the sheet, you’ll need to select “Refresh” from the PivotTable/Chart Analyze Ribbon |

    <img src='{{ '/assets/images/M26_2.png' | relative_url }}' alt='Selecting refresh under the ribbon' title='' width='352' height='249' />

27. Next, let’s take a look at our chart - it’s also struggling to display this much information in a single view! Let’s help it out a bit.  Select the downward area beside the **Column Labels** on the Pivot Table, or the **Route** label in your Chart. **Select** “Value Filters” and “Top 10…”

    <img src='{{ '/assets/images/M27.png' | relative_url }}' alt='In the context menu select value filters and then top 10' title='' width='590' height='501' />

28. In the pop-up window that appears, choose top 10 Items by “Sum of Min Delay”.

    <img src='{{ '/assets/images/image191.png' | relative_url }}' alt='Finished filter settings have been selected' title='' width='419' height='234' />

29. Our Chart (and our Table) have now been filtered to show only the 10 Bus Routes with the most total minute delays in 2024. This is a lot more useful!

    <img src='{{ '/assets/images/M29_0.png' | relative_url }}' alt='Original filtering has been done, this is the result' title='' width='1694' height='1260' />

30. Let’s say, however, that we’re not interested in the total or sum of minute delays, but how many delays each route reports. So, for example, we don’t want to know that the #32 bus had 28076 minutes of delays - we want to know that it had 1175 delays in total.

31. Click the 'i' icon beside the **Sum of Min Delays** in the **Values box** to set the **Value** field.

    <img src='{{ '/assets/images/image195.png' | relative_url }}' alt='Value field settings is selected after clicking into the Sum of Min Delay box' title='' width='706' height='430' />

32. This presents us with a popup that contains a number of options. We can re-label our Value; Change what we’re summarizing our value by; or show our value in a different way such as a percentage. Under “Summarize Values By” **select Count**. **Click OK**.

    <img src='{{ '/assets/images/M32.png' | relative_url }}' alt='Summarize Values by and Count are highlighted' title='' width='313' height='331' />

33. Our Table and Chart have now been updated to show the Routes with the top 10 total number of incidents, which you’ll notice is quite a different result (although some of the same routes are still in our top 10).

    <img src='{{ '/assets/images/M33_0.png' | relative_url }}' alt='Chart result showing the correct outcome' title='' width='1694' height='1260' />

34. Let’s filter the data further to show the routes with the more delay incidents reported during the weekend (Sat, Sun) in the Winter of 2024 (Jan, Feb, March). Click the filter icon arrows on your chart to filter for only these data. Select **OK** for each filter when done. Our Chart and table is now only showing us the top 10 routes with the most incidents reported on these days.

    <img src='{{ '/assets/images/M34_0.png' | relative_url }}' alt='Context menu with only Sunday & Saturday selected' title='' width='294' height='362' />

35. Now, let’s add a couple of **filters** to our Table and Chart. This will allow anyone interacting with the chart to quickly filter the data, without us needing to add these fields to our chart axes. **Drag and drop** “Incident” into the Filters box.

36. The filter has appeared at the top of our chart. Now we can, for example, see only the top 10 bus routes impacted by “Diversion” delays.

    <img src='{{ '/assets/images/M36.png' | relative_url }}' alt='Incident box in the top left is highlighted, as well as the filter on the right side' title='' width='1796' height='770' />

37. Finally, let’s say we want to add a title to our Chart that includes this incident information. We want a dynamic title that reflects the incident type we are filtering for. We can do this using an IF formula!

38. **Select an empty cell** at the top of your worksheet below your table (for example, E1) and type the function below:

    ```
    ="TTC Weekend Delays, Incident Type: "&IF(B1="(All)", "All",IF(B1="(Multiple Items)", "Multiple", B1))
    ```

39. The formula starts with a text string, "TTC Weekend Delays, Incident Type: “. The ampersand operator ( & )  joins that string with the text that follows it.

40. Our nested IF formula ensures that the ending of our string depends on what is selected in the pivot table’s Region filter:

    | a. If ALL incident types are selected, the formula ends with “All” |
    | b. If multiple incident types are selected, the formula ends with “Multiple” |
    | c. If only one incident type is selected, that incident type, in cell B1, is shown. |

41. Once you’ve entered your formula, hit **Enter**. Our cell is now showing “Incident Type: All”

    <img src='{{ '/assets/images/M41_0.png' | relative_url }}' alt='Cell E1 with the resulting text of "TTC Weekend Delays, Incident type: all"' title='' width='657' height='104' />  
 
42. Next, let's add a chart title. From the **Design** ribbon, select **Add Chart Element**, **Chart Title**, **Above Chart**.

    <img src='{{ '/assets/images/M42.png' | relative_url }}' alt='Chart Title selection is highlighted' title='' width='490' height='226' />

43. Click on your Chart Title to adjust. With your chart titles selected, in the formulas bar **type** = and **select** the cell containing your formula. Hit **Enter**. Your Chart Title will now adjust as we select / filter for different Incident Types.

    | **Text Styling Note**     Double clicking on any section of your chart, either the image or the text, will open a Format Data Series tab on the right side of your screen. This will allow you to adjust the colours of your text, as well as other styling options such as applying borders or shadows. |

    <img src='{{ '/assets/images/M43.png' | relative_url }}' alt='Legend button is highlighted within the Format Legend context menu' title='' width='256' height='350' />

44. You did it! We now have an interactive Pivot Table and Pivot Chart that we can use to quickly analyze what buses have the most delays on any day of the week, and why.

    <img src='{{ '/assets/images/M44.png' | relative_url }}' alt='Final result is shown' title='' width='1118' height='695' />  