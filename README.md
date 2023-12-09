# Maven-Power-Outage-Challenge

This project seeks to clean and analyze the Department of Energy dataset provided by Maven Analytics. The dataset spans multiple years and contains information on electric utility incidents from the North American Electric Reliability Corporation (NERC). I utilized R to extract and clean the data, SQL Server for storage, and Power BI to provide in-depth analysis and a corresponding dashboard. This deliverable visualizes trends in electric utility events, quantifies their impact on communities, and pinpoints potential vulnerabilities in the grid.

Project Components:
1.	R Script: as a preprocessing pipeline designed to extract, clean, and standardize the data to address numerous data quality issues. The script includes readxl, tidyr, odbc, and DBI as required libraries and defines custom functions to handle unique irregularities in the headers, metadata, and datatypes that result from changing data quality standards over the time this data was collected. Anomalies and inconsistent values are identified and corrected where possible.
2.	SQL Server Database: hosted locally in SQL Server Express to store the data before loading in Power BI. While it would have been possible to access the data directly via R script in Power BI, this would repeat the extraction and cleaning process in R whenever the data was refreshed in Power BI and increase refresh times.
3.	Power BI: used to define KPIs, develop visualizations, and analyze the cleaned data.

Assumptions: 
1.	Approximations were considered final values when considering dates, times, demand loss, and customers impacted.
2.	Incidents listed as “Ongoing” were treated as unknown values due to a lack of updates.
3.	Discrepancies in restoration time were resolved by discarding conflicting values.
4.	Weather and Natural Disasters were treated as separate categories to account for differences in the severity of the events.
5.	Physical events, such as theft, attacks, vandalism, and sabotage, were grouped together as Physical Attacks due to significant overlap between each of these more granular incident definitions.
6.	Inconsistencies between NERC Regions were addressed by replacing Reliability Council values with corresponding NERC Regional Entities and by replacing defunct NERC Regional Entities with their modern successors. An Incident ID value was implemented to avoid the double-counting of customers impacted and demand loss.

Main Findings:
1.	The analysis revealed a significant uptick in incidents, coupled with a significant improvement in restoration times across all major NERC Regions.
2.	The trend persisted through 2023, witnessing a 4.41% increase of incidents over the same period in 2022. 
3.	33% of incidents resulted in demand loss and 49% had a measurable impact on customers among incidents where these values were reported.

While the resilience of the grid in the face of major disasters has improved, weather remains the largest concern for electrical utilities overall. The analysis demonstrates an improving response time that indicates advancements in addressing power outage incidents efficiently but also suggests an increased need for proactive action to prevent power disruptions and enhance the overall resilience of the electrical infrastructure. 


