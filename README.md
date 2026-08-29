# global-health-analytics-pipeline

This project uses a large global healthcare dataset with 85,171 rows of data. 

Instead of just running simple searches, I built a complete system inside SQL Server (T-SQL) to clean and organize the data. The backend uses advanced queries, time-saving window formulas, and lag tools to find trends. To keep things running fast, I created automated stored procedures that run the heavy math, clear out old records, and save the final answers into either temp tables or permanent tables. I then built  database views so that Excel dashboards can read the final answers instantly without slowing down.

### Technical Stack
* Database Layer: SQL Server (T-SQL)
* Automation Layer: Stored Procedures, Temporary Tables, and Permanent Storage Tables
* Delivery Layer: Permanent Database Views
* Reporting Layer: Excel Pivot Tables, Pivot Charts, and Interactive Slicers on Excel


## Part 1: Population Health & Infection Acceleration Pipeline

* **The Goal for Part 1a:** To isolate a clean, streamlined dataset of core healthcare fields (cases, deaths, population) across specific countries and dates, removing junk tracking rows so external reporting tools can load the data quickly.
* **The Goal for Part 1b:** To uncover hidden survival patterns globally. Instead of just looking at raw deaths, this stage creates an original calculation to track actual survivor volumes, groups countries into performance tiers, and ranks them internally to find out which regions handled the crisis best.
* **The Goal for Part 1c:** To track how fast the virus accelerated day by day. It uses a look back tool to find the exact difference between today's cases and yesterday's cases, tags them with warning labels, and saves the answers into a permanent table then accessed via a view. 

### Part 1a: Core Column Selection View - see file named "View 1a"
* **What it does:** Filters out unused data tracking fields and narrows down the table to the exact columns needed for reporting. It sorts the results by country and date.

```sql
USE Project2COVID
GO

DROP VIEW IF EXISTS V_DataSelectionForCovidDeathsTables
GO

CREATE VIEW V_DataSelectionForCovidDeathsTables
AS
Select iso_code,  continent, location, date, total_cases, new_cases, total_deaths, new_deaths, reproduction_rate, population
from [Project2COVID].dbo.CovidDeaths


GO

Select *
from V_DataSelectionForCovidDeathsTables
Order by 3,4
```

### Part 1b: Country Survival Tier Analysis (Chained CTEs, Window Functions (ranking functions and aggregates), Case statements,  Views) - See File named "View 1b"
* **What it does:** Calculates the highest total cases for each country, subtracts deaths to isolate actual survivors, and groups countries into performance tiers ("Excellent", "Good", or "Requires Investigation") before ranking them.

```sql
USE Project2COVID
GO


DROP VIEW IF EXISTS V_CountryCaseVsDeathTier
GO

CREATE VIEW V_CountryCaseVsDeathTier
AS
WITH FindTotalSums AS (
Select Distinct  CD.location, 
Max(CD.Total_cases) over (partition by location) as MaxCases, 
Max(CD.Total_Deaths) over (partition by  location) as MaxDeaths
from [Project2COVID].dbo.CovidDeaths as CD) 
,
NumberOfCasesNoDeath AS (
Select FTS.Location, FTS.MaxCases, FTS.MaxDeaths, (FTS.MaxCases - FTS.MaxDeaths) as NoDeathCases
from FindTotalSums as FTS)
,
CountryDeathTier AS (
Select NOCND.Location, NOCND.MaxCases, NOCND.MaxDeaths, NOCND.NoDeathCases, Case 
When (NOCND.NoDeathCases) > 0.9 * NOCND.MaxCases then 'Excellent'
When (NOCND.NoDeathCases) > 0.7 * NOCND.MaxCases then 'Good'
Else 'Requires Investigation'
End As CountryCovidDeathTiers
from NumberOfCasesNoDeath as NOCND)
,
RankCountryCovidDeathTiers AS (
Select row_number() over (partition by CDT.CountryCovidDeathTiers order by CDT.NoDeathCases) as RowNumberTier, Dense_rank() over (partition by CDT.CountryCovidDeathTiers order by CDT.NoDeathCases) as RankNumberTier, *
from CountryDeathTier as CDT) 

Select *
From RankCountryCovidDeathTiers 
GO

select*
from V_CountryCaseVsDeathTier
```

### Part 1c: Daily Infection Acceleration Tracker (Stored Procedure & Temp Tables, Views, Lag function, CASE statements, Convert data types) - see file named  "view 1c"
* **What it does:** Uses a stored procedure to calculate the daily case changes using a look-back function, wipes old data to save a fresh copy inside a permanent staging table, and builds a permanent shortcut view for Excel to read.

```sql
--- Part 1c  Finding the Difference in Today's New Cases and Yesterday's New Cases (daily infection spikes)


USE Project2COVID
GO

DROP PROCEDURE IF EXISTS  DailyInfectionSpikes
GO

CREATE PROCEDURE DailyInfectionSpikes
AS
BEGIN
WITH ConvertDataTypes AS(
Select Convert(varchar(255), Location) as Location,    Convert(Date, Date) as Date, Convert(int, New_Cases) as TodayCases, Lag(Convert(int,New_Cases),1) over (partition by location order by date) as YeserdayCases
from [Project2COVID].dbo.CovidDeaths)
,
DailyChangeInCases AS (
Select ConvertDataTypes.Location, ConvertDataTypes.Date, ConvertDataTypes.TodayCases, ConvertDataTypes.YeserdayCases, Case 
when (ConvertDataTypes.TodayCases - ConvertDataTypes.YeserdayCases) >50 then 'Investigate'
when (ConvertDataTypes.TodayCases - ConvertDataTypes.YeserdayCases) > 10 then 'Monitor'
when (ConvertDataTypes.TodayCases - ConvertDataTypes.YeserdayCases) > 4 then 'Standard'
when (ConvertDataTypes.TodayCases - ConvertDataTypes.YeserdayCases) >1 then 'Good Measures in Place'
Else  'Excellent Measures in Place'
end as DailyCasesSpikeTiers
from ConvertDataTypes)

Select *
from DailyChangeInCases
END

go

USE Project2COVID
GO

DROP TABLE IF EXISTS #DAILYINFECTIONSPIKES
go


CREATE TABLE #DAILYINFECTIONSPIKES
(Location varchar(255), Date date, TodayCases int, YesterdayCases int, DailyCasesSpikeTiers varchar(255))
Go

INSERT INTO #DAILYINFECTIONSPIKES
EXEC DailyInfectionSpikes


Select *
from #DAILYINFECTIONSPIKES

GO

USE Project2COVID
GO

CREATE TABLE P_DailyInfectionSpikes
(Location varchar(255), Date date, TodayCases int, YesterdayCases int, DailyCasesSpikeTiers varchar(255))
Go


TRUNCATE TABLE  P_DailyinfectionSpikes 
Go


INSERT INTO P_DailyInfectionSpikes 
SELECT *
FROM #DAILYINFECTIONSPIKES


select *
from P_DailyInfectionSpikes
GO

USE Project2COVID
GO

DROP VIEW IF EXISTS v_DailyInfectionSpikes
GO

CREATE VIEW v_DailyInfectionSpikes
AS
select *
from [Project2COVID].dbo.P_DailyInfectionSpikes


GO

select *
from v_DailyInfectionSpikes
Order by 1
```
## Part 2: Development Index Tracking & Testing Volume Surge Models
Before running the data pipelines, the analytical goals for this section were broken down into clear targets:
* **The Goal for Part 2a:** To isolate a clean, streamlined dataset of testing, demographic, and development metrics (like HDI and life expectancy), filtering out unused fields so our downstream reporting tools can process the tables quickly.
* **The Goal for Part 2b:** To figure out how a country's wealth and development tier relates to its positive test results. This query groups countries into three dynamic brackets based on their development index, aggregates millions of tests safely, and creates a custom percentage ratio to show the relationship between health data and country wealth.
* **The Goal for Part 2c:** To build an automated tracking tool to flag sudden spikes in the daily testing workload. This pipeline looks back exactly 7 days to see what the testing volume was on the same day last week, calculates the exact difference, and automatically saves high-priority spikes into a permanent database repository table.
This phase focuses on tracking workloads, population health, and wealth metrics to analyze how a country's overall development level impacted its healthcare response.

---

### Part 2a: Core Vaccination & Development Column Selection View  - See File named "View 2a"
* **What it does:** Filters out unused data tracking fields and isolates the core columns needed to analyze testing volume alongside country development and health metrics (like human development index, life expectancy, and median age).

```sql
USE Project2COVID
GO

DROP VIEW IF EXISTS V_DataSelectionForCovidVaccinationsTables
GO

CREATE VIEW V_DataSelectionForCovidVaccinationsTables
AS
select location, date, new_tests, total_tests, positive_rate, population, median_age, life_expectancy, human_development_index, diabetes_prevalence
from [Project2COVID].dbo.CovidVaccinations.   
Order by 1
```
### Part 2b: Testing Efficiency vs. Country Development (Chained CTEs)
* **What it does:** Groups countries into development tiers using their Human Development Index (HDI) scores, totals up millions of tests using safe memory limits, and calculates a custom percentage metric showing the relationship between positive test rates and country development.

```sql
USE Project2COVID
GO

DROP VIEW IF EXISTS V_DevelopmentIndexTrackingAgainstTesting
GO

CREATE VIEW V_DevelopmentIndexTrackingAgainstTesting
AS 
WITH CountriesVsDevelopmentIndex AS (
Select Distinct Location, SUM(CONVERT(BIGint, new_tests)) over (partition by location )   as Total_Tests_Country, Max(CONVERT(float, human_development_index)) over (partition by location ) as Max_Index_Country, Case
WHEN  Max(human_development_index) over (partition by location )> 0.8 then 'high development'
WHEN Max(human_development_index) over (partition by location )>0.6 then 'medium development'
ELSE 'low development'
END  as DevelopmentTiers, positive_rate
From  [Project2COVID].dbo.CovidVaccinations )
,
AverageTestRate AS (
Select Distinct CDI.location, CDI.Total_Tests_Country, CDI.Max_Index_Country, CDI.DevelopmentTiers, Avg(CONVERT(float, CDI.positive_rate)) over (partition by location ) as Average_Positive_Rate_Country
FROM CountriesVsDevelopmentIndex as CDI)
,
PositiveTestDevelopmentRatio AS (
SELECT DISTINCT ATR.LOCATION, ATR.Total_Tests_Country, ATR.Max_Index_Country, ATR.DevelopmentTiers,   ATR.Average_Positive_Rate_Country, ((ATR.Average_Positive_Rate_Country/ATR.Max_Index_Country)*100) AS percentage_Positive_rate_Over_Dev_Index
FROM AverageTestRate as ATR)

SELECT *
FROM PositiveTestDevelopmentRatio

GO

SELECT *
FROM V_DevelopmentIndexTrackingAgainstTesting
WHERE Max_Index_Country IS NOT NULL
AND
Total_Tests_Country IS NOT NULL
ORDER BY 1
GO
```
### Part 2c: Weekly Testing Workload Alerts (etl pipeline)
* **What it does:** Automates a 7-day look-back calculation to track operational surges, dumps the output straight into a permanent data warehouse table using a truncate-and-load model, and establishes a secure virtual view layer for instant dashboard access.

```sql
USE Project2COVID
GO

DROP PROCEDURE IF EXISTS WEEK_DAY_TESTING_PROCEDURE
GO

CREATE PROCEDURE WEEK_DAY_TESTING_PROCEDURE
AS
BEGIN
WITH WEEK_DAY_LAG_TABLE AS (
SELECT CONVERT(VARCHAR(255), LOCATION) as location ,  CONVERT(DATE, DATE) as date, CONVERT(INT, new_tests) as new_tests, LAG(CONVERT(INT, new_tests), 7) over (partition by LOCATION ORDER BY Date) AS WEEK_AGO
FROM [Project2COVID].DBO.CovidVaccinations as CV) 
,
WEEK_Day_Tracking AS (
SELECT WDLT.LOCATION, WDLT.DATE, WDLT.new_tests, WDLT.WEEK_AGO, (WDLT.WEEK_AGO-WDLT.new_tests) AS IncreaseInNewTests, CASE 
WHEN (WDLT.WEEK_AGO-WDLT.new_tests) > 6000 THEN 'HIGH PRIORITY'
WHEN (WDLT.WEEK_AGO-WDLT.new_tests)  BETWEEN -2000 AND 2000 THEN 'MEDIUM PRIORITY'
ELSE 'LOW PRIORITY'
END AS PRIORITYTRACKING 
FROM WEEK_DAY_LAG_TABLE AS WDLT)

SELECT *
FROM WEEK_Day_Tracking 
END
GO

USE Project2COVID
GO

DROP TABLE IF EXISTS P_Week_Test_Tracker
GO


CREATE TABLE P_Week_Test_Tracker
(LOCATION varchar(255), DATE date, new_tests int, WEEK_AGO int, IncreaseInNewTests int, PRIORITYTRACKING varchar(255))
GO

INSERT INTO P_Week_Test_Tracker 
EXEC WEEK_DAY_TESTING_PROCEDURE
go

USE Project2COVID
GO

DROP VIEW IF EXISTS V_Week_Test_Tracker 
GO

CREATE VIEW V_Week_Test_Tracker 
AS
SELECT*
FROM [Project2COVID].DBO.P_Week_Test_Tracker 
GO

SELECT *
FROM V_Week_Test_Tracker
```
## Part 3: Deep Data Analysis (Nested Subqueries)

This final section brings all of our independent tracking numbers together into a single global reporting tool.

### Part 3a: Multi-Layer Nested Subquery Storage Pipeline
* **What it does:** Runs an advanced 4-layer deep nested subquery procedure to link our country metrics together. It calculates testing numbers against a 7-day look-back infection trend, assigns active efficiency labels, and saves the data straight down into a clean permanent table for reporting.

```sql
--- stored procedure
USE Project2COVID
GO

DROP PROCEDURE IF EXISTS MULTIJOINWORKINGCODE
GO

CREATE PROCEDURE MULTIJOINWORKINGCODE
AS
BEGIN
SELECT final.Location, final.DATE, final.DevelopmentTiers, final.Max_Index_Country, final.CountryCovidDeathTiers, final.TodayCases, final.Week_Ago_Case,  final.IncreaseInNewTests,  (final.TodayCases-final.Week_Ago_Case)  as IncreaseInNewCases, CASE WHEN (final.IncreaseInNewTests - (final.TodayCases-final.Week_Ago_Case) ) >500 THEN 'Testing is Moderate' WHEN (final.IncreaseInNewTests - (final.TodayCases-final.Week_Ago_Case)) > 200 THEN 'Effective' ELSE 'Not Effective' END AS TESTINGEFFECTIVNESS
FROM 
(SELECT DIS.Date, SECONDJOIN.Location, SECONDJOIN.DevelopmentTiers, SECONDJOIN.Max_Index_Country, SECONDJOIN.CountryCovidDeathTiers, SECONDJOIN.IncreaseInNewTests, DIS.TodayCases, lag(DIS.TodayCases, 7)  over (partition by SECONDJOIN.location Order by date) as Week_Ago_Case
FROM
(SELECT FIRSTJOINDEATHSANDINDEX.Location, FIRSTJOINDEATHSANDINDEX.DevelopmentTiers, FIRSTJOINDEATHSANDINDEX.Max_Index_Country, FIRSTJOINDEATHSANDINDEX.MaxDeaths, FIRSTJOINDEATHSANDINDEX.CountryCovidDeathTiers, VP3CTEST.IncreaseInNewTests 
FROM 
(SELECT  DITTABLE.Location,  DITTABLE.DevelopmentTiers, DITTABLE.Max_Index_Country, CCVDT.MaxDeaths, CCVDT.CountryCovidDeathTiers
FROM
(SELECT  DIT.Location, DIT.DevelopmentTiers, DIT.Max_Index_Country
FROM (
SELECT Location, DevelopmentTiers, Max_Index_Country
FROM  dbo.V_DevelopmentIndexTrackingAgainstTesting) as DIT) as  DITTABLE
INNER JOIN [Project2COVID].dbo.V_CountryCaseVsDeathTier as CCVDT
on DITTABLE.Location = CCVDT.Location) AS FIRSTJOINDEATHSANDINDEX
inner join [Project2COVID].dbo.V_P3C_Week_Test_Tracker  as VP3CTEST
on FIRSTJOINDEATHSANDINDEX.Location = VP3CTEST.Location ) as SECONDJOIN
INNER JOIN  [Project2COVID].dbo.P_dailyinfectionspikes as DIS
on SECONDJOIN.location = DIS.Location ) AS final
END

EXEC MULTIJOINWORKINGCODE
GO

--creating permenant table

USE Project2COVID
GO

DROP TABLE IF EXISTS P_GlobalHealthReporting
GO

CREATE TABLE P_GlobalHealthReporting
(Location varchar(255), Date Date, DevelopmentTiers varchar(255), Max_Index_Country float, CountryCovidDeathTiers varchar(255), Today_Cases int, Week_Ago_Case int,  IncreaseInNewTests int,  IncreaseInNewCases int, TESTINGEFFECTIVNESS varchar(255))
GO

INSERT INTO P_GlobalHealthReporting
EXEC MULTIJOINWORKINGCODE
GO

-- Creating a View
USE Project2COVID
GO

DROP VIEW IF EXISTS V_GlobalHealthReporting
GO

CREATE VIEW V_GlobalHealthReporting
AS
SELECT *
FROM P_GlobalHealthReporting
go

SELECT *
FROM V_GlobalHealthReporting
```
