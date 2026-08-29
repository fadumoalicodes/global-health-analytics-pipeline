# global-health-analytics-pipeline

This project takes a huge list of global healthcare tracking data, analyses it inside SQL Server (T-SQL), and pulls results into Excel to build an interactive dashboard.

### 💻 Tools Used
* **Database:** SQL Server (T-SQL)
* **Dashboard:** Excel Pivot Tables, Pivot Charts, and Interactive Slicers


## Part 1: Advanced Data Exploration with SQL

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
## Part 2: Vaccination, Testing & Country Development Exploration with SQL
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
