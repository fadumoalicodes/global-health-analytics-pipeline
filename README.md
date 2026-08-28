# global-health-analytics-pipeline

This project takes a huge list of global healthcare tracking data, analyses it inside SQL Server (T-SQL), and pulls results into Excel to build an interactive dashboard.

### 💻 Tools Used
* **Database:** SQL Server (T-SQL)
* **Dashboard:** Excel Pivot Tables, Pivot Charts, and Interactive Slicers


## Part 1: Advanced Data Exploration with SQL

* **The Goal for Part 1a:** To isolate a clean, streamlined dataset of core healthcare fields (cases, deaths, population) across specific countries and dates, removing junk tracking rows so external reporting tools can load the data quickly.
* **The Goal for Part 1b:** To uncover hidden survival patterns globally. Instead of just looking at raw deaths, this stage creates an original calculation to track actual survivor volumes, groups countries into performance tiers, and ranks them internally to find out which regions handled the crisis best.

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

### Part 1c: Daily Infection Acceleration Tracker (Stored Procedure & Cache)
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
