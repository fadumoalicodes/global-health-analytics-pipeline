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

