-- Run ONLY on an isolated, disposable PowerBIAnalytics_DEV lab database.
-- Existing table/data are left unchanged.
USE [PowerBIAnalytics_DEV];
GO
IF OBJECT_ID(N'dbo.SSIS_Benchmark_Results',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SSIS_Benchmark_Results
    (
      BenchmarkID BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
      OrderID INT NULL,
      ProductID INT NULL,
      Quantity INT NULL,
      SalesAmount FLOAT NULL,
      LoadTimestamp DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
    );
END;
GO
SELECT OBJECT_ID(N'dbo.SSIS_Benchmark_Results',N'U') AS BenchmarkTableObjectID;
