# Data Pipeline Architecture

## Overview

This project uses a structured Excel Business Intelligence pipeline to transform raw trading data into interactive dashboard visualizations.

The architecture follows a layered analytics workflow:

```text
Raw Trade Data
    ↓
Power Query Cleaning
    ↓
Data Model / Power Pivot
    ↓
DAX Measures
    ↓
Pivot Table Controller
    ↓
Cube Formulas
    ↓
Dashboard Visualizations
```

Each stage performs a specific role in the reporting and analytics process.

# 2. Power Query Cleaning

Power Query is used as the ETL (Extract, Transform, Load) layer of the project.

Key transformations include:

- Data type conversion
- Removal of null values
- Standardization of date formats
- Trade duration calculations
- Session classification
- Conditional columns
- Profit/Loss categorization
- Time-based transformations

Example Power Query operations:

- Custom Columns
- Conditional Columns
- Merge Queries
- Append Queries
- DateTime transformations
- Duration calculations

This stage ensures that all downstream analytics are based on clean, structured, and reliable data.

# 3. Data Model / Power Pivot

The cleaned data is loaded into the Excel Data Model using Power Pivot.

The data model is designed using a relational structure to improve scalability and analytical performance.

Model components include:

| Sheet1 Trable | Stores transactional trade data |
| Power Pivot Calendar Table | Supports time intelligence calculations 

Relationships are established between the tables to enable advanced filtering and aggregation.

The Power Pivot layer acts as the analytical engine of the dashboard.

# 4. DAX Measures

DAX (Data Analysis Expressions) measures are used to calculate trading and operational metrics dynamically.

Key measures include:

Cumulative P&L:=VAR StartDate = MIN('Calendar'[Date])
VAR EndDate   = MAX('Calendar'[Date])

RETURN
IF(
    ISBLANK(
        CALCULATE(
            SUM(Sheet1[Net_P&L]),
            DATESBETWEEN(
                'Calendar'[Date],
                StartDate,
                EndDate
            )
        )
    ),
    BLANK(),

    CALCULATE(
        SUM(Sheet1[Net_P&L]),

        FILTER(
            ALL('Calendar'),
            'Calendar'[Date] <= MAX('Calendar'[Date])
                &&
            'Calendar'[Date] >= StartDate
        )
    )
)

Drawdown %:=
VAR DrawdownTable =
    ADDCOLUMNS(
        VALUES('Calendar'[Date]),

        "Equity",
            CALCULATE([Cumulative P&L]),

        "RunningPeak",
            CALCULATE(
                MAXX(
                    FILTER(
                        ALLSELECTED('Calendar'[Date]),
                        'Calendar'[Date] <= MAX('Calendar'[Date])
                    ),
                    [Cumulative P&L]
                )
            )
    )

VAR MaxDD =
    MINX(
        DrawdownTable,
        DIVIDE(
            [Equity] - [RunningPeak],
            [RunningPeak],
            0
        )
    )

RETURN
MaxDD

Drawdown % Prior Month
:=CALCULATE(
    [Drawdown %],
    DATEADD('Calendar'[Date], -1, MONTH)
)

Drawdown % vs Prior
:=[Drawdown %]
-
[Drawdown % Prior Month]

Drawdown % vs Prior %
:=DIVIDE(
    [Drawdown % vs Prior],
    ABS([Drawdown % Prior Month]),
    0
)

Total P&L
:=SUM('Sheet1'[Net_P&L])

Total Trades
:=VAR StartDate = MIN('Calendar'[Date])
VAR EndDate   = MAX('Calendar'[Date])

VAR _Count =
    CALCULATE(
        DISTINCTCOUNT(Sheet1[Trade_ID]),
        DATESBETWEEN(
            'Calendar'[Date],
            StartDate,
            EndDate
        )
    )

RETURN
IF(ISBLANK(_Count), 0, _Count)

Win Rate
:=VAR StartDate = MIN('Calendar'[Date])
VAR EndDate   = MAX('Calendar'[Date])

VAR _Wins =
    CALCULATE(
        COUNTROWS(Sheet1),
        Sheet1[Trade Outcome] = "Win",
        DATESBETWEEN(
            'Calendar'[Date],
            StartDate,
            EndDate
        )
    )

VAR _TotalTrades =
    CALCULATE(
        COUNTROWS(Sheet1),
        DATESBETWEEN(
            'Calendar'[Date],
            StartDate,
            EndDate
        )
    )

RETURN
DIVIDE(_Wins, _TotalTrades, 0)

Sharpe Ratio
:=VAR AvgReturn = AVERAGE(Sheet1[Net_P&L])
VAR StdDevReturn = STDEV.P(Sheet1[Net_P&L])
RETURN
DIVIDE(AvgReturn, StdDevReturn, 0) * SQRT(252)

Profit Factor
:=VAR StartDate = MIN('Calendar'[Date])
VAR EndDate   = MAX('Calendar'[Date])

VAR GrossProfit =
    CALCULATE(
        SUM(Sheet1[Net_P&L]),
        Sheet1[Trade Outcome] = "Win",
        DATESBETWEEN(
            'Calendar'[Date],
            StartDate,
            EndDate
        )
    )

VAR GrossLoss =
    ABS(
        CALCULATE(
            SUM(Sheet1[Net_P&L]),
            Sheet1[Trade Outcome] = "Loss",
            DATESBETWEEN(
                'Calendar'[Date],
                StartDate,
                EndDate
            )
        )
    )

RETURN
DIVIDE(GrossProfit, GrossLoss, 0)

Calmer Ratio
:=VAR TotalDays = COUNTROWS('Calendar')
VAR CumulativeReturnPct = DIVIDE([Total P&L], AVERAGE('Sheet1'[Starting_Balance]), 0)
VAR AnnualizedReturn = (1 + CumulativeReturnPct) ^ (365 / TotalDays) - 1
VAR MaxDD = ABS([Drawdown %])
RETURN
DIVIDE(AnnualizedReturn, MaxDD, 0)

DAX measures allow all calculations to respond dynamically to slicers, timelines, and dashboard filters.

# 5. Pivot Table Controller

A Pivot Table is used as the central controller for slicers and timeline interactions.

The Pivot Table connects directly to the Data Model and enables:

- Centralized filtering
- Timeline synchronization
- Slicer control
- Dynamic context filtering

This architecture allows dashboard components to remain synchronized without exposing Pivot Tables directly on the dashboard interface.

# 6. Cube Formulas

Cube Functions are used to extract dynamic values directly from the Data Model.

Functions used include:
| CUBEVALUE | Returns aggregated measure values |
| CUBEMEMBER | Returns dimension members |
| CUBESET | Creates custom filtered sets |
| CUBERANKEDMEMBER | Returns the members of the cubeset |

Cube Functions provide several advantages:

- Flexible dashboard layouts
- Custom KPI cards
- Independent visual positioning
- Dynamic filtering support
- Advanced dashboard customization

This reporting layer enables the dashboard to behave more like a professional Business Intelligence application rather than a standard Excel report.

# 7. Dashboard Visualizations

The final layer presents the analytical outputs through interactive visualizations.

Dashboard components include:

- KPI Cards
- Equity Curve
- Heatmaps
- Monthly Performance Tables
- Session Analytics
- Trade Duration Analysis
- Interactive Slicers
- Timeline Filters

The dashboard is designed to provide clear operational and trading insights while maintaining a professional institutional-style layout.

# Pipeline Benefits

This architecture provides several advantages:

- Automated reporting workflows
- Scalable analytics structure
- Centralized calculation logic
- Dynamic filtering capabilities
- Reduced manual reporting effort
- Improved dashboard flexibility
- Enhanced analytical performance

The layered architecture also improves maintainability by separating:

- Data transformation
- Data modeling
- Calculations
- Reporting
- Visualization

# Future Improvements

Potential future enhancements include:

- Python integration
- SQL database integration
- VBA automation
- Real-time trade ingestion
- API connectivity
- Risk analytics expansion
- Multi-asset reporting