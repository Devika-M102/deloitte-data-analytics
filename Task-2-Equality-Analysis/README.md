# Task 2: Equality Analysis in Excel

## Project Overview
This task involved classifying employee compensation equality scores across factories and job roles using Microsoft Excel.

## Objective
Create an `Equality class` column to categorize each equality score according to its absolute value.

## Classification Rules

| Equality Score | Equality Class |
|---|---|
| -10 to +10 | Fair |
| Below -10 or above +10, up to -20 or +20 | Unfair |
| Below -20 or above +20 | Highly Discriminative |

## Tools Used
- Microsoft Excel
- Excel IF and ABS functions

## Formula
```excel
=IF(ABS(C2)>20,"Highly Discriminative",IF(ABS(C2)>10,"Unfair","Fair"))
```

The `ABS` function evaluates positive and negative scores consistently, while nested `IF` functions assign the appropriate category.

## Key Learning
Practiced Excel-based data classification, conditional logic, and interpreting equality scores.

*Completed as part of the Deloitte Data Analytics Virtual Job Simulation on Forage.*
