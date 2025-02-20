# Email Campaign Data Segmentation Tool

## Overview
This tool processes and analyzes email campaign data from multiple sources to generate engagement metrics and segment recipients based on their interaction levels.

## Features
- Merges contact data from multiple Excel sheets
- Cleans and standardizes email addresses and contact information
- Segments recipients into engagement categories:
  - Highly Engaged (>5 opens or >2 link clicks)
  - Moderately Engaged (≥1 open or ≥1 link click)
  - Not Engaged (no interaction)
- Generates comprehensive engagement metrics and analytics
- Handles duplicate entries and multiple email formats

## Input Files Required
1. `ALL_hands_on_deck.xlsx` - Main contact database
2. `Irene Email campaign list.xlsx` - Campaign contact lists
3. `Recipient_Engagement_from_20250120_20250218 (1).csv` - Engagement metrics

## Output Files
1. `Segmented_Recipients_Complete.csv` - Complete recipient list with engagement data
2. `Dashboard_Metrics.csv` - Summary metrics for dashboard
3. `Partial_Matches.csv` - Records with engagement but missing contact info

## Usage
```python
python data_segmentation.py
```

## Dependencies
- pandas
- numpy

## Key Metrics Tracked
- Total Recipients
- Engagement Summary
- Average Opens
- Total Opens
- Average Link Clicks
- Total Link Clicks
- Engagement Rate
- Top Locations

## Data Processing Pipeline
1. Load and clean contact data
2. Validate and standardize email addresses
3. Merge contact and engagement data
4. Segment recipients based on engagement
5. Calculate metrics and generate reports

## Notes
- Handles multiple email formats and duplicates
- Preserves data from both main database and campaign lists
- Automatically handles missing data fields