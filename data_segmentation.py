import pandas as pd
import numpy as np
from datetime import datetime

def load_and_clean_data():
    # Load the main Excel file and clean up column names
    print("Loading All Hands on Deck Excel file...")
    main_df = pd.read_excel("ALL_hands_on_deck.xlsx", sheet_name=None)
    main_df = pd.concat(main_df.values(), ignore_index=True)
    main_df.columns = main_df.columns.str.lower().str.strip()
    main_df = main_df.loc[:, ~main_df.columns.duplicated()]

    # Load Irene's campaign list (all sheets)
    print("Loading campaign list (all sheets)...")
    campaign_sheets = []
    excel_file = pd.ExcelFile("Irene Email campaign list.xlsx")

    # Define column mapping for standardization
    column_mapping = {
        'first name ': 'first name',
        'first name': 'first name',
        'First Name': 'first name',
        'First name': 'first name',
        'last name': 'last name',
        'Last Name': 'last name',
        'Last name': 'last name',
        'email': 'email',
        'Email': 'email',
        'EMAIL': 'email',
        'secondarymail': 'email2',
        'position': 'position',
        'Position': 'position',
        'country': 'country',
        'Country': 'country',
        'COUNTRY': 'country'
    }

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        # Rename columns using the mapping
        df.columns = df.columns.str.strip()
        df = df.rename(columns=column_mapping)

        # Ensure primary email column exists and is named correctly
        email_variants = ['email', 'Email', 'EMAIL']
        email_col = next((col for col in df.columns if col in email_variants), None)
        if email_col and email_col != 'email':
            df = df.rename(columns={email_col: 'email'})

        # Drop any secondary email columns
        if 'email2' in df.columns:
            df = df.drop(columns=['email2'])

        campaign_sheets.append(df)

    campaign_df = pd.concat(campaign_sheets, ignore_index=True)
    campaign_df = campaign_df.loc[:, ~campaign_df.columns.duplicated()]

    # Print debugging information
    print(f"\nTotal records in campaign data: {len(campaign_df)}")
    print("Campaign columns after cleaning:", sorted(campaign_df.columns.tolist()))


    # Merge main_df with campaign_df
    print("\nMerging contact data...")
    main_df = pd.merge(
        main_df,
        campaign_df,
        on='email',
        how='outer',
        suffixes=('_main', '_campaign')
    )

    # Remove duplicate columns and standardize names
    main_df = main_df.loc[:, ~main_df.columns.duplicated()]

    # Print column names for debugging
    print("\nAvailable columns after merge:")
    print(sorted(main_df.columns.tolist()))

    # Select and coalesce columns (prefer main, fallback to campaign)
    main_df['first name'] = main_df['first name_main'].fillna(main_df['first name_campaign'])
    main_df['last name'] = main_df['last name_main'].fillna(main_df['last name_campaign'])
    main_df['position'] = main_df['position_main'].fillna(main_df['position_campaign'])
    main_df['country'] = main_df['country_main'].fillna(main_df['country_campaign'])

    # Select final columns
    final_columns = [
        'first name', 'last name', 'position', 'city_main',
        'state_main', 'country', 'email', 'phone1_main', 'linked in._main'
    ]

    # Rename remaining _main columns
    main_df = main_df[final_columns].rename(columns={
        'city_main': 'city',
        'state_main': 'state',
        'phone1_main': 'phone1',
        'linked in._main': 'linked in.'
    })

    # Load and clean engagement data
    print("Loading engagement CSV file...")
    engagement_df = pd.read_csv("Recipient_Engagement_from_20250120_20250218 (1).csv")
    engagement_df.columns = engagement_df.columns.str.lower().str.strip()

    return main_df, engagement_df

def clean_and_validate_emails(main_df, engagement_df):
    print("\nValidating and cleaning email data...")

    # Store original totals for validation
    original_metrics = {
        'total_opens': engagement_df['# opens'].sum(),
        'total_clicks': engagement_df['# link clicks'].sum()
    }

    # Check for multiple emails in single fields
    multi_emails = engagement_df[engagement_df['recipient email'].str.contains(',', na=False)]
    if not multi_emails.empty:
        print(f"Found {len(multi_emails)} rows with multiple emails")
        new_rows = []

        # Split and expand multiple email entries
        for idx, row in multi_emails.iterrows():
            emails = [email.strip().lower() for email in row['recipient email'].split(',')]
            for email in emails:
                new_row = row.copy()
                new_row['recipient email'] = email
                new_rows.append(new_row)

        # Remove original rows with multiple emails
        engagement_df = engagement_df[~engagement_df['recipient email'].str.contains(',', na=False)]

        # Add new split rows
        engagement_df = pd.concat([engagement_df, pd.DataFrame(new_rows)], ignore_index=True)

    # Clean email fields
    main_df['email'] = main_df['email'].str.lower().str.strip()
    engagement_df['recipient email'] = engagement_df['recipient email'].str.lower().str.strip()

    # Aggregate metrics for duplicate emails
    engagement_df = engagement_df.groupby('recipient email').agg({
        '# of engaged emails': 'sum',
        '# opens': 'sum',
        '# link clicks': 'sum',
        '# attachment views': 'sum',
        'most recent event': 'max'  # Take the most recent date
    }).reset_index()

    # Validate totals
    new_metrics = {
        'total_opens': engagement_df['# opens'].sum(),
        'total_clicks': engagement_df['# link clicks'].sum()
    }

    print("\nMetrics Validation:")
    print(f"Original opens: {original_metrics['total_opens']:.0f}")
    print(f"New opens: {new_metrics['total_opens']:.0f}")
    print(f"Original clicks: {original_metrics['total_clicks']:.0f}")
    print(f"New clicks: {new_metrics['total_clicks']:.0f}")
    print(f"Unique emails after processing: {len(engagement_df)}")

    return main_df, engagement_df

def merge_and_process_data(main_df, engagement_df):
    # Merge datasets
    merged_df = pd.merge(
        main_df,
        engagement_df,
        left_on="email",
        right_on="recipient email",
        how='outer',
        indicator=True
    )

    print("\nDetailed Merge Statistics:")
    print(merged_df['_merge'].value_counts())

    # Handle partial matches
    partial_matches = merged_df[merged_df['_merge'] == 'right_only'].copy()
    if not partial_matches.empty:
        print(f"\nFound {len(partial_matches)} emails with engagement but no contact info")
        partial_matches = partial_matches.fillna({
            'first name': 'Unknown',
            'last name': 'Unknown',
            'position': 'Unknown',
            'city': 'Unknown',
            'state': 'Unknown',
            'country': 'Unknown',
            'phone1': 'Unknown',
            'linked in.': 'Unknown'
        })
        partial_matches['email'] = partial_matches['recipient email']

    # Combine all data
    complete_df = pd.concat([
        merged_df[merged_df['_merge'] == 'both'],
        merged_df[merged_df['_merge'] == 'left_only'],
        partial_matches
    ], ignore_index=True)

    # Clean up final dataset
    complete_df = complete_df.fillna({
        'first name': 'Unknown',
        'last name': 'Unknown',
        'position': 'Unknown',
        'city': 'Unknown',
        'state': 'Unknown',
        'country': 'Unknown',
        'phone1': 'Unknown',
        'linked in.': 'Unknown',
        '# of engaged emails': 0,
        '# opens': 0,
        '# link clicks': 0,
        '# attachment views': 0
    })
    complete_df = complete_df.drop(columns=['_merge', 'recipient email'])
    complete_df = complete_df.drop_duplicates(subset=['email'])

    return complete_df, partial_matches

def segment_engagement(row):
    if row["# opens"] > 5 or row["# link clicks"] > 2:
        return "Highly Engaged"
    elif row["# opens"] >= 1 or row["# link clicks"] >= 1:
        return "Moderately Engaged"
    else:
        return "Not Engaged"

def calculate_metrics(complete_df):
    # Apply segmentation
    complete_df["engagement_segment"] = complete_df.apply(segment_engagement, axis=1)

    # Calculate segment analysis
    segment_analysis = complete_df.groupby("engagement_segment").agg({
        "email": "count",
        "# opens": ["mean", "sum"],
        "# link clicks": ["mean", "sum"]
    }).round(2)

    # Calculate dashboard metrics
    dashboard_metrics = {
        "Total Recipients": len(complete_df),
        "Engagement Summary": {
            "Highly Engaged": len(complete_df[complete_df['engagement_segment'] == "Highly Engaged"]),
            "Moderately Engaged": len(complete_df[complete_df['engagement_segment'] == "Moderately Engaged"]),
            "Not Engaged": len(complete_df[complete_df['engagement_segment'] == "Not Engaged"])
        },
        "Key Metrics": {
            "Average Opens": complete_df['# opens'].mean(),
            "Total Opens": complete_df['# opens'].sum(),
            "Average Link Clicks": complete_df['# link clicks'].mean(),
            "Total Link Clicks": complete_df['# link clicks'].sum(),
            "Engagement Rate": (len(complete_df[complete_df['engagement_segment'].isin(["Highly Engaged", "Moderately Engaged"])])
                              / len(complete_df) * 100)
        },
        "Top Locations": complete_df.groupby('country')['email'].count().nlargest(5).to_dict()
    }

    return segment_analysis, dashboard_metrics

def save_results(complete_df, dashboard_metrics, partial_matches):
    # Define output columns
    output_columns = [
        'first name', 'last name', 'position', 'city', 'state', 'country',
        'email', 'phone1', 'linked in.', '# of engaged emails', '# opens', '# link clicks',
        '# attachment views', 'most recent event', 'engagement_segment'
    ]

    # Sort by engagement
    engagement_order = ["Highly Engaged", "Moderately Engaged", "Not Engaged"]
    complete_df['engagement_segment'] = pd.Categorical(
        complete_df['engagement_segment'],
        categories=engagement_order,
        ordered=True
    )
    complete_df = complete_df.sort_values(
        by=['engagement_segment', '# opens', '# link clicks'],
        ascending=[True, False, False]
    )

    # Save files
    complete_df[output_columns].to_csv("Segmented_Recipients_Complete6.csv", index=False)
    pd.DataFrame([dashboard_metrics]).to_csv("Dashboard_Metrics6.csv", index=False)
    if not partial_matches.empty:
        partial_matches[['recipient email', '# opens', '# link clicks']].to_csv(
            "Partial_Matches6.csv", index=False
        )

def main():
    # Load and clean data
    main_df, engagement_df = load_and_clean_data()

    # Clean and validate emails
    main_df, engagement_df = clean_and_validate_emails(main_df, engagement_df)

    # Merge and process data
    complete_df, partial_matches = merge_and_process_data(main_df, engagement_df)

    # Calculate metrics
    segment_analysis, dashboard_metrics = calculate_metrics(complete_df)

    # Print analysis results
    print("\nSegment Analysis:")
    print(segment_analysis)
    print("\nDashboard Metrics:")
    print(pd.DataFrame([dashboard_metrics]))

    # Save results
    save_results(complete_df, dashboard_metrics, partial_matches)

    print("\nFiles saved:")
    print(f"Total records processed: {len(complete_df)}")
    print("1. Segmented_Recipients_Complete.csv - Complete recipient list with engagement data")
    print("2. Dashboard_Metrics.csv - Summary metrics for Airtable dashboard")
    print("3. Partial_Matches.csv - Emails with engagement but missing contact info")

if __name__ == "__main__":
    main()