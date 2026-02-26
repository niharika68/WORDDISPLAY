"""
Data Source Module
==================
Responsible for generating synthetic healthcare procurement data.
This module contains NO report generation logic - only data creation.
"""

import random
from datetime import datetime, timedelta
from typing import Tuple

import pandas as pd


# Constants for data generation
HOSPITALS = [
    "City General Hospital",
    "St. Mary's Medical Center",
    "University Health System",
    "Regional Medical Center",
    "Community Health Hospital",
    "Metropolitan Care Center",
    "Valley View Hospital",
    "Riverside Medical Center",
]

PHARMACIES = [
    "MedSupply Plus",
    "PharmaCare Direct",
    "HealthRx Solutions",
    "National Drug Distributors",
    "Premier Pharmacy Services",
    "MedLine Wholesale",
]

DRUGS = [
    ("Amoxicillin 500mg", 12.50, 150),
    ("Lisinopril 10mg", 8.75, 200),
    ("Metformin 850mg", 6.25, 300),
    ("Omeprazole 20mg", 15.00, 180),
    ("Atorvastatin 40mg", 22.50, 120),
    ("Amlodipine 5mg", 9.00, 250),
    ("Metoprolol 50mg", 11.25, 175),
    ("Losartan 100mg", 18.75, 140),
    ("Gabapentin 300mg", 14.00, 160),
    ("Sertraline 50mg", 10.50, 190),
    ("Hydrochlorothiazide 25mg", 5.50, 280),
    ("Pantoprazole 40mg", 16.25, 130),
]

SAVINGS_INDICATORS = [
    "Saved 8% vs last month",
    "Saved 12% vs last month",
    "Saved 5% vs last month",
    "Costs increased 3%",
    "Saved 15% vs last month",
    "Costs stable",
    "Saved 10% vs last month",
    "Saved 6% vs last month",
]


def _generate_random_date(start_date: datetime, end_date: datetime) -> datetime:
    """Generate a random date between start_date and end_date."""
    delta = end_date - start_date
    random_days = random.randint(0, delta.days)
    return start_date + timedelta(days=random_days)


def _generate_orders_data(num_records: int = 100) -> pd.DataFrame:
    """
    Generate synthetic orders dataset.
    
    Args:
        num_records: Number of order records to generate.
        
    Returns:
        DataFrame containing orders with calculated Order Value.
    """
    random.seed(42)  # For reproducibility
    
    # Date range: last 6 months
    end_date = datetime.now()
    start_date = end_date - timedelta(days=180)
    
    orders = []
    for _ in range(num_records):
        # Select random drug with base price and typical units
        drug_name, base_price, typical_units = random.choice(DRUGS)
        
        # Add some variation to price and units
        price = round(base_price * random.uniform(0.9, 1.1), 2)
        units = int(typical_units * random.uniform(0.5, 1.5))
        
        order = {
            "Hospital": random.choice(HOSPITALS),
            "Pharmacy": random.choice(PHARMACIES),
            "Drug": drug_name,
            "Price": price,
            "Units": units,
            "Date Ordered": _generate_random_date(start_date, end_date),
            "Invoiced": random.choice(["Yes", "Yes", "Yes", "No"]),  # 75% invoiced
        }
        orders.append(order)
    
    # Create DataFrame
    df = pd.DataFrame(orders)
    
    # Add calculated column
    df["Order Value"] = (df["Price"] * df["Units"]).round(2)
    
    # Sort by date
    df = df.sort_values("Date Ordered", ascending=False).reset_index(drop=True)
    
    return df


def _generate_summary_data(orders_df: pd.DataFrame) -> pd.DataFrame:
    """
    Generate summary dataset from orders data.
    
    Args:
        orders_df: The orders DataFrame to summarize.
        
    Returns:
        DataFrame containing monthly summary statistics.
    """
    # Extract month from date
    orders_df = orders_df.copy()
    orders_df["Month"] = orders_df["Date Ordered"].dt.to_period("M")
    
    # Group by month
    summary = orders_df.groupby("Month").agg(
        Total_Orders=("Drug", "count"),
        Total_Spend=("Order Value", "sum")
    ).reset_index()
    
    # Convert period to string for cleaner display
    summary["Month"] = summary["Month"].astype(str)
    
    # Rename columns
    summary.columns = ["Month", "Total Orders", "Total Spend"]
    
    # Round total spend
    summary["Total Spend"] = summary["Total Spend"].round(2)
    
    # Add savings indicator (synthetic)
    random.seed(123)
    summary["Savings Indicator"] = [
        random.choice(SAVINGS_INDICATORS) for _ in range(len(summary))
    ]
    
    # Sort by month descending
    summary = summary.sort_values("Month", ascending=False).reset_index(drop=True)
    
    return summary


def get_report_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Main entry point for data generation.
    
    Returns:
        Tuple of (summary_dataframe, orders_dataframe)
    """
    # Generate orders data
    orders_df = _generate_orders_data(num_records=100)
    
    # Generate summary from orders
    summary_df = _generate_summary_data(orders_df)
    
    # Format date for display in orders
    orders_df["Date Ordered"] = orders_df["Date Ordered"].dt.strftime("%Y-%m-%d")
    
    return summary_df, orders_df


# For testing this module independently
if __name__ == "__main__":
    summary, orders = get_report_data()
    
    print("=" * 60)
    print("SUMMARY DATA")
    print("=" * 60)
    print(summary.to_string(index=False))
    
    print("\n" + "=" * 60)
    print("ORDERS DATA (First 10 rows)")
    print("=" * 60)
    print(orders.head(10).to_string(index=False))
    
    print(f"\nTotal orders: {len(orders)}")
    print(f"Total months: {len(summary)}")
