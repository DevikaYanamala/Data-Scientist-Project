# recommendation_engine.py
"""
Recommendation Engine for Standard Bank – Next Best Action project

This script reads the dummy recommendation dataset (CSV) and provides a
simple API to retrieve the best product suggestion for a given churn type or
customer ID. It can be extended to integrate with a real‑time service.

Typical usage:
    python recommendation_engine.py --input data/recommendations/dummy_recommendations.csv \
        --customer_id 3

The script will print the recommended product and the expected conversion rate.
"""

import argparse
import pathlib
import pandas as pd

def load_recommendations(csv_path: pathlib.Path) -> pd.DataFrame:
    """Load the dummy recommendations CSV.

    The CSV must contain the columns:
    - customer_id
    - churn_type
    - recommended_product
    - expected_conversion_rate
    """
    df = pd.read_csv(csv_path)
    return df

def get_recommendation(df: pd.DataFrame, customer_id: int = None, churn_type: str = None) -> pd.Series:
    """Return a single recommendation row.

    - If ``customer_id`` is provided, we look for that exact row.
    - If ``customer_id`` is not provided but ``churn_type`` is, we return the
      first row matching the churn type.
    - If neither is provided, the function raises a ValueError.
    """
    if customer_id is not None:
        match = df[df["customer_id"] == customer_id]
        if not match.empty:
            return match.iloc[0]
        raise ValueError(f"No recommendation found for customer_id={customer_id}")
    if churn_type is not None:
        match = df[df["churn_type"].str.lower() == churn_type.lower()]
        if not match.empty:
            return match.iloc[0]
        raise ValueError(f"No recommendation found for churn_type='{churn_type}'")
    raise ValueError("Either customer_id or churn_type must be supplied")

def main() -> None:
    parser = argparse.ArgumentParser(description="Simple recommendation lookup")
    parser.add_argument("--input", type=pathlib.Path, required=True,
                        help="Path to dummy_recommendations.csv")
    parser.add_argument("--customer_id", type=int, default=None,
                        help="Customer ID to look up")
    parser.add_argument("--churn_type", type=str, default=None,
                        help="Churn type to look up (e.g., 'High Value')")
    args = parser.parse_args()

    df = load_recommendations(args.input)
    try:
        rec = get_recommendation(df, args.customer_id, args.churn_type)
        print("--- Recommendation ---")
        print(f"Customer ID          : {rec['customer_id']}")
        print(f"Churn Type           : {rec['churn_type']}")
        print(f"Recommended Product  : {rec['recommended_product']}")
        print(f"Expected Conv. Rate  : {rec['expected_conversion_rate']:.2%}")
    except ValueError as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
