"""
Recommendation Engine for Standard Bank – Next Best Action project
(Advanced Logic Edition for MSc / Mid-Level Portfolio)

This script demonstrates a Next Best Action (NBA) logic:
1. It loads the real CHURNDATA Excel file to retrieve actual customer features.
2. It applies a heuristic/rule-based clustering to map the customer into a "Churn Persona".
   (In a production system, this could be an unsupervised K-Means model or Uplift model).
3. It maps the persona to the `dummy_recommendations.csv` to fetch the optimal product and expected conversion rate.

Usage:
    python recommendation_engine.py --customer_id XXXXXX
"""

import argparse
import pathlib
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format="%(message)s")

def assign_churn_persona(customer_data: pd.Series) -> str:
    """
    Dynamically assigns a churn persona based on real customer financial features.
    This simulates a clustering or uplift model.
    """
    # Example logic based on CHURNDATA features
    transactions = customer_data.get('total transactions', 0)
    income = customer_data.get('CUS_Month_Income', 0)
    
    # Handle missing or non-numeric safely
    try:
        transactions = float(transactions)
    except:
        transactions = 0
        
    try:
        income = float(income)
    except:
        income = 0

    if transactions < 5:
        return "Inactive"
    elif transactions >= 5 and income > 15000:
        return "High Value"
    elif transactions >= 5 and transactions <= 15:
        return "Low Engagement"
    else:
        return "Recent Complaints" # Fallback

def get_next_best_action(customer_id: str, real_data_path: pathlib.Path, dummy_rec_path: pathlib.Path):
    """Fetches real customer data, assigns persona, and returns recommendation."""
    
    # 1. Load Real Data to get customer context
    try:
        df_real = pd.read_excel(real_data_path)
    except Exception as e:
        logging.error(f"Failed to load real data from {real_data_path}: {e}")
        return

    customer_row = df_real[df_real['CIF'].astype(str) == str(customer_id)]
    
    if customer_row.empty:
        logging.error(f"Customer ID '{customer_id}' not found in {real_data_path.name}")
        return
        
    customer_data = customer_row.iloc[0]
    
    # 2. Assign Persona
    persona = assign_churn_persona(customer_data)
    
    # 3. Lookup Recommendation
    try:
        df_rec = pd.read_csv(dummy_rec_path)
    except Exception as e:
        logging.error(f"Failed to load dummy recommendations from {dummy_rec_path}: {e}")
        return
        
    match = df_rec[df_rec['churn_type'].str.lower() == persona.lower()]
    
    if match.empty:
        logging.error(f"No recommendation mapped for persona '{persona}'")
        return
        
    rec = match.iloc[0]
    
    # Output the Next Best Action
    print("\n" + "="*50)
    print(" 🎯 NEXT BEST ACTION (NBA) RECOMMENDATION")
    print("="*50)
    print(f"Customer CIF         : {customer_id}")
    print(f"Age                  : {customer_data.get('AGE', 'N/A')}")
    print(f"Monthly Income       : {customer_data.get('CUS_Month_Income', 'N/A')}")
    print(f"Total Transactions   : {customer_data.get('total transactions', 'N/A')}")
    print("-" * 50)
    print(f"Derived Persona      : ** {rec['churn_type'].upper()} **")
    print(f"Recommended Product  : {rec['recommended_product']}")
    print(f"Expected Uplift (CR) : {rec['expected_conversion_rate']:.2%}")
    print("="*50 + "\n")

def main():
    parser = argparse.ArgumentParser(description="Next Best Action Recommendation Lookup")
    parser.add_argument("--customer_id", type=str, required=True, help="Customer CIF from CHURNDATA")
    parser.add_argument("--real_data", type=pathlib.Path, default=pathlib.Path("data/churn/CHURNDATA (1).xlsx"))
    parser.add_argument("--dummy_rec", type=pathlib.Path, default=pathlib.Path("data/recommendations/dummy_recommendations.csv"))
    args = parser.parse_args()

    get_next_best_action(args.customer_id, args.real_data, args.dummy_rec)

if __name__ == "__main__":
    main()
