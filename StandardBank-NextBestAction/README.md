# Standard Bank - Next Best Action Project

## Overview
This repository contains an end-to-end Machine Learning pipeline to predict customer churn and a **Next Best Action (NBA) Recommendation Engine** to suggest products that maximize retention conversion rates.

This project goes beyond basic prediction by combining a high-performing classification engine with actionable, data-driven recommendations, demonstrating a mature approach to commercial Data Science.

## Technical Highlights (MSc / Mid-Level Standard)
1. **Classification Engine (`classification_engine.py`)**
   - **Data Preprocessing**: Scikit-Learn `ColumnTransformer` (StandardScaler, OneHotEncoder).
   - **Imbalance Handling**: `SMOTE` (Synthetic Minority Over-sampling Technique) embedded inside an `imblearn` pipeline to prevent data leakage during cross-validation.
   - **Algorithm**: `XGBoost` Classifier, highly optimized for tabular financial data.
   - **Hyperparameter Tuning**: `GridSearchCV` for optimal learning rate, depth, and tree count.
   - **Explainability**: `SHAP` (SHapley Additive exPlanations) values to interpret model decisions—crucial for regulatory compliance in banking.

2. **Recommendation Engine (`recommendation_engine.py`)**
   - Implements dynamic **Churn Persona Assignment** based on real financial features (e.g., transaction volume, income).
   - Maps personas to a curated recommendations matrix (`dummy_recommendations.csv`) to surface the Next Best Action and expected uplift (Conversion Rate).

## Repository Structure
- `scripts/classification_engine.py` – Advanced ML pipeline with SMOTE & SHAP.
- `scripts/recommendation_engine.py` – Dynamic NBA lookup tool.
- `scripts/generate_presentation.py` – Auto-generates a polished PowerPoint deck (`NextBestAction.pptx`) using `python-pptx`.
- `data/churn/CHURNDATA (1).xlsx` – Real banking dataset.
- `data/recommendations/dummy_recommendations.csv` – NBA mapping matrix.
- `model/` – Contains exported `.pkl` pipelines, metrics, and SHAP plots.

## Quick Start
```powershell
# 1. Install dependencies
pip install -r requirements.txt

# 2. Train the advanced classifier (generates model, metrics, and SHAP plots)
python scripts/classification_engine.py --data "data/churn/CHURNDATA (1).xlsx" --output model

# 3. Test the Recommendation Engine (using a real Customer CIF from the dataset)
python scripts/recommendation_engine.py --customer_id XXXXXX

# 4. Generate the Executive Presentation
python scripts/generate_presentation.py
```
