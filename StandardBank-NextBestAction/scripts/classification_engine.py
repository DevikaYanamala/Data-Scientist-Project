"""
Classification Engine for Standard Bank – Next Best Action project
(Advanced Pipeline Edition for MSc / Mid-Level Portfolio)

This script implements a production-grade machine learning pipeline:
1. Data Ingestion: Reads the real CHURNDATA excel file.
2. Preprocessing: Scikit-Learn ColumnTransformer for numerical scaling and categorical encoding.
3. Imbalance Handling: SMOTE (Synthetic Minority Over-sampling Technique) inside an imblearn Pipeline.
4. Modeling: XGBoost Classifier.
5. Tuning: GridSearchCV for hyperparameter optimization.
6. Explainability: SHAP (SHapley Additive exPlanations) values to interpret model decisions.
7. Artifacts: Saves the best model, metrics, and SHAP plots.
"""

import argparse
import json
import logging
import pathlib
from typing import Tuple

import joblib
import numpy as np
import pandas as pd
import shap
import matplotlib.pyplot as plt

from sklearn.compose import ColumnTransformer
from sklearn.metrics import (accuracy_score, confusion_matrix, f1_score,
                             precision_score, recall_score, roc_auc_score, classification_report)
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.pipeline import Pipeline as SklearnPipeline
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from imblearn.pipeline import Pipeline as ImbPipeline
from imblearn.over_sampling import SMOTE
from xgboost import XGBClassifier

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

def load_and_clean_data(file_path: pathlib.Path) -> Tuple[pd.DataFrame, pd.Series]:
    """Loads the real churn Excel dataset and prepares features and target."""
    logger.info(f"Loading data from {file_path}")
    df = pd.read_excel(file_path)
    
    # Target variable mapping
    if "Status" not in df.columns:
        raise ValueError("Dataset must contain a 'Status' column.")
    
    # Map CHURN -> 1, ACTIVE -> 0
    df["churn_label"] = df["Status"].apply(lambda x: 1 if str(x).strip().upper() == "CHURN" else 0)
    y = df["churn_label"]
    
    # Feature Engineering / Selection
    # Drop IDs and date columns that are hard to use directly without extensive FE
    cols_to_drop = ["CIF", "CUS_DOB", "CUS_Customer_Since", "Status", "churn_label"]
    X = df.drop(columns=[c for c in cols_to_drop if c in df.columns])
    
    # Fill basic missing values if any
    for col in X.select_dtypes(include=np.number).columns:
        X[col].fillna(X[col].median(), inplace=True)
    for col in X.select_dtypes(include=['object', 'category']).columns:
        X[col].fillna("Unknown", inplace=True)
        
    logger.info(f"Data shape after cleaning: X={X.shape}, y={y.shape}")
    return X, y

def build_pipeline(X: pd.DataFrame) -> ImbPipeline:
    """Constructs a Scikit-Learn/Imblearn pipeline with preprocessing, SMOTE, and XGBoost."""
    numeric_features = X.select_dtypes(include=np.number).columns.tolist()
    categorical_features = X.select_dtypes(include=['object', 'category']).columns.tolist()
    
    preprocessor = ColumnTransformer(
        transformers=[
            ("num", StandardScaler(), numeric_features),
            ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), categorical_features),
        ]
    )
    
    # We use Imblearn's Pipeline to ensure SMOTE is only applied to training data during CV
    pipeline = ImbPipeline(steps=[
        ("preprocessor", preprocessor),
        ("smote", SMOTE(random_state=42)),
        ("classifier", XGBClassifier(objective="binary:logistic", eval_metric="logloss", random_state=42))
    ])
    
    return pipeline

def train_and_tune(pipeline: ImbPipeline, X_train: pd.DataFrame, y_train: pd.Series):
    """Performs Grid Search CV to find the best hyperparameters."""
    logger.info("Starting GridSearchCV for hyperparameter tuning...")
    
    param_grid = {
        'classifier__n_estimators': [100, 200],
        'classifier__max_depth': [3, 5],
        'classifier__learning_rate': [0.05, 0.1]
    }
    
    grid_search = GridSearchCV(
        pipeline, param_grid, cv=3, scoring='roc_auc', n_jobs=-1, verbose=1
    )
    grid_search.fit(X_train, y_train)
    
    logger.info(f"Best parameters found: {grid_search.best_params_}")
    logger.info(f"Best CV ROC-AUC: {grid_search.best_score_:.4f}")
    
    return grid_search.best_estimator_

def evaluate_model(model, X_test: pd.DataFrame, y_test: pd.Series) -> dict:
    """Evaluates the model on the hold-out test set."""
    logger.info("Evaluating optimal model on test set...")
    proba = model.predict_proba(X_test)[:, 1]
    preds = model.predict(X_test)
    
    metrics = {
        "accuracy": accuracy_score(y_test, preds),
        "roc_auc": roc_auc_score(y_test, proba),
        "precision": precision_score(y_test, preds),
        "recall": recall_score(y_test, preds),
        "f1": f1_score(y_test, preds),
        "confusion_matrix": confusion_matrix(y_test, preds).tolist()
    }
    logger.info(f"\nClassification Report:\n{classification_report(y_test, preds)}")
    logger.info(f"Test Set Metrics: {json.dumps(metrics, indent=2)}")
    return metrics

def generate_shap_explanations(model, X_train: pd.DataFrame, output_dir: pathlib.Path):
    """Generates SHAP values for model explainability."""
    logger.info("Generating SHAP explainability plots...")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Extract the preprocessor and transform the data
    preprocessor = model.named_steps['preprocessor']
    X_train_transformed = preprocessor.transform(X_train)
    
    # Get feature names after one-hot encoding
    cat_encoder = preprocessor.named_transformers_['cat']
    num_features = preprocessor.transformers_[0][2]
    cat_features = preprocessor.transformers_[1][2]
    
    if len(cat_features) > 0:
        cat_feature_names = cat_encoder.get_feature_names_out(cat_features)
        feature_names = num_features + list(cat_feature_names)
    else:
        feature_names = num_features
        
    X_train_df = pd.DataFrame(X_train_transformed, columns=feature_names)
    
    # Extract the XGBoost model
    xgb_model = model.named_steps['classifier']
    
    # Calculate SHAP values (using a sample for speed if dataset is large)
    sample_size = min(500, len(X_train_df))
    X_sample = X_train_df.sample(sample_size, random_state=42)
    
    explainer = shap.TreeExplainer(xgb_model)
    shap_values = explainer.shap_values(X_sample)
    
    # Generate and save Summary Plot
    plt.figure(figsize=(10, 6))
    shap.summary_plot(shap_values, X_sample, feature_names=feature_names, show=False)
    plt.tight_layout()
    plot_path = output_dir / "shap_summary.png"
    plt.savefig(plot_path, dpi=300, bbox_inches='tight')
    plt.close()
    logger.info(f"SHAP summary plot saved to {plot_path}")

def save_artifacts(model, metrics: dict, output_dir: pathlib.Path):
    """Persists the trained pipeline and metrics."""
    model_path = output_dir / "churn_pipeline.pkl"
    metrics_path = output_dir / "metrics.json"
    
    joblib.dump(model, model_path)
    with metrics_path.open("w", encoding="utf-8") as f:
        json.dump(metrics, f, indent=2)
    logger.info(f"Pipeline saved to {model_path}")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", type=pathlib.Path, required=True, help="Path to CHURNDATA Excel file")
    parser.add_argument("--output", type=pathlib.Path, default=pathlib.Path("model"), help="Output directory")
    args = parser.parse_args()

    args.output.mkdir(parents=True, exist_ok=True)

    X, y = load_and_clean_data(args.data)
    
    # Stratified split to maintain class distribution
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )
    
    pipeline = build_pipeline(X_train)
    best_model = train_and_tune(pipeline, X_train, y_train)
    metrics = evaluate_model(best_model, X_test, y_test)
    
    generate_shap_explanations(best_model, X_train, args.output)
    save_artifacts(best_model, metrics, args.output)

if __name__ == "__main__":
    main()
