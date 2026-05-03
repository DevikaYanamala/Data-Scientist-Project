# classification_engine.py
"""
Classification Engine for Standard Bank – Next Best Action project

This script:
1. Loads a churn dataset (CSV) – path supplied via `--data` argument.
2. Performs a train / test split with a fixed random_state for reproducibility.
3. Trains an XGBoost classifier (high‑performance, handles tabular data well).
4. Evaluates using accuracy, ROC‑AUC, precision, recall, F1.
5. Saves the trained model to `model/churn_classifier.pkl`.
6. Logs the results to a JSON file (`model/metrics.json`).

The script is deliberately written for clarity, type hints and logging so it can be
included directly in a portfolio or demo.
"""

import argparse
import json
import logging
import pathlib
from typing import Tuple

import joblib
import numpy as np
import pandas as pd
from sklearn.metrics import (accuracy_score, confusion_matrix, f1_score,
                             precision_score, recall_score, roc_auc_score)
from sklearn.model_selection import train_test_split
from xgboost import XGBClassifier

# ---------------------------------------------------------------------------
# Logging configuration – simple console output, suitable for a portfolio demo
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s – %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def load_data(csv_path: pathlib.Path) -> Tuple[pd.DataFrame, pd.Series]:
    """Load CSV and separate features / target.

    The CSV is expected to contain a column named `churn` (1 = churn, 0 = retained).
    All other columns are treated as features. Categorical columns are one‑hot
    encoded using pandas `get_dummies`.
    """
    logger.info(f"Loading data from {csv_path}")
    df = pd.read_csv(csv_path)
    if "churn" not in df.columns:
        raise ValueError("Dataset must contain a 'churn' column as the target")
    y = df["churn"].astype(int)
    X = df.drop(columns=["churn"])
    # Simple one‑hot encoding for any object dtype columns
    X = pd.get_dummies(X, drop_first=True)
    logger.info(f"Dataset shape after encoding: {X.shape}")
    return X, y


def train_classifier(X: pd.DataFrame, y: pd.Series) -> XGBClassifier:
    """Train an XGBoost classifier with sensible defaults.

    A fixed `random_state` ensures reproducible results across runs.
    """
    logger.info("Training XGBoost classifier")
    model = XGBClassifier(
        n_estimators=300,
        max_depth=5,
        learning_rate=0.05,
        subsample=0.9,
        colsample_bytree=0.9,
        objective="binary:logistic",
        eval_metric="logloss",
        n_jobs=-1,
        random_state=42,
        use_label_encoder=False,
    )
    model.fit(X, y)
    return model


def evaluate(model: XGBClassifier, X_test: pd.DataFrame, y_test: pd.Series) -> dict:
    """Generate a dictionary of evaluation metrics.
    """
    logger.info("Evaluating model on test set")
    proba = model.predict_proba(X_test)[:, 1]
    preds = (proba >= 0.5).astype(int)
    metrics = {
        "accuracy": accuracy_score(y_test, preds),
        "roc_auc": roc_auc_score(y_test, proba),
        "precision": precision_score(y_test, preds, zero_division=0),
        "recall": recall_score(y_test, preds, zero_division=0),
        "f1": f1_score(y_test, preds, zero_division=0),
        "confusion_matrix": confusion_matrix(y_test, preds).tolist(),
    }
    logger.info(f"Metrics: {json.dumps(metrics, indent=2)}")
    return metrics


def save_artifacts(model: XGBClassifier, metrics: dict, output_dir: pathlib.Path) -> None:
    """Persist the trained model and metrics.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    model_path = output_dir / "churn_classifier.pkl"
    metrics_path = output_dir / "metrics.json"
    joblib.dump(model, model_path)
    logger.info(f"Model saved to {model_path}")
    with metrics_path.open("w", encoding="utf-8") as f:
        json.dump(metrics, f, indent=2)
    logger.info(f"Metrics saved to {metrics_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Train churn classification model")
    parser.add_argument("--data", type=pathlib.Path, required=True, help="Path to churn CSV dataset")
    parser.add_argument("--output", type=pathlib.Path, default=pathlib.Path("model"), help="Directory to store model and metrics")
    args = parser.parse_args()

    X, y = load_data(args.data)
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )
    model = train_classifier(X_train, y_train)
    metrics = evaluate(model, X_test, y_test)
    save_artifacts(model, metrics, args.output)


if __name__ == "__main__":
    main()
