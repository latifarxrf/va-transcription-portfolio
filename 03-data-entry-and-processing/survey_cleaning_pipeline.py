#!/usr/bin/env python3
"""
Automated Survey Data Cleaning & Quality Control Pipeline
Author: Latifar (@latifarxrf)
Description:
    Processes raw, noisy survey responses (N=2,500).
    Performs de-duplication, speeder/careless respondent detection,
    straight-lining evaluation, logic consistency checks, and produces
    clean, analysis-ready datasets with structured audit logs.
"""

import sys
import pandas as pd
import numpy as np
from datetime import datetime

def clean_survey_dataset(input_file: str, output_file: str, log_file: str = "cleaning_audit_log.txt"):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting Survey Cleaning Pipeline...")
    audit_log = []
    
    # 1. Load Raw Dataset
    try:
        df = pd.read_csv(input_file)
        initial_count = len(df)
        audit_log.append(f"Initial raw record count: {initial_count}")
        print(f"-> Loaded {initial_count} rows from {input_file}")
    except Exception as e:
        print(f"Error loading file: {e}")
        return

    # 2. De-duplication by Respondent ID
    if 'respondent_id' in df.columns:
        df_unique = df.drop_duplicates(subset=['respondent_id'], keep='first')
        dupes_removed = initial_count - len(df_unique)
        audit_log.append(f"Duplicates identified and purged: {dupes_removed}")
        print(f"-> Removed {dupes_removed} duplicates")
        df = df_unique

    # 3. Speeder Detection (Completion duration < 25% of median duration)
    if 'duration_seconds' in df.columns:
        median_duration = df['duration_seconds'].median()
        speed_threshold = median_duration * 0.25
        speeders = df[df['duration_seconds'] < speed_threshold]
        audit_log.append(f"Median duration: {median_duration:.1f}s | Speed threshold: {speed_threshold:.1f}s")
        audit_log.append(f"Careless speeders eliminated: {len(speeders)}")
        print(f"-> Removed {len(speeders)} speeders (took < {speed_threshold:.1f}s)")
        df = df[df['duration_seconds'] >= speed_threshold]

    # 4. Straight-Lining Evaluation (Zero variance across Likert scale items)
    likert_columns = [col for col in df.columns if col.startswith('Q_likert_')]
    if likert_columns:
        df['likert_variance'] = df[likert_columns].var(axis=1)
        straight_liners = df[df['likert_variance'] == 0.0]
        audit_log.append(f"Identified {len(straight_liners)} straight-line responders across Likert matrix")
        print(f"-> Filtered out {len(straight_liners)} straight-liners")
        df = df[df['likert_variance'] > 0.0].drop(columns=['likert_variance'])

    # 5. Logic Consistency Check
    # Rule: If respondent states 'Never used product' (Q_usage == 1), Satisfaction (Q_sat) cannot be > 1
    if 'Q_usage' in df.columns and 'Q_satisfaction' in df.columns:
        logic_violations = df[(df['Q_usage'] == 1) & (df['Q_satisfaction'] > 1)]
        audit_log.append(f"Logical contradictions flagged: {len(logic_violations)}")
        # Impute contradiction with NaN or benchmark score
        df.loc[(df['Q_usage'] == 1) & (df['Q_satisfaction'] > 1), 'Q_satisfaction'] = np.nan
        print(f"-> Rectified {len(logic_violations)} logic contradictions")

    # 6. Final Export & Audit Trail Generation
    final_count = len(df)
    retention_rate = (final_count / initial_count) * 100
    audit_log.append(f"Final clean records retained: {final_count} ({retention_rate:.2f}%)")
    
    df.to_csv(output_file, index=False)
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("\n".join(audit_log))
        
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Pipeline Complete!")
    print(f"Clean data saved to: {output_file}")
    print(f"Audit log saved to: {log_file}")
    print(f"Retention rate: {retention_rate:.2f}%\n")

if __name__ == "__main__":
    print("Survey Data Cleaning Pipeline Module — Ready for Execution.")
