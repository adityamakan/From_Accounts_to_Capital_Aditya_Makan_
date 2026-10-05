# From Accounts to Capital
### The Rise of India's Retail-Demand Economy After COVID-19

This repository contains the full replication package and supplementary materials for the undergraduate economics research paper **"From Accounts to Capital"** submitted for the Fall 2026 cycle.

## 📁 Repository Structure
* `data/`: Contains verified datasets (`verified_data_ledger.csv`, `verified_data_tables.xlsx`) tracking retail demat holders, mutual fund AUM, and savings distributions.
* `figures/`: Sub-directory storing the 16 analytical charts generated during code execution.
* `build_final_audited.py`: Core architecture script that computes tracking metrics, builds data charts, and compiles the final academic PDF manuscript layout.
* `ABSTRACT.txt`: The summary abstract outlining framework bounds, sources, and scope definitions.

## 🛠️ Execution & Replication Instructions
To verify data structures and fully rebuild the document asset tree, execute the script via your terminal environment.

### 1. Install Prerequisites
Ensure you have the required empirical processing and layout rendering engines installed:
```bash
pip install matplotlib pandas numpy reportlab pypdf openpyxl
```

### 2. Run the Audit Architecture
Run the programmatic workflow to re-verify numbers and dynamically compile the PDF paper:
```bash
python build_final_audited.py
```
