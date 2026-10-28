# MNC-Ready Data Science & Data Analyst Project

> **SYNTHETIC DEMONSTRATION — NOT REAL COMPANY DATA.** This is a reproducible demonstration created only because synthetic data was explicitly requested. It is not a real company dataset and no real-world business conclusions should be drawn from it.

## Deliverables

- `01_Manual_Statistics_Probability.pdf` — manual-first calculations for every standalone statistics/probability question.
- `02_Cleaned_Data.xlsx` — cleaned dimension/fact sheets, enriched order table, quality checks and data dictionary.
- `03_Excel_Management_Dashboard.xlsx` — KPI formulas with dropdown/date controls, charts and pivot-ready group summaries. Includes Power Query steps/M snippets and DAX measures in Build Notes.
- `04_SQL_Assessment.sql` — commented PostgreSQL solutions for the listed SQL topics.
- `05_Python_Foundations.ipynb` — core Python, functions, file I/O, OOP, exceptions and context manager.
- `06_Pandas_NumPy_EDA.ipynb` — profiling, cleaning, joins, feature derivation, manual-stat validation, required NumPy operations and eight-chart EDA.
- `07_Feature_Engineering.ipynb` — RFM, monthly features, bins, encoding/scaling, ColumnTransformer, Pipeline, evaluation demo and joblib persistence.
- `08_Streamlit_App/` — interactive dashboard app, data dependency requirements and this README's deployment instructions.
- `09_Executive_Report.pdf` and `.docx` — summary, caveats, recommendations and 10-minute interview outline.
- `10_README.md` — this file. Supporting CSVs, charts and outputs are included.

## Synthetic data and assumptions

- Fixed random seed: `27102026`. Period: 2024-01-01 to 2025-12-31. Tables: customers, orders, products, payments, delivery and support. Raw exports intentionally include small amounts of whitespace, missing demographics, and duplicate rows for cleaning practice.
- `data/raw/` is deliberately imperfect. `data/cleaned/` is normalized, trimmed and deduplicated, with reproducible derived measures. No source system data is represented.
- Recognized revenue = order value only when `order_status == Completed` and `payment_status == Success`. `profit` = recognized revenue − quantity × cost_price; it is a gross contribution proxy, not accounting profit. Returned orders are excluded from recognized revenue; returns remain in the denominator of overall return rate.
- On-time delivery rate = On time / (On time + Late). Churn is a generated snapshot label at 2025-12-31. Synthetic churn probability was constructed with correlated late delivery/returns/support factors; the association analysis is descriptive, not causal.
- Payment composite rate = Failed or Refunded / payment records. Failed and refunded rates are also reported separately.
- RFM uses completed, successful orders as of 2025-12-31. Missing RFM activity is set to zero frequency/spend and long recency for a no-purchase customer.
- Employee-manager hierarchy data was not in the suggested input schema; SQL includes a standalone recursive CTE template for an `employees` table.

## Rebuild and run

Python 3.10+ is recommended. From the project root:

```bash
pip install pandas numpy matplotlib seaborn scikit-learn joblib openpyxl xlsxwriter nbformat nbclient ipykernel python-docx
python build_mnc_demo.py
````

Open notebooks in Jupyter and run top to bottom. Notebook data paths work from project root or `notebooks/`. `04_SQL_Assessment.sql` targets PostgreSQL; replace `:revenue_threshold` with a bound parameter in the client.

## Streamlit Cloud deployment

The repository root should contain `app/streamlit_app.py`, `app/requirements.txt`, and the `data/cleaned/` directory. In Streamlit Community Cloud, choose `app/streamlit_app.py` as the app entry point and include project data files in the connected repository. From a local environment run `streamlit run app/streamlit_app.py`. The dashboard offers multi-select category/city/segment filters and date range controls, KPI cards, four interactive Plotly charts, a filtered dataframe and CSV download.

## Excel feature notes

The dashboard workbook uses live Excel formulas (SUMIFS/COUNTIFS/IFERROR), list/date validation and native charts. XlsxWriter does not author native PivotTable caches/slicers or embed executable Power Query M queries. For transparency, the workbook includes pivot-ready grouped tables, a Power Query applied-step recipe/M snippets and DAX measure examples. In desktop Excel, create a native PivotTable from the `Orders` sheet table and add slicers/timeline if required. The formulas are pre-calculated/cached for the default all-data view; Excel recalculates after filter control changes.

## Key demo findings

- Recognized revenue: **$976,190**; recognized gross contribution proxy: **$207,580**; AOV: **$160.08**.
- Highest revenue segment: **Core**; revenue **$436,351**, profit **$100,919**.
- Lowest simulated category contribution margin: **Electronics**, **7.6%**.
- Synthetic churn: customers with any late delivery **11.7%** vs **7.4%** without; not causal.
- Highest combined failed/refunded payment method: **Wallet**, **14.1%**.
- Slowest-resolving support issue: **Refund/payment**, mean **4.6 days**.

## QA and limitations

Data joins validate many-to-one dimensions; raw/clean table row counts and nulls are included in `02_Cleaned_Data.xlsx`. Statistical calculations are manually displayed in the PDF and recalculated in the notebook afterward. Business outputs are deterministic with the fixed seed. The logistic regression is only a learning exercise; its target construction induces feature correlations and does not establish generalization. Use real data, approved definitions, time-based validation, privacy review and experimentation before any operational decision.
