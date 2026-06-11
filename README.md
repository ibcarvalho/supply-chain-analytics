
# Supply Chain Analytics — Procurement Analysis (2024)

End-to-end data analytics project focused on supplier performance, procurement spending, delivery delays, and purchasing efficiency.

The dataset was generated in Python to simulate a realistic procurement operation throughout 2024, enabling the application of SQL and Excel analytics techniques in a controlled business environment.

The analysis was developed using both Excel and SQL to compare approaches, validate results, and demonstrate how different tools can support procurement decision-making.

---

## Business Problem

An educational institution needs to monitor supplier performance, purchasing concentration, and delivery reliability to improve procurement planning and reduce operational risks.

The objective of this project is to:

* Identify critical suppliers
* Measure delivery performance against SLAs
* Analyze purchasing concentration
* Detect spending patterns throughout the year
* Generate actionable recommendations for procurement management

---

## Skills Demonstrated

* Data Generation and Simulation
* Data Analysis
* SQL Analytics
* Excel Modeling
* KPI Development
* Supplier Performance Analysis
* ABC Analysis
* Window Functions
* Procurement Analytics
* Business Recommendations

---

## Project Structure

```text
supply-chain-analytics/
│
├── sql/
│   ├── 01.exploracao_inicial.sql
│   ├── 02.analise_fornecedor.sql
│   ├── 03.analise_atraso.sql
│   ├── 04.analise_temporal.sql
│   ├── 05.window_function.sql 
│   └── conclusions_and_recommendations.md
│
├── procurement_analysis.xlsx
├── generate_dataset.py
└── README.md
```

---

## Tools & Technologies

| Tool                   | Purpose                                                                    |
| ---------------------- | -------------------------------------------------------------------------- |
| Python (Pandas, NumPy) | Synthetic dataset generation                                               |
| Excel                  | Lookups, conditional logic, KPI reporting, business analysis               |
| SQL                    | Data exploration, aggregation, joins, CTEs, window functions, ABC analysis |

---

## Dashboard

A visual dashboard is currently being finalized and will be included in a future update of this repository.

The current version already includes:

* Excel-based KPI reporting
* Procurement performance analysis
* Supplier evaluation
* SQL business insights
* Actionable recommendations

---

## Key Metrics (2024)

| Metric                  | Result       |
| ----------------------- | ------------ |
| Total Purchase Orders   | 300          |
| Total Procurement Spend | R$ 3,466,688 |
| On-Time Delivery Rate   | 77.0%        |
| Average Lead Time       | 8.8 days     |
| Delayed Orders          | 69 (23.0%)   |

---

## Key Findings

### Critical Suppliers

LimpaTudo recorded a 100% delay rate throughout 2024, exceeding contracted lead times by an average of 9 days. No purchase order was delivered within the agreed SLA.

EduMaterial (34.2%) and MobiDesk (28.6%) also exceeded the acceptable delay threshold of 20%.

### Spending Concentration

Three categories account for approximately 95% of procurement spending:

* Technology (35.5%)
* Audiovisual Equipment (33.8%)
* Furniture (25.8%)

Within these categories, only five products represent more than 76% of total procurement value.

### Seasonality

Procurement spending peaks occurred in September and December, coinciding with academic cycle closures.

Furniture and Audiovisual Equipment showed the highest lead times, indicating a need for earlier procurement planning during peak periods.

### Structural Delivery Issues

Delay rates remained stable between the first and second semesters, suggesting that delivery performance issues are supplier-related rather than seasonal.

---

## Recommendations

| Priority | Recommendation                                                                            |
| -------- | ----------------------------------------------------------------------------------------- |
| High     | Review supplier performance and initiate alternative supplier qualification for LimpaTudo |
| High     | Negotiate long-term contracts for high-spend products                                     |
| Medium   | Implement monthly monitoring for EduMaterial and MobiDesk                                 |
| Medium   | Anticipate Furniture and Audiovisual purchases before peak demand periods                 |
| Low      | Consolidate low-value purchases to improve purchasing efficiency                          |

Detailed findings are available in:

```text
sql/conclusions_and_recommendations.md
```

---

## How to Run

### Clone the repository

```bash
git clone https://github.com/ibcarvalho/supply-chain-analytics.git
cd supply-chain-analytics
```

### Install dependencies

```bash
pip install pandas numpy openpyxl
```

### Generate the dataset

```bash
python generate_dataset.py
```

### Explore the analysis

* Open `Analise-supply-chain.xlsx`
* Import the CSV files into your preferred SQL environment
* Run the SQL scripts in the `/sql` folder

---

## Future Improvements

* Interactive dashboard
* Power BI version of the analysis
* Additional procurement KPIs
* Forecasting and demand-planning scenarios

---

## Author

**Iago Carvalho**

GitHub: https://github.com/ibcarvalho
