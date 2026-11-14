# 📊 SS Demand Planner Pro V21 — Multi-Model Excel VBA & Solver Forecasting Engine
**Author:** Sukhen Singha | Lead Consultant – Supply Chain Planning (16+ Yrs Exp) | Executive PG (IIM Ranchi)  
**LinkedIn:** [linkedin.com/in/sukhen-singha-84773880](https://www.linkedin.com/in/sukhen-singha-84773880)

---
## 🎯 Business Problem Solved
In large-scale FMCG and manufacturing supply chains (3,000+ SKUs across multiple distribution tiers), relying on a single static forecasting formula produces high forecast error (MAPE) because SKUs exhibit vastly different demand behaviors—steady, seasonal, trending, or intermittent/lumpy.

**SS Demand Planner Pro V21** is an automated, macro-enabled Excel & Solver workbench designed to clean historical sales data, run 7 competing statistical models simultaneously, rank them on a live MAPE leaderboard, and automatically feed the winning forecast into multi-echelon inventory planning.

---
## 🚀 Key Features & Technical Architecture
* **4-Click End-to-End Automated Workflow:**  
  1. `1. Update Master` — Ingests SKU sales history and dynamically validates header mappings (evaluates new file uploads fresh whenever column headers change).  
  2. `2. RUN SOLVER` — Executes automated non-linear Solver optimization to tune smoothing parameters (α, β, γ, ϕ) for minimum MAPE.  
  3. `3. Update Inventory` — Feeds the lowest-error forecast into cycle stock and dynamic safety stock calculations.  
  4. `4. Publish Report` — Generates executive S&OP comparison charts and exportable planning tables.
* **7-Model Automated Leaderboard (MAPE % Ranking):**  
  * **Simple Moving Average (SMA)** — Baseline short-horizon stabilizer.  
  * **Single Exponential Smoothing (SES)** — Level-based smoothing for stable SKUs.  
  * **Holt-Winters (Standard Additive/Multiplicative)** — Captures level, trend, and seasonality.  
  * **Holt-Winters (Damped Trend - ϕ)** — Prevents over-forecasting on maturing promotional SKUs.  
  * **Autoregressive AR(1)** — Captures lag-1 serial correlation in high-velocity FMCG lines.  
  * **Croston’s Intermittent Method** — Separates demand size and inter-arrival interval for slow-moving/niche SKUs (achieved **12.90% Best MAPE** on intermittent test SKUs).  
  * **Blended Ensemble (Best-2 Average)** — Combines top two performers to reduce variance risk.
* **Built-In Demand Sensing & Data Hygiene:**  
  * **Outlier Correction (Z-Score Toggle):** Automatically detects and caps abnormal one-off sales spikes.  
  * **Promotional Demand Sensing (ON/OFF):** Adjusts baseline trajectory for planned marketing events.  
  * **Zero-Sales Bridging:** Distinguishes true demand drops from stockout-induced zero sales.  
  * **Time-Window Optimization:** Automatically identifies the optimal historical training window (e.g., Week 79 to Week 156) to prevent outdated structural regimes from skewing current forecasts.

---
## 🛠️ Tech Stack
* **Engine:** Microsoft Excel VBA Macros (`.xlsm`), Excel Solver Optimization Add-in
* **Statistical Methods:** Time-Series Decomposition, Z-Score Normalization, Croston's Method, Holt-Winters, AR(1)
* **Domain Application:** S&OP Demand Validation, DFU-Level Tuning, Multi-Echelon Safety Stock Norms
