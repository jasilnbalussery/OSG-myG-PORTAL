# Detailed Work Summary: Q2 2026 (April, May, June)

This document provides a comprehensive breakdown of the engineering, database optimization, frontend development, and artificial intelligence integration completed across all workspaces during April, May, and June 2026.

---

## 1. OSG-myG-PORTAL (Claims & WhatsApp Engineering)
Significant architectural work was done to stabilize, enhance, and debug the Onsitego claims management system and its integration with external messaging APIs.

* **WhatsApp API Integration (Telinfy/GreenAds Global):**
  * Engineered automated messaging pipelines linking the portal's events to WhatsApp API endpoints.
  * Debugged payload routing and delivery issues, ensuring reliable outbound notifications.
  * Rebuilt data fetching logic utilizing Postman API networks.
* **Claims Processing & Data Forensics:**
  * Developed targeted investigative scripts to trace and resolve specific claim anomalies (e.g., missing follow-ups, stuck statuses).
  * Built and deployed repair utilities to retroactively fix corrupted claim states.
* **Large-Scale Data Ingestion (OSID):**
  * Engineered massive data parsers for legacy Excel imports.
  * Automated data sanitization pipelines to standardize European/American date collisions.
* **Google Apps Script Bridge:**
  * Developed a seamless sync between operational Google Sheets and the Django backend, maintaining data integrity across platforms.

---

## 2. Enterprise AI Agent (Loyalty Portal)
Transformed the Loyalty Analytics Portal into an intelligent, conversational Business Intelligence tool by injecting Large Language Models directly into the database querying pipeline.

* **Multi-Layered Agent Architecture:**
  * Built a **SQL Agent** capable of translating natural language business questions into complex PostgreSQL queries.
  * Built an **AI Analyst** layer to interpret the raw database output and generate readable business insights.
* **LLM Integration & Optimization:**
  * Integrated **NVIDIA's API** endpoints to utilize powerful models for code generation.
  * Overcame severe API latency (reducing 5-10 minute wait times to seconds) through aggressive prompt engineering and query scope reduction.
  * Hardened the AI against database schema hallucinations by explicitly feeding it the database schema context.

---

## 3. Database Administration & Optimization (12.6M+ Rows)
Massive improvements were made to how the Django backend communicates with the DigitalOcean PostgreSQL database, dramatically reducing dashboard load times.

* **Materialized View Overhaul:**
  * Diagnosed a critical cross-join bug in `mv_yearly_cohort` that was duplicating rows and falsely inflating LTV (Lifetime Value) metrics.
  * Built a robust refresh system that successfully regenerates all 26 massive materialized views concurrently without locking up the database.
* **Caching Infrastructure:**
  * Integrated and configured **Redis** and **LocMemCache** to cache heavy API responses, reducing page load times for 12.6M+ row calculations to milliseconds.
  * Fixed issues where local environments were holding onto stale memory caches.
* **DB Manager Enhancements:**
  * Created custom Python scripts to bypass web-server timeouts when uploading massive `DSR MAY 2026` Excel files.
  * Automated the scrubbing of anomalous data, successfully filtering out internal store transactions before they corrupted sales analytics.

---

## 4. Machine Learning & Predictive Forecasting
Replaced static dashboard placeholders with live, mathematically rigorous Python forecasting engines.

* **Model Deployment (Scikit-Learn):**
  * Configured and deployed **Random Forest**, **MLPRegressor** (Neural Network proxy), and **GradientBoostingRegressor** models directly into the API endpoint.
* **Feature Engineering:**
  * Built the highly customized `MalayalamCalendarFeaturizer` to convert dates into multi-dimensional arrays, allowing the AI to factor local Kerala events (like the proximity to *Onam*) into its predictive scoring.
* **Dormant Customer Reactivation System:**
  * Built SQL pre-aggregation engines to accurately bucket customers who purchased in 2024 but went silent in 2025/2026.
  * Designed complex UI visualizers including Plotly **Probability Gauges**, **Semi-Donuts**, and **Dormancy Risk Meters**.

---

## 5. SHE START - Applicant Evaluation Dashboard
Built a complete, isolated dashboard exclusively for managing the "She Start - Her Dreams Start Here" startup program.

* **Live Google Sheets Synchronization:**
  * Utilized the `gspread` library to create a live-syncing engine that perfectly mirrors the applicant Google Sheet to the Django portal every 25 seconds.
* **Advanced Scoring Algorithm:**
  * Programmed a custom mathematical algorithm for 6-panelist scoring: the system automatically identifies and drops the absolute highest and lowest scores, then averages the remaining 4 to prevent bias.
* **Interactive Dashboard UI:**
  * Built interactive, inline editing for the applicant attributes.
  * Engineered a silent saving mechanism to write panelist scores directly to the local Postgres database without page reloads.
  * Implemented an automated badging system that categorizes startups (e.g., *Strong Final Selection*, *Waitlist Consideration*).

---

## 6. Enterprise Retail Dashboard Enhancements
* **High-Speed Reporting:** Swapped out standard `openpyxl` exporters for the Rust-based **`calamine`** engine, drastically accelerating Excel parsing and generation (5-10x faster).
* **DataTables Integration:** Replaced static tables with dynamic, searchable, and sortable DataTables grids.
* **Advanced Metrics:** Integrated complex business logic calculations directly into the Pandas dataframes, including dynamic **Average Selling Price (ASP)** mapping and product-to-category cross-referencing.

---

## 7. BIGG BOSS & FoneFlix Registration Portals
Built a highly customized Flask-based web application initially designed for 'Bigg Boss Season 8 – Agnipareeksha' auditions, which was later successfully pivoted into the 'myG FoneFlix Mobile Phone Short Film Contest 2026'.

* **Video Upload Architecture (Google Drive OAuth):**
  * Engineered a custom OAuth 2.0 Client ID integration that allowed users to authenticate and upload large video files directly into their personal Google Drive folders, bypassing server limits.
* **Dynamic UI/UX Design:**
  * Designed a complex, responsive hero section using precise grid layouts matching OTT reality-show aesthetics.
  * Implemented neon-glow typography, dark glassmorphism form backgrounds, and floating 3D particle animations for brand elements.
* **Project Pivot & Rebranding:**
  * Successfully transitioned the entire codebase from the Bigg Boss branding (Navy Blue & Neon Pink) to the FoneFlix cinematic branding (Dark Browns & Orange).
  * Refactored form inputs, consent clauses, and loading UI states to match the new short film contest requirements.
