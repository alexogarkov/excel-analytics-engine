# Excel Analytics Engine: From API to Insight

A professional-grade Excel solution that connects directly to the World Bank Open Data API, processes large datasets with Power Query (M), and delivers insights through dynamic dashboards powered by DAX.

## 🔍 Features
- REST API integration via native VBA (no external libraries)
- Alternative Power Query–based API loader
- Aggregation logic using weighted averages for relative indicators
- Semantic data modeling in Power Pivot
- McKinsey-style dashboard with slicers and Top-N logic

## 🧠 Technologies
- Excel (Power Query, Power Pivot, DAX)
- VBA (REST API + JSON)
- M Language
- No Power BI, Python, or external tools required

## 🗂 File Structure
See `/code` for all scripts and `/excel` for the interactive workbook.

## 🚀 Quick Start
1. Open `ExcelAnalyticsEngine.xlsm`
2. Click "Upload World Bank Data"
3. Refresh Power Query if needed
4. Navigate the dashboard

## ⚠️ How to Unblock Macros in Excel
If macros are blocked after downloading this .xlsm file, follow these steps:
- Close Excel completely.
- Right-click the downloaded file → select Properties.
- At the bottom, check "Unblock" under Security.
- Click Apply → OK.
- Reopen the file in Excel. Macros will now run properly.
✅ No external libraries are required — just enable macros from a trusted source.
