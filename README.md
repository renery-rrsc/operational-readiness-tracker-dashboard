## 📊 Operational Readiness & Project Tracker Dashboard
An interactive, data-driven dashboard built with Python and Streamlit designed to help internal Project Managers from my company to track deliverable adherence, resource planning, and timeline progress. This tool transforms static Excel data into dynamic S-Curves, KPI cards, and delay analysis tables.

## 🚀 Key Features
1. Dynamic S-Curve Visualization: Compares cumulative planned tasks against actual progress with adjustable granularity (Weekly, Monthly, Yearly). It also overlays key project milestones directly on the timeline.

2. Interactive Drill-Downs: Filter data by Area and Work Package to view specific adherence metrics.

3. Automated Delay Detection: Identifies delayed deliverables based on planned finish dates versus current status, classifying them by severity (Minor, Moderate, Major).

4. Visual Status Cards: Color-coded cards providing a quick "at-a-glance" status for specific work packages.

5. KPI Metrics: Real-time calculation of Total Deliverables, Completion Rate, On-Track Percentage, and Total Delayed Items.

## 🛠️ Technical Architecture
The application is structured into two primary classes to separate data processing from UI rendering:

### 1. ProjectDataProcessor
Handles data ingestion and transformation:

Expansion Logic: Converts high-level management plans (mgmtLevel) into weekly operational data points (op_level) by distributing planned quantities over time.

Progress Calculation: Computes actual progress based on status weights (e.g., "Done" = 1.0, "On going" = 0.5).

Delay Logic: Compare planned/actual dates to generate a "Severity" classification for the delay table.

### 2. ProjectDashboard
Manages the Streamlit UI and Plotly figures:

Generates interactive Pie Charts for Area/Package distribution.

Renders the S-Curve using plotly.graph_objects.

Creates CSS-styled HTML cards for the status grid.

## 📦 Installation & Setup
Clone the repository, then type in bash:

  git clone https://github.com/yourusername/project-tracker-dashboard.git
  cd project-tracker-dashboard
  pip install streamlit plotly pandas numpy openpyxl

Configure Data Path Open dashboard.py and update the file path in the main() function to point to your local Excel file:
  
  processor = ProjectDataProcessor("path/to/your/data.xlsx")

Run the Dashboard by calling it in bash:
  
  streamlit run dashboard.py
