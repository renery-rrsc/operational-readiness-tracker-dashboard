import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from datetime import datetime
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import uuid
from dateutil.relativedelta import relativedelta

### ================> DATA PROCESSOR <================ ###
class ProjectDataProcessor:
    def __init__(self, excel_file_path):
        """Initialize the program with excel file contents"""
        self.excel_file = excel_file_path
        self.mgmtLevel = None
        self.absolute = None
        self.plan_status = None
        self.milestones = None

    def load_data(self):
        """Load the excel sheet"""
        print("Loading data from excel...")

        self.mgmtLevel = pd.read_excel(self.excel_file, sheet_name='mgmtLevel')
        self.absolute = pd.read_excel(self.excel_file, sheet_name='absolute')
        self.plan_status = pd.read_excel(self.excel_file, sheet_name='planStatus')
        self.milestones = pd.read_excel(self.excel_file, sheet_name='milestones')

        print(f"Loaded {len(self.mgmtLevel)} records from table 'mgmtLevel'")
        return self
    
    def create_mgmt_level_ids(self):
        """Create unique IDs for the mgmtLevel table"""
        self.mgmtLevel['mgmt_ID'] = [f"MGMT_{i:05d}" for i in range(len(self.mgmtLevel))]

        print("Created management level IDs")
        return self
    
    def expand_to_op_level(self):
        """Expand the mgmtLevel table to operation level by distributing the quantities over months and weeks"""
        print("Expanding mgmtLevel to opLevel...")

        expanded_records = []
        for idx, row in self.mgmtLevel.iterrows():
            try:
                start_date = datetime(int(row['Planned Start Year']), int(row['Planned Start Month']), 1)
                end_date = datetime(int(row['Planned Finish Year']), int(row['Planned Finish Month']), 1)

                if end_date.month == 12:
                    end_date = end_date.replace(year=end_date.year + 1, month=1) - timedelta(days=1)
                else:
                    end_date = end_date.replace(month=end_date.month + 1) - timedelta(days=1)

                total_days = (end_date - start_date).days + 1
                total_weeks = max(1, int(np.ceil(total_days/7)))
                d_quantityPlanned = row['quantityPlanned'] / total_weeks

                current_date = start_date
                week_counter = 0

                while current_date <= end_date and week_counter < total_weeks:
                    week_num = current_date.isocalendar()[1]
                    month_num = current_date.month
                    year_num = current_date.year

                    expanded_record = {
                        'Area': row['Area'],
                        'Package': row['Package'],
                        'WorkTrack': row['WorkTrack'],
                        'Deliverable Type': row['Deliverable Type'],
                        'Deliverable Name': row['Deliverable Name'],
                        'Task': row['Task'],
                        'status': row['status'],
                        'mgmt_ID': row['mgmt_ID'],
                        #new fields
                        'op_id': f"OP_{uuid.uuid4().hex[:8].upper()}",
                        'week_num' : week_num if week_num <= 52 else 52,
                        'month_num' : month_num,
                        'year_num' : year_num,
                        'distributed_quantity': d_quantityPlanned,
                        'cumulative_planned': 0,
                        'completion' : None
                    }

                    expanded_records.append(expanded_record)
                    current_date += timedelta(weeks=1)
                    week_counter += 1

            except Exception as e:
                print(f"Error processing row {idx}: {e}")
                continue

            self.op_level = pd.DataFrame(expanded_records)

        if len(self.op_level) > 0:
            self.op_level = self.op_level.sort_values(['year_num', 'month_num', 'week_num'])
            self.op_level['cumulative_planned'] = self.op_level['distributed_quantity'].cumsum()
            print(f"Successfully to {len(self.op_level)} operation level records")
            print("\nSample of expanded data: ")
            print(self.op_level.head(10))
        else:
            print("No operation level records were created.")

        return self
        
    def _get_weeks_in_month(self, date):
        """Calculate number of weeks in a given month"""
        days_in_month = (date.replace(month=date.month % 12 + 1, day = 1) - timedelta(days=1)).day
        return 5 if days_in_month > 30 else 4
        
    def calculate_actual_progress(self):
        """Calculate actual progress based on status and actual dates"""
        print("Calculating actual progress...")

        status_completion = {
            'Done' : 1.0,
            'On going' : 0.5,
            'Not started' : 0.0,
            'Delayed' : 0.3
        }

        self.op_level['completion_rate'] = self.op_level['status'].map(status_completion).fillna(0)
        self.op_level['actual_quantity'] = self.op_level['distributed_quantity']*self.op_level['completion_rate']
        self.op_level['cumulative_actual'] = self.op_level['actual_quantity'].cumsum()

        return self
    
    def identify_delays(self):
        """Identify delayed deliverables for summary table"""
        print("Identifying delays... ")

        def safe_date_creation(year_col, month_col, default_day=1):
            """To hande NaN values"""
            dates = []
            for year, month in zip(year_col, month_col):
                try:
                    if pd.notna(year) and pd.notna(month):
                        year_int = int(float(year))
                        month_int = int(float(month))
                        if 1 <= month_int <= 12 and year_int > 1900:
                            dates.append(datetime(year_int, month_int, default_day))
                        else:
                            dates.append(pd.NaT)
                    else:
                        dates.append(pd.NaT)
                except (ValueError, TypeError, OverflowError):
                    dates.append(pd.NaT)
            return pd.Series(dates)

        self.mgmtLevel['planned_finish_date'] = safe_date_creation(
            self.mgmtLevel['Planned Finish Year'].astype(str),
            self.mgmtLevel['Planned Finish Month'].astype(str)
        )

        self.mgmtLevel['actual_finish_date'] = safe_date_creation(
            self.mgmtLevel['Actual Finish Year'].astype(str),
            self.mgmtLevel['Actual Finish Month'].astype(str)
        )

        self.mgmtLevel['planned_start_date'] = safe_date_creation(
            self.mgmtLevel['Planned Start Year'],
            self.mgmtLevel['Planned Start Month']
        )

        current_date = datetime.now()

        delayed_conditions = (
            (self.mgmtLevel['status'] == 'Delayed') |
            (
                (self.mgmtLevel['planned_finish_date'].notna()) &
                (self.mgmtLevel['planned_finish_date'] < current_date) &
                (self.mgmtLevel['status'] != 'Done')
            ) |
            (self.mgmtLevel['actual_finish_date'].notna()) &
            (self.mgmtLevel['planned_finish_date'].notna()) &
            (self.mgmtLevel['actual_finish_date'] > self.mgmtLevel['planned_finish_date'])
        )

        self.delayed_deliverables = self.mgmtLevel[delayed_conditions].copy()

        if len(self.delayed_deliverables) > 0:

            self.delayed_deliverables['planned_week_num'] = self.delayed_deliverables['planned_start_date'].apply(
                lambda x: x.isocalendar()[1] if pd.notna(x) else None
            )

            self.delayed_deliverables['planned_start_combined'] = self.delayed_deliverables.apply(
                lambda row: f"{int(row['Planned Start Year'])}-{int(row['Planned Start Month']):02d}"
                if pd.notna(row['Planned Start Year']) and pd.notna(row['Planned Start Month'])
                else None, axis=1
            )

            mask_both_dates = (
                self.delayed_deliverables['actual_finish_date'].notna() & 
                self.delayed_deliverables['planned_finish_date'].notna()
            )
        
            self.delayed_deliverables.loc[mask_both_dates, 'delay_days'] = (
                self.delayed_deliverables.loc[mask_both_dates, 'actual_finish_date'] - 
                self.delayed_deliverables.loc[mask_both_dates, 'planned_finish_date']
            ).dt.days
        
            mask_overdue = (
                self.delayed_deliverables['planned_finish_date'].notna() & 
                self.delayed_deliverables['actual_finish_date'].isna() &
                (self.delayed_deliverables['planned_finish_date'] < current_date) &
                (self.delayed_deliverables['status'] != 'Done')
            )
        
            self.delayed_deliverables.loc[mask_overdue, 'delay_days'] = (
                current_date - self.delayed_deliverables.loc[mask_overdue, 'planned_finish_date']
            ).dt.days
        
            self.delayed_deliverables['delay_days'] = self.delayed_deliverables['delay_days'].fillna(0)
        
            def classify_delay_severity(delay_days):
                if pd.isna(delay_days) or delay_days <= 0:
                    return 'No Delay'
                elif delay_days <= 30:
                    return 'Minor Delay'
                elif delay_days <= 90:
                    return 'Moderate Delay'
                else:
                    return 'Major Delay'
            
            self.delayed_deliverables['delay_severity'] = self.delayed_deliverables['delay_days'].apply(classify_delay_severity)
    
        print(f"Identified {len(self.delayed_deliverables)} delayed deliverables.")
        return self
    
    def process_all(self):
        """Run complete data processing pipeline"""
        return (self.load_data()
                    .create_mgmt_level_ids()
                    .expand_to_op_level()
                    .calculate_actual_progress()
                    .identify_delays()
                )

### ================> DASHBOARD GENERATOR <================ ###
class ProjectDashboard:
    def __init__(self, processor):
        self.processor = processor
        self.color_scheme = {
            'Done': '#28a745',
            'On going': '#ffc107', 
            'Not Started': '#dc3545',
            'Delayed': '#fd7e14',
            'Not Planned': '#6c757d'
        }

    def create_areas_pie_chart(self):
        """Create the main Areas distribution pie chart"""
        print("Creating Areas pie chart...")
        dept_totals = self.processor.absolute.groupby('Areas')['Total'].sum().reset_index()
        
        package_breakdown = {}

        for area in dept_totals['Areas'].unique():
            area_packages = self.processor.absolute[self.processor.absolute['Areas'] == area]
            area_packages_sorted = area_packages.sort_values('Total', ascending=False)

            breakdown_lines =[]
            for _, pkg in area_packages_sorted.iterrows():
                breakdown_lines.append(f"{pkg['Packages']}: {pkg['Total']} deliverables")

            package_breakdown[area] = "<br>".join(breakdown_lines)

        dept_totals['package_breakdown'] = dept_totals['Areas'].map(package_breakdown)

        fig = px.pie(
            dept_totals,
            values='Total',
            names='Areas',
            title='Documents by Area<br><sub>Hover to see package breakdown</sub>',
            color_discrete_sequence=px.colors.qualitative.Set3,
            hole=0.3
        )
        
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label+value',
            hovertemplate='<b>%{label}</b><br>' +
                         'Total Deliverables: %{value}<br>' +
                         'Percentage: %{percent}<br>' +
                         '<br><b>Package Breakdown:</b><br>' +
                         ' %{customdata}<br>' +
                         '<extra></extra>',
            customdata=dept_totals['package_breakdown']
        )
        
        fig.update_layout(
            height=500,
            showlegend=True,
            font=dict(size=12),
            title_x=0.5
        )
        
        print("Areas pie chart created.")
        return fig
    
    def create_packages_pie_chart(self, selected_area=None):
        """Create packages pie chart for selected area or all areas"""
        if selected_area and selected_area != 'All Areas':
            package_data = self.processor.absolute[self.processor.absolute['Areas'] == selected_area]
            title = f'Package Distribution - {selected_area}'
        else:
            package_data = self.processor.absolute
            title = 'All Packages Distribution'
        
        if len(package_data) == 0:
            # Create empty chart
            fig = go.Figure()
            fig.update_layout(
                title=f"No data available for {selected_area}",
                height=500
            )
            return fig
        
        fig = px.pie(
            package_data,
            values='Total',
            names='Packages',
            title=title,
            color_discrete_sequence=px.colors.qualitative.Pastel,
            hole=0.3
        )
        
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label+value',
            hovertemplate='<b>%{label}</b><br>' +
                         'Deliverables: %{value}<br>' +
                         'Percentage: %{percent}<br>' +
                         '<extra></extra>'
        )
        
        fig.update_layout(
            height=500,
            showlegend=True,
            font=dict(size=12),
            title_x=0.5
        )
        
        return fig
    
    def create_status_cards(self, area_filter="All Areas", package_filter="All Packages"):
        """Visualization 2: cards with quantity of deliverables with status colors"""
        print("Creating status cards...")

        cards_data = []
        filtered_plan_status = self.processor.plan_status.copy()
        if area_filter != "All Areas":
            filtered_plan_status = filtered_plan_status[filtered_plan_status['Area'] == area_filter]

        if package_filter != "All Packages":
            filtered_plan_status = filtered_plan_status[filtered_plan_status['Package'] == package_filter]

        for _, row in filtered_plan_status.iterrows():
            status_row = self.processor.plan_status[
                (self.processor.plan_status['Area'] == row['Area']) &
                (self.processor.plan_status['Package'] == row['Package'])
            ]

            status = status_row['Status'].iloc[0] if not status_row.empty else 'Not Planned'

            cards_data.append({
                'area' : row['Area'],
                'package' : row['Package'],
                'status' : status,
                'color' : self.color_scheme.get(status, '#6c757d')
            })

        print("Status cards created.")
        return cards_data
    
    def create_s_curve(self, timeline_view="Monthly", area_filter="All Areas", package_filter="All Packages"):
        """Visualization 3: S-curve with planned vs actual progress"""
        print(f"Creating S-curve with {timeline_view} resolution, Area: {area_filter}, Package: {package_filter}")
        
        if not hasattr(self.processor, 'op_level') or self.processor.op_level is None:
            print("Warning: op_level not found, creating empty S-curve.")
            fig = go.Figure()
            fig.update_layout(title="S-Curve: No Data Available")
            return fig
        
        filtered_data = self.processor.op_level.copy()
        
        if area_filter != "All Areas":
            filtered_data = filtered_data[filtered_data['Area'] == area_filter]
        
        if package_filter != "All Packages":
            filtered_data = filtered_data[filtered_data['Package'] == package_filter]
        
        if len(filtered_data) == 0:
            fig = go.Figure()
            fig.update_layout(
                title=f"S-Curve: No Data Available for {area_filter} - {package_filter}",
                height=600
            )
            return fig
        
        if timeline_view == "Weekly":
            timeline_data = filtered_data.groupby(['year_num', 'week_num']).agg({
                'cumulative_planned': 'max',
                'distributed_quantity': 'sum'
            }).reset_index()
            
            timeline_data['timeline_label'] = timeline_data['week_num'].astype(str) + '-' + timeline_data['year_num'].astype(str)
            timeline_data = timeline_data.sort_values(['year_num', 'week_num'])
            x_axis_title = 'Timeline (Week-Year)'
            
        elif timeline_view == "Monthly":
            timeline_data = filtered_data.groupby(['year_num', 'month_num']).agg({
                'cumulative_planned': 'max',
                'distributed_quantity': 'sum'
            }).reset_index()
            
            timeline_data['timeline_label'] = timeline_data['month_num'].astype(str) + '-' + timeline_data['year_num'].astype(str)
            timeline_data = timeline_data.sort_values(['year_num', 'month_num'])
            x_axis_title = 'Timeline (Month-Year)'
            
        else:
            timeline_data = filtered_data.groupby(['year_num']).agg({
                'cumulative_planned': 'max',
                'distributed_quantity': 'sum'
            }).reset_index()
            
            timeline_data['timeline_label'] = timeline_data['year_num'].astype(str)
            timeline_data = timeline_data.sort_values(['year_num'])
            x_axis_title = 'Timeline (Year)'
        
        timeline_data['cumulative_planned'] = timeline_data['distributed_quantity'].cumsum()
        
        status_completion = {
            'Done': 1.0,
            'On going': 0.5,
            'Not Started': 0.0,
            'Delayed': 0.3
        }
        timeline_data['cumulative_actual'] = timeline_data['cumulative_planned'] * 0.7
        timeline_data['planned_display'] = np.ceil(timeline_data['cumulative_planned']).astype(int)
        timeline_data['actual_display'] = np.ceil(timeline_data['cumulative_actual']).astype(int)
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=timeline_data['timeline_label'],
            y=timeline_data['cumulative_planned'],
            mode='lines+markers',
            name='Planned Tasks',
            line=dict(color='blue', width=3),
            customdata=timeline_data['planned_display'],
            hovertemplate='<b>Planned</b><br>Period: %{x}<br>Tasks: %{customdata}<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=timeline_data['timeline_label'],
            y=timeline_data['cumulative_actual'],
            mode='lines+markers',
            name='Actual Progress',
            line=dict(color='green', width=3),
            customdata=timeline_data['actual_display'],
            hovertemplate='<b>Actual</b><br>Period: %{x}<br>Tasks: %{customdata}<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=timeline_data['timeline_label'].tolist() + timeline_data['timeline_label'].tolist()[::-1],
            y=timeline_data['cumulative_planned'].tolist() + timeline_data['cumulative_actual'].tolist()[::-1],
            fill='tonexty',
            fillcolor='rgba(255, 0, 0, 0.2)',
            line=dict(color='rgba(255, 255, 255, 0)'),
            name='Delay Gap',
            hoverinfo='skip'
        ))
        
        if hasattr(self.processor, 'milestones') and self.processor.milestones is not None and len(self.processor.milestones) > 0:
            print(f"Debug: Found {len(self.processor.milestones)} milestones")
            print(f"Debug: Milestone columns: {list(self.processor.milestones.columns)}")
            print(f"Debug: Timeline labels available: {timeline_data['timeline_label'].tolist()}")
            
            filtered_milestones = self.processor.milestones.copy()
            
            # Apply filters
            if area_filter != "All Areas":
                filtered_milestones = filtered_milestones[filtered_milestones['Area'] == area_filter]
            
            if package_filter != "All Packages":
                filtered_milestones = filtered_milestones[filtered_milestones['Package'] == package_filter]
            
            if len(filtered_milestones) > 0:
                milestones_colors = [
                    '#FF0000',
                    '#FF8C00',
                    '#8A2BE2',
                    '#A52A2A',
                    '#FF1493',
                    '#00CED1',
                    '#32CD32',
                    '#FFD700',
                    '#DC143C',
                    '#4169E1'
                ]
                max_y = max(timeline_data['cumulative_planned'].max(), timeline_data['cumulative_actual'].max())
                milestone_count = 0
                
                for idx, milestone in filtered_milestones.iterrows():
                    try:
                        year = milestone['Year']
                        month = milestone['Month']
                        day = milestone['Day']
                        
                        if pd.notna(year) and pd.notna(month) and pd.notna(day):
                            year_int = int(float(str(year)))
                            month_int = int(float(str(month)))
                            day_int = int(float(str(day)))
                            
                            if year_int > 1900 and 1 <= month_int <= 12 and 1 <= day_int <= 31:
                                milestone_date = datetime(year_int, month_int, day_int)
                                
                                if timeline_view == "Weekly":
                                    milestone_week = milestone_date.isocalendar()[1]
                                    milestone_x = f"{milestone_week}-{milestone_date.year}"
                                elif timeline_view == "Monthly":
                                    milestone_x = f"{milestone_date.month}-{milestone_date.year}"
                                else:
                                    milestone_x = str(milestone_date.year)
                                
                                if milestone_x in timeline_data['timeline_label'].values:
                                    milestone_color = milestones_colors[milestone_count % len(milestones_colors)]
                                    
                                    milestone_name = f"{milestone['Name']}-{milestone['Package']}"
                                    
                                    print(f"Debug: Adding milestone line at {milestone_x} with color {milestone_color}")
                                    
                                    fig.add_vline(
                                        x=str(milestone_x),
                                        line_dash='dash',
                                        line_color=milestone_color,
                                        line_width=2,
                                        #annotation_text=str(milestone_name),
                                        #annotation_position='top'
                                    )

                                    fig.add_annotation(
                                        x=str(milestone_x),
                                        y=max_y * 1.1,
                                        text=f"{milestone_name}",
                                        showarrow=False,
                                        font=dict(size=9, color="white"),
                                        bgcolor=milestone_color,
                                        bordercolor=milestone_color,
                                        borderwidth=1,
                                        xanchor="center",
                                        yanchor="bottom"
                                    )
                                    
                                    milestone_count += 1
                                else:
                                    print(f"Debug: Milestone {milestone_x} not found in timeline range")
                            else:
                                print(f"Invalid date components for milestone {milestone.get('Name', 'Unknown')}: {year_int}-{month_int}-{day_int}")
                        else:
                            print(f"Missing date data for milestone {milestone.get('Name', 'Unknown')}")
                            
                    except (ValueError, TypeError, KeyError) as e:
                        print(f"Error processing milestone {milestone.get('Name', 'Unknown')}: {e}")
                        print(f"Milestone data: Year={milestone.get('Year')}, Month={milestone.get('Month')}, Day={milestone.get('Day')}")
                        continue
            else:
                print("No milestones remain after filtering")
        else:
            print("No milestones table found or table is empty")
        
        title_parts = ['S-Curve: Planned Tasks vs Actual Progress']
        if area_filter != "All Areas":
            title_parts.append(f"Area: {area_filter}")
        if package_filter != "All Packages":
            title_parts.append(f"Package: {package_filter}")
        title_parts.append(f"({timeline_view})")
        
        fig.update_layout(
            title=' | '.join(title_parts),
            xaxis_title=x_axis_title,
            yaxis_title='Quantity of Tasks',
            height=600,
            hovermode='x unified',
            legend=dict(x=0.02, y=0.98),
            title_x=0.5
        )
        
        print("S-curve created successfully.")
        return fig
    
    def create_delay_table(self):
        """Visualization 4: table with delayed deliverables"""
        print("Creating X9 table...")

        if not hasattr(self.processor, 'delayed_deliverables') or len(self.processor.delayed_deliverables) == 0:
            return pd.DataFrame()
        
        summary_columns = [
            'Area',
            'Package',
            'Deliverable Name',
            'Task',
            'planned_start_combined',
            'planned_week_num',
            'delay_severity'
        ]

        delay_summary = self.processor.delayed_deliverables[summary_columns].copy()
        delay_summary = delay_summary.sort_values(['Area', 'Package','planned_start_combined']).reset_index(drop=True)

        delay_summary = delay_summary.rename(columns={
            'planned_start_combined': 'Planned Start',
            'planned_week_num': 'Planned Week Number',
            'delay_severity': 'Severity'
        })

        print("X9 table created successfully.")
        return delay_summary
    
    def display_cards_streamlit(self, cards_data, area_filter="All Areas", package_filter="All Packages"):
        """Display cards in Streamlit with style"""

        if len(cards_data) == 0:
            st.info(f"No cards to display for the selected filters (Area: {area_filter}, Package: {package_filter})")
            return

        # Create filter info display
        filter_info = []
        if area_filter != "All Areas":
            filter_info.append(f"Area: {area_filter}")
        if package_filter != "All Packages":
            filter_info.append(f"Package: {package_filter}")
        
        if filter_info:
            st.markdown(f"**Filtered by:** {' | '.join(filter_info)}")
        
        area_groups = {}
        for card in cards_data:
            area = card['area']
            if area not in area_groups:
                area_groups[area] = []
            area_groups[area].append(card)
        
        for area, cards in area_groups.items():
            st.subheader(f"🏢 {area}")
            cols = st.columns(min(len(cards), 6))  # Using 6 columns for compact display
            for idx, card in enumerate(cards):
                with cols[idx % 6]:
                    st.markdown(f"""
                    <div style="
                        background: linear-gradient(135deg, {card['color']}, {card['color']}dd);
                        padding: 0.8rem;
                        border-radius: 8px;
                        color: white;
                        text-align: center;
                        margin: 0.3rem 0;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                        transition: transform 0.2s;
                        min-height: 80px;
                        "onmouseover="this.style.transform='scale(1.02)'" 
                        onmouseout="this.style.transform='scale(1)'">
                        <h5 style="margin: 0 0 0.3rem 0; color: white; font-weight: bold; font-size: 0.9rem;">{card['package']}</h5>
                        <div style="margin: 0.3rem 0; padding: 0.2rem 0.5rem; 
                                background: rgba(255,255,255,0.25); border-radius: 10px; 
                                font-size: 0.7rem; display: inline-block;">
                            {card['status']}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

def main():
    st.set_page_config(
        page_title = 'Operational Readiness Deliverables Tracker',
        layout = 'wide',
        initial_sidebar_state = 'expanded'
    )

    st.title("Operational Readiness FFEx MOC - Deliverables Adherence Monitor")
    st.markdown('---')

    with st.sidebar:
        st.header("Filters")
        st.info("Interactive filters section.")

        processor = ProjectDataProcessor("C:\\Users\\RNYC\\OneDrive - Novo Nordisk\\deliverablesTracker_dev.xlsx")
        processor.process_all()       

        timeline_view = st.selectbox(
            "Timeline drill",
            ["Monthly", "Weekly", "Yearly"],
            index=0
        )
        
        st.markdown("---")
        area_options = ['All Areas'] + sorted(list(processor.absolute['Areas'].unique()))
        selected_area_filter = st.selectbox(
            "Area Adherence Drill",
            options=area_options,
            index=0,
            key="area_adherence_filter"
        )
        
        if selected_area_filter == 'All Areas':
            package_options = ['All Packages'] + sorted(list(processor.absolute['Packages'].unique()))
        else:
            area_packages = processor.absolute[processor.absolute['Areas'] == selected_area_filter]['Packages'].unique()
            package_options = ['All Packages'] + sorted(list(area_packages))
        
        selected_package_filter = st.selectbox(
            "Package Adherence Drill",
            options=package_options,
            index=0,
            key="package_adherence_filter"
        )

    try:
        with st.spinner("Loading and processing data..."):
            dashboard = ProjectDashboard(processor)

            if 'show_packages' not in st.session_state:
                st.session_state.show_packages = False
            if 'selected_area' not in st.session_state:
                st.session_state.selected_area = 'All Areas'

            col1, col2 = st.columns([1, 2])

            with col1:
                if not st.session_state.show_packages:
                    areas_fig = dashboard.create_areas_pie_chart()
                    st.plotly_chart(areas_fig, use_container_width=True)
                    
                    if st.button("🔍 Show Package Distribution", type="primary", use_container_width=True):
                        st.session_state.show_packages = True
                        st.rerun()
                
                else:
                    st.subheader("📦 Package Distribution")
                    
                    area_options = ['All Areas'] + sorted(list(processor.absolute['Areas'].unique()))
                    selected_area = st.selectbox(
                        "Select Area:",
                        options=area_options,
                        index=area_options.index(st.session_state.selected_area) if st.session_state.selected_area in area_options else 0,
                        key="area_selector"
                    )
                    st.session_state.selected_area = selected_area
                    
                    packages_fig = dashboard.create_packages_pie_chart(selected_area)
                    st.plotly_chart(packages_fig, use_container_width=True)
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.button("⬅️ Back to Areas", use_container_width=True):
                            st.session_state.show_packages = False
                            st.rerun()
                    
                    with col_btn2:
                        if st.button("🔄 Refresh", use_container_width=True):
                            st.rerun()
                    
                    if selected_area != 'All Areas':
                        st.markdown("---")
                        st.subheader("📋 Package Details")
                        area_data = processor.absolute[processor.absolute['Areas'] == selected_area]
                        package_details = area_data[['Packages', 'Total']].sort_values('Total', ascending=False)
                        package_details.columns = ['Package', 'Deliverables']
                        package_details.index = range(1, len(package_details) + 1)
                        st.dataframe(package_details, use_container_width=True, height=200)

            with col2:
                st.subheader("S-Curve Progress")
                s_curve_fig = dashboard.create_s_curve(timeline_view, selected_area_filter, selected_package_filter)
                st.plotly_chart(s_curve_fig, use_container_width = True)

            st.markdown('---')
            st.header("Area & Package Status")
            cards_data = dashboard.create_status_cards(selected_area_filter, selected_package_filter)
            dashboard.display_cards_streamlit(cards_data, selected_area_filter, selected_package_filter)

            st.markdown('---')
            st.header("Delayed Deliverables Summary")
            delay_table = dashboard.create_delay_table()
            if not delay_table.empty:
                st.dataframe(
                    delay_table,
                    use_container_width = True,
                    height = 400
                )
            else:
                st.success("No delayed deliverables found!")

            st.markdown('---')
            st.header("Key Performance Indicators")

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                total_deliverables = processor.absolute['Total'].sum()
                st.metric("Total Deliverables", f"{total_deliverables}")
            
            with col2:
                completed_pct = len(processor.mgmtLevel[processor.mgmtLevel['status'] == 'Done'])*100 / len(processor.mgmtLevel)
                st.metric("Completion Rate", f"{completed_pct:.2f}%")

            with col3:
                delayed_count = len(delay_table) if not delay_table.empty else 0
                st.metric("Delayed Items", f"{delayed_count}")

            with col4:
                on_track_pct = len(processor.mgmtLevel[processor.mgmtLevel['status'].isin(['Done', 'On going'])])*100 / len(processor.mgmtLevel)
                st.metric("On track", f"{on_track_pct:.2f}%")
    except Exception as e:
        st.error(f"Error : {str(e)}")
        st.info("Please check your excel file and ensure all required sheets are present.")

if __name__ == "__main__":
    main()
