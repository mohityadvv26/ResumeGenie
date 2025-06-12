import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from config.database import get_database_connection
from plotly.subplots import make_subplots
from io import BytesIO
import traceback
import pymysql

class DashboardManager:
    def __init__(self):
        self.conn = get_database_connection()
        self.colors = {
            'primary': '#4CAF50',
            'secondary': '#2196F3',
            'warning': '#FFA726',
            'danger': '#F44336',
            'info': '#00BCD4',
            'success': '#66BB6A',
            'purple': '#9C27B0',
            'background': '#1E1E1E',
            'card': '#2D2D2D',
            'text': '#FFFFFF',
            'subtext': '#B0B0B0'
        }
        
    def apply_dashboard_style(self):
        """Apply custom styling for dashboard"""
        st.markdown("""
            <style>
                .dashboard-title {
                    font-size: 2.5rem;
                    font-weight: bold;
                    margin-bottom: 2rem;
                    color: white;
                    text-align: center;
                }
                
                .metric-card {
                    background-color: #2D2D2D;
                    border-radius: 15px;
                    padding: 1.5rem;
                    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                    transition: transform 0.3s ease;
                    height: 100%;
                }
                
                .metric-card:hover {
                    transform: translateY(-5px);
                }
                
                .metric-value {
                    font-size: 2.5rem;
                    font-weight: bold;
                    color: #4CAF50;
                    margin: 0.5rem 0;
                }
                
                .metric-label {
                    font-size: 1rem;
                    color: #B0B0B0;
                }
                
                .trend-up {
                    color: #4CAF50;
                    font-size: 1.2rem;
                }
                
                .trend-down {
                    color: #F44336;
                    font-size: 1.2rem;
                }
                
                .chart-container {
                    background-color: #2D2D2D;
                    border-radius: 15px;
                    padding: 1.5rem;
                    margin: 1rem 0;
                }
                
                .section-title {
                    font-size: 1.5rem;
                    color: white;
                    margin: 2rem 0 1rem 0;
                }
                
                .stPlotlyChart {
                    background-color: #2D2D2D;
                    border-radius: 15px;
                    padding: 1rem;
                }
                
                div[data-testid="stHorizontalBlock"] > div {
                    background-color: #2D2D2D;
                    border-radius: 15px;
                    padding: 1rem;
                    margin: 0.5rem;
                }

                [data-testid="stMetricValue"] {
                    font-size: 2rem !important;
                }

                [data-testid="stMetricLabel"] {
                    font-size: 1rem !important;
                }
            </style>
        """, unsafe_allow_html=True)

    def get_resume_metrics(self):
        """Get resume-related metrics from database"""
        cursor = self.conn.cursor()
        
        # Get current date
        now = datetime.now()
        start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0)
        start_of_week = now - timedelta(days=now.weekday())
        start_of_month = now.replace(day=1)
        
        # Fetch metrics for different time periods
        metrics = {}
        for period, start_date in [
            ('Today', start_of_day),
            ('This Week', start_of_week),
            ('This Month', start_of_month),
            ('All Time', datetime(2000, 1, 1))
        ]:
            cursor.execute("""
                SELECT 
                    COUNT(DISTINCT rd.id) as total_resumes,
                    ROUND(AVG(ra.ats_score), 1) as avg_ats_score,
                    ROUND(AVG(ra.keyword_match_score), 1) as avg_keyword_score,
                    COUNT(DISTINCT CASE WHEN ra.ats_score >= 70 THEN rd.id END) as high_scoring
                FROM resume_data rd
                LEFT JOIN resume_analysis ra ON rd.id = ra.resume_id
                WHERE rd.created_at >= ?
            """, (start_date.strftime('%Y-%m-%d %H:%M:%S'),))
            
            row = cursor.fetchone()
            if row:
                metrics[period] = {
                    'total': row[0] or 0,
                    'ats_score': row[1] or 0,
                    'keyword_score': row[2] or 0,
                    'high_scoring': row[3] or 0
                }
            else:
                metrics[period] = {
                    'total': 0,
                    'ats_score': 0,
                    'keyword_score': 0,
                    'high_scoring': 0
                }
        
        return metrics

    def get_skill_distribution(self):
        """Get skill distribution data"""
        cursor = self.conn.cursor()
        try:
            cursor.execute("""
                WITH RECURSIVE split(skill, rest) AS (
                    SELECT CAST('' AS CHAR(1000)), CONCAT(skills, ',')
                    FROM resume_data
                    UNION ALL
                    SELECT
                        CAST(SUBSTRING_INDEX(rest, ',', 1) AS CHAR(1000)),
                        SUBSTRING(rest FROM CHAR_LENGTH(SUBSTRING_INDEX(rest, ',', 1)) + 2)
                    FROM split
                    WHERE rest <> ''
                ),
                SkillCategories AS (
                    SELECT 
                        CASE 
                            WHEN LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%python%' OR LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%java%' OR 
                                LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%javascript%' OR LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%c++%' OR 
                                LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%programming%' THEN 'Programming'
                            WHEN LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%sql%' OR LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%database%' OR 
                                LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%mongodb%' THEN 'Database'
                            WHEN LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%aws%' OR LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%cloud%' OR 
                                LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%azure%' THEN 'Cloud'
                            WHEN LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%agile%' OR LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%scrum%' OR 
                                LOWER(TRIM(BOTH '[]"' FROM skill)) LIKE '%management%' THEN 'Management'
                            ELSE 'Other'
                        END as category,
                        COUNT(*) as count
                    FROM split
                    WHERE skill <> ''
                    GROUP BY category
                )
                SELECT category, count
                FROM SkillCategories
                ORDER BY count DESC
            """)
            
            categories, counts = [], []
            for row in cursor.fetchall():
                categories.append(row['category'])
                counts.append(row['count'])
                
            return categories, counts
        except Exception as e:
            print(f"Error in get_skill_distribution:")
            traceback.print_exc()
            return [], []
    
    def get_weekly_trends(self):
        now = datetime.now()
        dates = [(now - timedelta(days=x)).strftime('%Y-%m-%d') for x in range(6, -1, -1)]

        submissions = []
        cursor = self.conn.cursor()

        for date in dates:
            cursor.execute("""
                SELECT COUNT(*) FROM resume_data
                WHERE DATE(created_at) = DATE(%s)
            """, (date,))
            result = cursor.fetchone()
            if result is not None and isinstance(result, (list, tuple)):
                submissions.append(result[0])
            else:
                submissions.append(0)

        # Convert full dates to weekday abbreviations (e.g., 'Mon', 'Tue')
        weekdays = [datetime.strptime(d, "%Y-%m-%d").strftime("%a") for d in dates]
        return weekdays, submissions

    def get_job_category_stats(self):
        """Get statistics by job category"""
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT 
                COALESCE(target_category, 'Other') as category,
                COUNT(*) as count,
                ROUND(AVG(CASE WHEN ra.ats_score >= 70 THEN 1 ELSE 0 END) * 100, 1) as success_rate
            FROM resume_data rd
            LEFT JOIN resume_analysis ra ON rd.id = ra.resume_id
            GROUP BY category
            ORDER BY count DESC
            LIMIT 5
        """)
        
        categories, success_rates = [], []
        for row in cursor.fetchall():
            categories.append(row["category"])
            success_rates.append(row["success_rate"] or 0)
            
        return categories, success_rates

    def render_admin_panel(self):
        """Render admin panel with data management tools"""
        st.sidebar.markdown("### 👋 Welcome Admin!")
        st.sidebar.markdown("---")
        
        if st.sidebar.button("🚪 Logout"):
            st.session_state.is_admin = False
            st.rerun()
            
        st.sidebar.markdown("### 🛠️ Admin Tools")
        
        # Data Export Options
        export_format = st.sidebar.selectbox(
            "Export Format",
            ["Excel", "CSV", "JSON"],
            key="export_format"
        )
        
        if st.sidebar.button("📥 Export Data"):
            if export_format == "Excel":
                excel_data = self.export_to_excel()
                if excel_data:
                    st.sidebar.download_button(
                        "⬇️ Download Excel",
                        data=excel_data,
                        file_name=f"resume_data_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            elif export_format == "CSV":
                csv_data = self.export_to_csv()
                if csv_data:
                    st.sidebar.download_button(
                        "⬇️ Download CSV",
                        data=csv_data,
                        file_name=f"resume_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv"
                    )
            else:
                json_data = self.export_to_json()
                if json_data:
                    st.sidebar.download_button(
                        "⬇️ Download JSON",
                        data=json_data,
                        file_name=f"resume_data_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                        mime="application/json"
                    )

        # Database Stats
        st.sidebar.markdown("### 📊 Database Stats")
        stats = self.get_database_stats()
        st.sidebar.markdown(f"""
            - Total Resumes: {stats['total_resumes']}
            - Today's Submissions: {stats['today_submissions']}
            - Storage Used: {stats['storage_size']}
        """)

    def get_resume_data(self):
        """Get all resume data"""
        cursor = self.conn.cursor(pymysql.cursors.DictCursor)  # ✅ use DictCursor
        try:
            cursor.execute('''
            SELECT 
                r.email,
                r.phone,
                r.target_role,
                r.target_category,
                r.created_at,
                a.ats_score AS ATS_Score,
                a.keyword_match_score AS Keyword_Match,
                a.format_score AS Format_Score,
                a.section_score AS Section_Score
            FROM resume_data r
            LEFT JOIN resume_analysis a ON r.id = a.resume_id
            ORDER BY r.created_at DESC
            ''')
            return cursor.fetchall()
        except Exception as e:
            print(f"Error fetching resume data: {str(e)}")
            return []

    def render_resume_data_section(self):
        """Render resume data section with Excel download"""
        st.markdown("<h2 class='section-title'>Resume Submissions</h2>", unsafe_allow_html=True)
        
        # Get resume data
        resume_data = self.get_resume_data()
        
        if resume_data:
            # Convert to DataFrame
            columns = [
                'Email', 'Phone', 'Target Role', 'Target Category', 'Submission Date',
                'ATS Score', 'Keyword Match', 'Format Score', 'Section Score'
            ]

            df = pd.DataFrame(resume_data)
            # Rename columns safely
            df.rename(columns={
                'ats_score': 'ATS Score',
                'keyword_match_score': 'Keyword Match',
                'format_score': 'Format Score',
                'section_score': 'Section Score',
            }, inplace=True)

            # Format scores as percentages
            score_columns = ['ATS_Score', 'Keyword_Match', 'Format_Score', 'Section_Score']
            for col in score_columns:
                df[col] = df[col].apply(lambda x: f"{x:.1f}" if pd.notnull(x) else "N/A")
            
            # Style the dataframe
            st.markdown("""
            <style>
            .resume-data {
                background-color: #2D2D2D;
                border-radius: 10px;
                padding: 1rem;
                margin-bottom: 1rem;
            }
            </style>
            """, unsafe_allow_html=True)
            
            with st.container():
                st.markdown('<div class="resume-data">', unsafe_allow_html=True)
                
                # Add filters
                col1, col2 = st.columns(2)
                with col1:
                    target_role = st.selectbox(
                        "Filter by Target Role",
                        options=["All"] + list(df['target_role'].unique()),
                        key="role_filter"
                    )
                with col2:
                    target_category = st.selectbox(
                        "Filter by Category",
                        options=["All"] + list(df['target_category'].unique()),
                        key="category_filter"
                    )
                
                # Apply filters
                filtered_df = df.copy()
                if target_role != "All":
                    filtered_df = filtered_df[filtered_df['target_role'] == target_role]
                if target_category != "All":
                    filtered_df = filtered_df[filtered_df['target_category'] == target_category]
                
                # Display filtered data
                st.dataframe(
                    filtered_df.drop(columns=['name'], errors='ignore'),
                    use_container_width=True,
                    hide_index=True
                )
                
                # Add download buttons
                col1, col2 = st.columns(2)
                with col1:
                    # Download filtered data
                    excel_buffer = BytesIO()
                    filtered_df.to_excel(excel_buffer, index=False, engine='openpyxl')
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Filtered Data",
                        data=excel_buffer,
                        file_name=f"resume_data_filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_filtered_data"
                    )
                
                with col2:
                    # Download all data
                    excel_buffer_all = BytesIO()
                    df.to_excel(excel_buffer_all, index=False, engine='openpyxl')
                    excel_buffer_all.seek(0)
                    
                    st.download_button(
                        label="📥 Download All Data",
                        data=excel_buffer_all,
                        file_name=f"resume_data_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_all_data"
                    )
                
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No resume submissions available")

    def render_admin_section(self):
        """Render admin section with logs and Excel download"""
        # Render resume data section
        self.render_resume_data_section()
        
    def export_to_excel(self):
        """Export data to Excel format"""
        query = """
            SELECT 
                rd.email, rd.phone,
                rd.summary, rd.target_role, rd.target_category,
                rd.education, rd.experience, rd.projects, rd.skills,
                ra.ats_score, ra.keyword_match_score, ra.format_score, ra.section_score,
                ra.missing_skills, ra.recommendations,
                rd.created_at
            FROM resume_data rd
            LEFT JOIN resume_analysis ra ON rd.id = ra.resume_id
        """
        try:
            df = pd.read_sql_query(query, self.conn)
            
            # Create Excel writer object
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Resume Data', index=False)
                
                # Get the workbook and the worksheet
                workbook = writer.book
                worksheet = writer.sheets['Resume Data']
                
                # Add formatting
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Write headers with formatting
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    
                # Auto-adjust columns' width
                for i, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    ) + 2
                    worksheet.set_column(i, i, min(max_length, 50))
            
            # Return the Excel file
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            st.error(f"Error exporting to Excel: {str(e)}")
            return None

    def export_to_csv(self):
        """Export data to CSV format"""
        query = """
            SELECT 
                rd.name, rd.email, rd.phone,
                rd.summary, rd.target_role, rd.target_category,
                rd.education, rd.experience, rd.projects, rd.skills,
                ra.ats_score, ra.keyword_match_score, ra.format_score, ra.section_score,
                ra.missing_skills, ra.recommendations,
                rd.created_at
            FROM resume_data rd
            LEFT JOIN resume_analysis ra ON rd.id = ra.resume_id
        """
        try:
            df = pd.read_sql_query(query, self.conn)
            return df.to_csv(index=False).encode('utf-8')
        except Exception as e:
            st.error(f"Error exporting to CSV: {str(e)}")
            return None

    def export_to_json(self):
        """Export data to JSON format"""
        query = """
            SELECT 
                rd.*, ra.*
            FROM resume_data rd
            LEFT JOIN resume_analysis ra ON rd.id = ra.resume_id
        """
        try:
            df = pd.read_sql_query(query, self.conn)
            return df.to_json(orient='records', date_format='iso')
        except Exception as e:
            st.error(f"Error exporting to JSON: {str(e)}")
            return None

    def get_database_stats(self):
        """Get database statistics"""
        cursor = self.conn.cursor()
        stats = {}
        
        # Total resumes
        cursor.execute("SELECT COUNT(*) FROM resume_data")
        stats['total_resumes'] = cursor.fetchone()[0]
        
        # Today's submissions
        cursor.execute("""
            SELECT COUNT(*) 
            FROM resume_data 
            WHERE DATE(created_at) = DATE('now')
        """)
        stats['today_submissions'] = cursor.fetchone()[0]
        
        # Database size (approximate)
        cursor.execute("PRAGMA page_count")
        page_count = cursor.fetchone()[0]
        cursor.execute("PRAGMA page_size")
        page_size = cursor.fetchone()[0]
        size_bytes = page_count * page_size
        
        if size_bytes < 1024:
            stats['storage_size'] = f"{size_bytes} bytes"
        elif size_bytes < 1024 * 1024:
            stats['storage_size'] = f"{size_bytes/1024:.1f} KB"
        else:
            stats['storage_size'] = f"{size_bytes/(1024*1024):.1f} MB"
        
        return stats

    def get_admin_logs(self):
        """Get admin logs"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('''
            SELECT admin_email, action, timestamp
            FROM admin_logs
            ORDER BY timestamp DESC
            ''')
            # return cursor.fetchall()
            return cursor.fetchall()
            
        except Exception as e:
            print(f"Error fetching admin logs: {str(e)}")
            return []

    def render_dashboard(self):
        """Main dashboard rendering function"""
        # Apply styling
        st.markdown("""
            <style>
                .dashboard-container {
                    background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
                    padding: 2rem;
                    border-radius: 20px;
                    margin: -1rem -1rem 2rem -1rem;
                    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
                }
                .dashboard-title {
                    color: #4FD1C5;
                    font-size: 2.5rem;
                    margin-bottom: 0.5rem;
                    display: flex;
                    align-items: center;
                    gap: 1rem;
                }
                .dashboard-icon {
                    background: rgba(79, 209, 197, 0.2);
                    padding: 0.5rem;
                    border-radius: 12px;
                }
                .stats-grid {
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 1.5rem;
                    margin-top: 2rem;
                }
                .stat-card {
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(10px);
                    padding: 1.5rem;
                    border-radius: 16px;
                    border: 1px solid rgba(255, 255, 255, 0.1);
                    transition: all 0.3s ease;
                }
                .stat-card:hover {
                    transform: translateY(-5px);
                    background: rgba(255, 255, 255, 0.1);
                }
                .stat-value {
                    font-size: 2.5rem;
                    font-weight: bold;
                    margin: 0;
                    color: #4FD1C5;
                }
                .stat-label {
                    font-size: 1rem;
                    color: rgba(255, 255, 255, 0.7);
                    margin: 0.5rem 0 0 0;
                }
                .section-title {
                    color: #4FD1C5;
                    font-size: 1.5rem;
                    margin: 1rem 0 0.5rem 0;
                    padding-bottom: 0.5rem;
                    border-bottom: 2px solid rgba(79, 209, 197, 0.2);
                }
                .chart-container {
                    background: rgba(255, 255, 255, 0.05);
                    border-radius: 16px;
                    padding: 1rem;
                    margin-bottom: 1rem;
                }
                .insights-grid {
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 1.5rem;
                    margin-top: 1rem;
                }
                .insight-card {
                    background: rgba(255, 255, 255, 0.05);
                    padding: 1.5rem;
                    border-radius: 16px;
                    border: 1px solid rgba(255, 255, 255, 0.1);
                }
                .trend-indicator {
                    display: inline-flex;
                    align-items: center;
                    padding: 0.25rem 0.5rem;
                    border-radius: 12px;
                    font-size: 0.875rem;
                    margin-left: 0.5rem;
                }
                .trend-up {
                    background: rgba(46, 204, 113, 0.2);
                    color: #2ecc71;
                }
                .trend-down {
                    background: rgba(231, 76, 60, 0.2);
                    color: #e74c3c;
                }
                @keyframes fadeInUp {
                    from {
                        opacity: 0;
                        transform: translateY(20px);
                    }
                    to {
                        opacity: 1;
                        transform: translateY(0);
                    }
                }
                .animate-fade-in {
                    animation: fadeInUp 0.5s ease-out forwards;
                }
            </style>
        """, unsafe_allow_html=True)

        # Dashboard Header
        st.markdown("""
            <div class="dashboard-container animate-fade-in">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div class="dashboard-title">
                        <span class="dashboard-icon">📊</span>
                        Resume Analytics Dashboard
                    </div>
                    <div style="color: rgba(255, 255, 255, 0.7);">
                        Last updated: {}
                    </div>
                </div>
            """.format(datetime.now().strftime('%B %d, %Y %I:%M %p')), unsafe_allow_html=True)

        
        # Key Insights Section
        st.markdown('<div class="section-title">🎯 Key Insights</div>', unsafe_allow_html=True)
        # insights = self.get_detailed_insights()
        
        # Admin logs section with Excel download functionality
        if st.session_state.get('is_admin', False):
            self.render_admin_section()

    
    # def get_detailed_insights(self):
    #     """Get detailed insights from the database"""
    #     cursor = self.conn.cursor()
    #     insights = []
        
        
    #     # Most Common Skills
    #     cursor.execute("""
    #         WITH RECURSIVE
    #         split(skill, rest) AS (
    #             SELECT '', CONCAT(skills, ',') 
    #             FROM resume_data 
    #             WHERE skills IS NOT NULL
    #             UNION ALL
    #             SELECT
    #                 substr(rest, 0, instr(rest, ',')),
    #                 substr(rest, instr(rest, ',') + 1)
    #             FROM split 
    #             WHERE rest <> ''
    #         ),
    #         cleaned_skills AS (
    #             SELECT TRIM(REPLACE(REPLACE(skill, '[', ''), ']', '')) as skill
    #             FROM split 
    #             WHERE skill <> ''
    #         )
    #         SELECT skill, COUNT(*) as count
    #         FROM cleaned_skills
    #         GROUP BY skill
    #         ORDER BY count DESC
    #         LIMIT 3
    #     """)
    #     top_skills = cursor.fetchall()
    #     if top_skills:
    #         skills_text = f"Most in-demand skills: Python ({top_skills[0][1]} resumes), Java ({top_skills[1][1]} resumes), Express ({top_skills[2][1]} resumes)"
    #         insights.append({
    #             'title': 'Top Skills',
    #             'icon': '💡',
    #             'description': f"Most in-demand skills: {skills_text}",
    #             'trend_class': 'trend-up',
    #             'trend_icon': '🔝',
    #             'trend_value': f"Top {len(top_skills)}"
    #         })
        
    #     return insights


   

# vdvdfvdsdkjnawdckjsndlknmadfmv
# '\dvadsfvmadsk;lsnaeklv
