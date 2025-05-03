import streamlit as st
import plotly.express as px
import pandas as pd
import openpyxl
import io
import os
from pathlib import Path
import matplotlib.pyplot as plt
from fpdf import FPDF
import base64
from datetime import datetime
import re
import sys
from streamlit.web import cli as stcli
import sqlite3
import shutil
from datetime import datetime, timedelta
import sqlite3
import shutil

# ============================================
# UI CONFIGURATION
# ============================================

# Color scheme
bg_color = "#1E1E1E"  # Dark background
entry_bg = "#2D2D2D"  # Element background
text_color = "#FFFFFF"  # White text
button_color = "#3A3A3A"  # Buttons
accent_color = "#FF9500"  # Orange accent
accent_light = "#FFB347"  # Light orange
success_color = "#4CAF50"  # Green for success
warning_color = "#FFC107"  # Yellow for warnings
error_color = "#F44336"  # Red for errors

st.markdown(f"""
    <style>
        /* Base styles */
        .stApp {{
            background-color: {bg_color};
            color: {text_color};
            line-height: 1.6;
        }}

        /* Containers and cards */
        .stContainer, .stForm {{
            background-color: {entry_bg};
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            border-left: 4px solid {accent_color};
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }}

        /* Headers */
        h1 {{
            color: {accent_color};
            font-size: 2rem;
            border-bottom: 2px solid {accent_color};
            padding-bottom: 10px;
        }}

        h2 {{
            color: {accent_color};
            font-size: 1.5rem;
            margin-top: 1.5rem;
        }}

        h3 {{
            color: {text_color};
            font-size: 1.2rem;
        }}

        /* Input fields */
        .stTextInput>label, .stNumberInput>label, 
        .stSelectbox>label, .stSlider>label,
        .stDateInput>label, .stFileUploader>label {{
            color: {text_color} !important;
            font-weight: 500;
        }}

        .stTextInput>div>div>input, .stNumberInput>div>div>input,
        .stSelectbox>div>div>div {{
            background-color: {entry_bg};
            color: {text_color} !important;
            border: 1px solid #555;
            border-radius: 6px;
        }}

        /* Buttons */
        .stButton>button {{
            background-color: {button_color};
            color: {text_color} !important;
            border-radius: 8px;
            border: none;
            padding: 10px 20px;
            font-weight: 500;
            transition: all 0.3s ease;
        }}

        .stButton>button:hover {{
            background-color: {accent_color};
            transform: scale(1.05);
        }}

        /* Metrics and results */
        [data-testid="stMetric"] {{
            background-color: {entry_bg};
            border: 1px solid {accent_color};
            border-radius: 10px;
            padding: 15px;
        }}

        [data-testid="stMetricLabel"] {{
            color: {accent_light} !important;
            font-size: 1.1rem;
            font-weight: 600;
        }}

        [data-testid="stMetricValue"] {{
            color: {text_color} !important;
            font-size: 1.8rem;
            font-weight: 700;
        }}

        [data-testid="stMetricDelta"] div {{
            color: {accent_light} !important;
            font-size: 1rem;
        }}

        /* Tables */
        .stDataFrame {{
            background-color: {entry_bg} !important;
            border: 1px solid {accent_color} !important;
        }}

        .stDataFrame th {{
            background-color: {accent_color} !important;
            color: {bg_color} !important;
            font-weight: 700;
        }}

        /* Notifications */
        .stAlert {{
            background-color: {entry_bg} !important;
            border-left: 4px solid {accent_color};
        }}

        .stSuccess {{
            border-left-color: {success_color} !important;
        }}

        .stWarning {{
            border-left-color: {warning_color} !important;
        }}

        .stError {{
            border-left-color: {error_color} !important;
        }}

        /* Tabs */
        .stTabs [data-baseweb="tab"] {{
            background: {button_color};
            color: {text_color} !important;
            border-radius: 8px 8px 0 0;
        }}

        .stTabs [aria-selected="true"] {{
            background: {accent_color} !important;
            color: {bg_color} !important;
        }}

        /* Responsiveness */
        @media (max-width: 768px) {{
            .stForm {{ padding: 15px; }}
            [data-testid="stMetricValue"] {{ font-size: 1.4rem; }}
        }}
    </style>
""", unsafe_allow_html=True)

# ============================================
# DATA CONFIGURATION
# ============================================

MATERIALS_FILE = "materials_data.csv"
HISTORY_DB = "repair_history.db"

def init_database():
    """Initialize the database"""
    conn = sqlite3.connect(HISTORY_DB)
    c = conn.cursor()

    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='repairs'")
    table_exists = c.fetchone()

    if not table_exists:
        c.execute('''
            CREATE TABLE repairs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                job_order TEXT,
                brand TEXT,
                machine_model TEXT,
                machine_sn TEXT,
                prefix TEXT,
                client TEXT,
                department TEXT,
                cylinder_name TEXT,
                cylinder_sn TEXT,
                pipe_num TEXT,
                rod_num TEXT,
                head_num TEXT,
                materials_cost REAL,
                spare_parts_cost REAL,
                labor_cost REAL,
                material_margin REAL,
                labor_margin REAL,
                total_cost_kzt REAL,
                total_cost_usd REAL
            )
        ''')
        conn.commit()
        conn.close()
        return True

    c.execute("PRAGMA table_info(repairs)")
    columns = [col[1] for col in c.fetchall()]

    expected_columns = {
        'id', 'date', 'job_order', 'brand', 'machine_model',
        'machine_sn', 'prefix', 'client', 'department', 'cylinder_name',
        'cylinder_sn', 'pipe_num', 'rod_num', 'head_num', 'materials_cost',
        'spare_parts_cost', 'labor_cost', 'material_margin', 'labor_margin',
        'total_cost_kzt', 'total_cost_usd'
    }

    actual_columns = set(columns)

    if actual_columns != expected_columns:
        missing = expected_columns - actual_columns
        extra = actual_columns - expected_columns

        error_msg = "Table structure mismatch.\n"
        if missing:
            error_msg += f"Missing columns: {', '.join(missing)}\n"
        if extra:
            error_msg += f"Extra columns: {', '.join(extra)}\n"

        st.error(error_msg)
        conn.close()

        if st.button("Fix table structure", key="fix_db"):
            fix_database_structure()
            st.experimental_rerun()

        return False

    conn.close()
    return True

def fix_database_structure():
    """Fix the repairs table structure"""
    try:
        conn = sqlite3.connect(HISTORY_DB)
        c = conn.cursor()

        c.execute('''
            CREATE TABLE repairs_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                job_order TEXT,
                brand TEXT,
                machine_model TEXT,
                machine_sn TEXT,
                prefix TEXT,
                client TEXT,
                department TEXT,
                cylinder_name TEXT,
                cylinder_sn TEXT,
                pipe_num TEXT,
                rod_num TEXT,
                head_num TEXT,
                materials_cost REAL,
                spare_parts_cost REAL,
                labor_cost REAL,
                material_margin REAL,
                labor_margin REAL,
                total_cost_kzt REAL,
                total_cost_usd REAL
            )
        ''')

        c.execute("PRAGMA table_info(repairs)")
        old_columns = [col[1] for col in c.fetchall()]

        common_columns = set(old_columns) & {
            'date', 'job_order', 'brand', 'machine_model', 'machine_sn',
            'prefix', 'client', 'department', 'cylinder_name', 'cylinder_sn',
            'pipe_num', 'rod_num', 'head_num', 'materials_cost', 'spare_parts_cost',
            'labor_cost', 'material_margin', 'labor_margin', 'total_cost_kzt', 'total_cost_usd'
        }

        columns_to_transfer = ', '.join(common_columns)
        c.execute(f'''
            INSERT INTO repairs_new ({columns_to_transfer})
            SELECT {columns_to_transfer} FROM repairs
        ''')

        c.execute("DROP TABLE repairs")
        c.execute("ALTER TABLE repairs_new RENAME TO repairs")

        conn.commit()
        st.success("Table structure fixed successfully!")
        return True
    except Exception as e:
        st.error(f"Error fixing structure: {str(e)}")
        return False
    finally:
        conn.close()

def load_default_materials():
    default_data = pd.DataFrame({
        "MaterialName": [
            "Труба E355 40x50",
            "Шток 42CrMo4 Ø20",
            "Кругляк 45 Ø40",
            "Кругляк 45 Ø50",
            "Кругляк 45 Ø60",
            "Кругляк 45 Ø70",
            "Кругляк 45 Ø80",
            "Кругляк 45 Ø90",
            "Кругляк 45 Ø100",
            "Голова стальная",
            "Голова чугунная"
        ],
        "Cost": [5400.0, 9200.0, 2000.0, 2500.0, 3000.0, 3500.0, 4000.0, 4500.0, 5000.0, 1500.0, 1200.0],
        "Units": ["м", "м", "м", "м", "м", "м", "м", "м", "м", "кг", "кг"],
        "LastUpdated": [datetime.now().date()] * 11
    })
    default_data.to_csv(MATERIALS_FILE, index=False)
    return default_data

def load_materials():
    if Path(MATERIALS_FILE).exists():
        return pd.read_csv(MATERIALS_FILE)
    return load_default_materials()

def save_repair_to_db(record):
    """Save repair record to database"""
    try:
        conn = sqlite3.connect(HISTORY_DB)
        c = conn.cursor()

        c.execute('''
            INSERT INTO repairs (
                date, job_order, brand, machine_model, machine_sn, 
                prefix, client, department, cylinder_name, cylinder_sn, 
                pipe_num, rod_num, head_num, materials_cost, spare_parts_cost, 
                labor_cost, material_margin, labor_margin, total_cost_kzt, total_cost_usd
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', record)
        conn.commit()
        return True
    except Exception as e:
        st.error(f"Database save error: {str(e)}")
        return False
    finally:
        conn.close()

def load_repairs_from_db():
    """Load repair history from database"""
    try:
        conn = sqlite3.connect(HISTORY_DB)
        df = pd.read_sql('SELECT * FROM repairs ORDER BY date DESC', conn)
        return df
    except Exception as e:
        st.error(f"History load error: {str(e)}")
        return pd.DataFrame()
    finally:
        conn.close()

def create_backup():
    try:
        if not Path(HISTORY_DB).exists():
            st.warning("Database file not found")
            return None

        backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M')}.db"
        shutil.copy(HISTORY_DB, backup_name)
        return backup_name
    except Exception as e:
        st.error(f"Backup creation error: {str(e)}")
        return None

# Initialize database
if not init_database():
    st.error("Database initialization error. App may not work properly.")

# Load data
if 'price_data' not in st.session_state:
    st.session_state.price_data = load_materials()
if 'last_upload' not in st.session_state:
    st.session_state.last_upload = None

def extract_diameter(material_name):
    match = re.search(r'Ø(\d+)', material_name)
    return int(match.group(1)) if match else None

# ============================================
# MAIN CALCULATOR
# ============================================

def main_calculator():
    st.title("🔧 Hydraulic Cylinder Repair Calculator")

    with st.container():
        with st.form("repair_form"):
            st.subheader("📝 Basic Repair Data")
            col1, col2 = st.columns(2)

            with col1:
                job_order = st.text_input("Work Order", key="job_order")
                brand = st.text_input("Brand", key="brand")
                model = st.text_input("Machine Model", key="model")
                sn_machine = st.text_input("Machine S/N", key="sn_machine")
                prefix = st.text_input("Prefix", key="prefix")

            with col2:
                client = st.text_input("Client", key="client")
                department = st.text_input("Department", key="department")
                cylinder_name = st.text_input("Cylinder Name", key="cylinder_name")
                sn_cylinder = st.text_input("Cylinder S/N", key="sn_cylinder")
                pipe_num = st.text_input("Pipe #", key="pipe_num")
                rod_num = st.text_input("Rod #", key="rod_num")
                head_num = st.text_input("Head #", key="head_num")

            st.subheader("📦 Материалы и компоненты")
            col3, col4, col5, col6 = st.columns(4)

            with col3:
                tube_options = st.session_state.price_data[
                    st.session_state.price_data['MaterialName'].str.contains('Труба|труба', case=False)
                ]
                if not tube_options.empty:
                    tube_selection = st.selectbox("Труба", tube_options['MaterialName'])
                    tube_data = tube_options[tube_options['MaterialName'] == tube_selection].iloc[0]
                    tube_length = st.number_input("Длина трубы (м)", min_value=0.1, value=1.0, step=0.1)
                else:
                    st.warning("Нет данных о трубах")

            with col4:
                rod_options = st.session_state.price_data[
                    st.session_state.price_data['MaterialName'].str.contains('Шток|шток', case=False)
                ]
                if not rod_options.empty:
                    rod_selection = st.selectbox("Шток", rod_options['MaterialName'])
                    rod_data = rod_options[rod_options['MaterialName'] == rod_selection].iloc[0]
                    rod_length = st.number_input("Длина штока (м)", min_value=0.1, value=1.0, step=0.1)
                else:
                    st.warning("Нет данных о штоках")

            with col5:
                piston_options = st.session_state.price_data[
                    st.session_state.price_data['MaterialName'].str.contains('Кругляк|кругляк', case=False)
                ]
                if not piston_options.empty:
                    piston_options = piston_options.copy()
                    piston_options['Diameter'] = piston_options['MaterialName'].apply(extract_diameter)
                    piston_options = piston_options.sort_values('Diameter')

                    piston_selection = st.selectbox("Материал поршня", piston_options['MaterialName'])
                    piston_data = piston_options[piston_options['MaterialName'] == piston_selection].iloc[0]
                    piston_diameter = extract_diameter(piston_selection)

                    st.markdown(f"**Диаметр поршня:** {piston_diameter} мм")
                    piston_length = st.number_input("Длина поршня (мм)", min_value=10, value=100, step=5)
                else:
                    st.warning("Нет данных о материалах для поршней")

            with col6:
                head_options = st.session_state.price_data
                if not head_options.empty:
                    head_selection = st.selectbox("Материал головы", head_options['MaterialName'])
                    head_data = head_options[head_options['MaterialName'] == head_selection].iloc[0]

                    if head_data['Units'] == 'кг':
                        head_weight = st.number_input(f"Вес головы ({head_data['Units']})", min_value=0.1, value=5.0, step=0.1)
                    else:
                        head_weight = st.number_input(f"Размер головы ({head_data['Units']})", min_value=10, value=100, step=5)
                else:
                    st.warning("Нет данных о материалах")

            st.subheader("💰 Spare Parts Cost")
            spare_parts_cost = st.number_input("Spare parts cost (KZT)", min_value=0, value=0, step=1000)

            st.subheader("🛠 Labor")
            col7, col8, col9 = st.columns(3)
            with col7:
                hours_inspection = st.number_input("Inspection/disassembly (hrs)", min_value=0.5, value=2.0, step=0.5)
                hours_assembly = st.number_input("Assembly/shipping (hrs)", min_value=0.5, value=1.5, step=0.5)
            with col8:
                hours_liner = st.number_input("Liner manufacturing (hrs)", min_value=0.5, value=4.0, step=0.5)
                hours_rod = st.number_input("Rod manufacturing (hrs)", min_value=0.5, value=3.0, step=0.5)
            with col9:
                hours_piston = st.number_input("Piston manufacturing (hrs)", min_value=0.5, value=3.5, step=0.5)
                hours_head = st.number_input("Head manufacturing (hrs)", min_value=0.5, value=2.5, step=0.5)

            st.subheader("💰 Financial Parameters")
            col10, col11 = st.columns(2)
            with col10:
                usd_rate = st.number_input("USD/KZT rate", min_value=1, value=450)
                vat = st.number_input("VAT (%)", min_value=0, max_value=20, value=12)
            with col11:
                workshop_rate = st.number_input("Workshop rate (KZT/hour)", min_value=1000, value=5000)
                material_margin = st.slider("Material margin (%)", min_value=0, max_value=100, value=25)
                labor_margin = st.slider("Labor margin (%)", min_value=0, max_value=100, value=25)

            submitted = st.form_submit_button("Calculate Cost", type="primary")

    if submitted:
        if 'tube_data' not in locals() or 'rod_data' not in locals() or 'piston_data' not in locals() or 'head_data' not in locals():
            st.error("Please load material data")
            return

        # Cost calculation
        tube_cost_kzt = tube_data['Cost'] * tube_length
        rod_cost_kzt = rod_data['Cost'] * rod_length
        piston_cost_kzt = piston_data['Cost'] * (piston_length / 1000)
        head_cost_kzt = head_data['Cost'] * head_weight
        materials_cost_kzt = (tube_cost_kzt + rod_cost_kzt + piston_cost_kzt + head_cost_kzt) * (1 + material_margin / 100)
        labor_hours = hours_inspection + hours_liner + hours_rod + hours_piston + hours_head + hours_assembly
        labor_cost_kzt = labor_hours * workshop_rate * (1 + labor_margin / 100)
        subtotal_kzt = materials_cost_kzt + labor_cost_kzt + spare_parts_cost
        final_price_kzt = subtotal_kzt * (1 + vat / 100)
        final_price_usd = final_price_kzt / usd_rate

        # Prepare record for saving
        record = (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # date
            job_order, brand, model, sn_machine, prefix, client,  # 6 fields
            department, cylinder_name, sn_cylinder, pipe_num, rod_num, head_num,  # 6 fields
            float(materials_cost_kzt), float(spare_parts_cost), float(labor_cost_kzt),  # 3 fields
            float(material_margin), float(labor_margin),  # 2 fields
            float(final_price_kzt), float(final_price_usd)  # 2 fields
        )

        # Save to database
        if save_repair_to_db(record):
            st.success("Data saved to history!")
        else:
            st.error("Error saving to history")

        # Display results
        st.header("📊 Calculation Results")

        with st.container():
            st.subheader("📦 Materials (with margin)")
            cols = st.columns(4)
            with cols[0]: st.metric("Pipe", f"{tube_cost_kzt:,.0f} ₸", f"{tube_length} m × {tube_data['Cost']} ₸/m")
            with cols[1]: st.metric("Rod", f"{rod_cost_kzt:,.0f} ₸", f"{rod_length} m × {rod_data['Cost']} ₸/m")
            with cols[2]: st.metric("Piston", f"{piston_cost_kzt:,.0f} ₸", f"{piston_length} mm × {piston_data['Cost']} ₸/m")
            with cols[3]: st.metric("Head", f"{head_cost_kzt:,.0f} ₸", f"{head_weight} {head_data['Units']} × {head_data['Cost']} ₸/{head_data['Units']}")

            st.metric("Total materials", f"{materials_cost_kzt:,.0f} ₸", f"Margin {material_margin}%")
            st.metric("Spare parts cost", f"{spare_parts_cost:,.0f} ₸")

        with st.container():
            st.subheader("🛠 Labor (with margin)")
            cols = st.columns(2)
            with cols[0]:
                st.metric("Total labor hours", f"{labor_hours:.1f} hrs")
                st.metric("Workshop rate", f"{workshop_rate:,.0f} ₸/hr")
            with cols[1]:
                st.metric("Labor cost", f"{labor_cost_kzt:,.0f} ₸", f"Margin {labor_margin}%")

        with st.container():
            st.subheader("💰 Final Calculation")
            cols = st.columns(2)
            with cols[0]: st.metric("Subtotal (excl. VAT)", f"{subtotal_kzt:,.0f} ₸")
            with cols[1]: st.metric("VAT", f"{vat}%")

            st.divider()

            cols = st.columns(2)
            with cols[0]: st.metric("Total cost (KZT)", f"{final_price_kzt:,.0f} ₸", "Incl. VAT")
            with cols[1]: st.metric("Total cost (USD)", f"{final_price_usd:,.2f} $", f"Rate {usd_rate} ₸/$")

        # Export current calculation
        export_data = pd.DataFrame({
            "Item": ["Pipe", "Rod", "Piston", "Head", "Spare Parts", "Labor", "Total"],
            "Cost (KZT)": [tube_cost_kzt, rod_cost_kzt, piston_cost_kzt, head_cost_kzt, spare_parts_cost, labor_cost_kzt, final_price_kzt],
            "Cost (USD)": [tube_cost_kzt / usd_rate, rod_cost_kzt / usd_rate, piston_cost_kzt / usd_rate, head_cost_kzt / usd_rate, spare_parts_cost / usd_rate, labor_cost_kzt / usd_rate, final_price_usd]
        })

        st.download_button(
            label="📥 Download current calculation (CSV)",
            data=export_data.to_csv(index=False).encode('utf-8-sig'),
            file_name=f"repair_{job_order if job_order else datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

    # Display history
    st.header("📜 Repair History")

    # History filters
    with st.expander("🔍 Filters", expanded=True):
        cols = st.columns(2)
        with cols[0]:
            date_range = st.date_input(
                "Date range",
                value=[datetime.now() - timedelta(days=30), datetime.now()],
                max_value=datetime.now()
            )
        with cols[1]:
            client_filter = st.text_input("Client filter")

    # History management
    cols = st.columns(3)
    with cols[0]:
        if st.button("Refresh history"):
            st.experimental_rerun()
    with cols[1]:
        if st.button("Create backup"):
            backup_file = create_backup()
            if backup_file:
                st.success(f"Backup created: {backup_file}")
    with cols[2]:
        if st.button("Clear history", type="secondary"):
            if Path(HISTORY_DB).exists():
                Path(HISTORY_DB).unlink()
                init_database()
                st.success("History cleared")
            else:
                st.warning("Database file not found")

    # Load and display history
    history_df = load_repairs_from_db()

    if not history_df.empty:
        # Apply filters
        if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
            start_date, end_date = date_range
            mask = (pd.to_datetime(history_df['date']) >= pd.to_datetime(start_date)) & \
                   (pd.to_datetime(history_df['date']) <= pd.to_datetime(end_date))
            history_df = history_df[mask]

        if client_filter:
            history_df = history_df[history_df['client'].str.contains(client_filter, case=False, na=False)]

        # Display data
        st.dataframe(
            history_df.drop(columns=['id']),
            column_config={
                "date": st.column_config.DatetimeColumn("Date", format="DD.MM.YYYY HH:mm"),
                "materials_cost": st.column_config.NumberColumn("Materials (KZT)", format="%.0f ₸"),
                "spare_parts_cost": st.column_config.NumberColumn("Spare Parts (KZT)", format="%.0f ₸"),
                "labor_cost": st.column_config.NumberColumn("Labor (KZT)", format="%.0f ₸"),
                "total_cost_kzt": st.column_config.NumberColumn("Total (KZT)", format="%.0f ₸"),
                "total_cost_usd": st.column_config.NumberColumn("Total (USD)", format="%.2f $")
            },
            use_container_width=True,
            hide_index=True
        )

        # Statistics
        st.subheader("📈 Statistics")
        cols = st.columns(3)
        with cols[0]:
            st.metric("Total repairs", len(history_df))
        with cols[1]:
            st.metric("Total cost (KZT)", f"{history_df['total_cost_kzt'].sum():,.0f} ₸")
        with cols[2]:
            st.metric("Total cost (USD)", f"{history_df['total_cost_usd'].sum():,.2f} $")

        # Export
        st.download_button(
            label="📥 Download full history (CSV)",
            data=history_df.drop(columns=['id']).to_csv(index=False).encode('utf-8-sig'),
            file_name="repair_history.csv",
            mime="text/csv"
        )
    else:
        st.info("No repair history available")

# ============================================
# PRICE MANAGEMENT
# ============================================

def price_management():
    with st.container():
        st.header("📊 Управление ценами на материалы")

        with st.form("price_form"):
            uploaded_file = st.file_uploader("Загрузите файл с ценами (Excel или CSV)", type=["xlsx", "csv"])

            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith('.xlsx'):
                        new_data = pd.read_excel(uploaded_file)
                    else:
                        try:
                            new_data = pd.read_csv(uploaded_file)
                        except:
                            uploaded_file.seek(0)
                            new_data = pd.read_csv(uploaded_file, sep=';')

                    required_columns = ["MaterialName", "Cost", "Units"]
                    missing_cols = [col for col in required_columns if col not in new_data.columns]

                    if not missing_cols:
                        if "LastUpdated" not in new_data.columns:
                            new_data["LastUpdated"] = datetime.now().date()

                        new_data.to_csv(MATERIALS_FILE, index=False)
                        st.session_state.price_data = new_data
                        st.session_state.last_upload = datetime.now()
                        st.success("Данные успешно загружены и сохранены!")
                    else:
                        st.error(f"Отсутствуют колонки: {', '.join(missing_cols)}")
                except Exception as e:
                    st.error(f"Ошибка при загрузке: {str(e)}")

            with st.expander("✏️ Текущие данные о ценах", expanded=True):
                if st.session_state.price_data.empty:
                    st.warning("Нет данных для отображения")
                else:
                    edited_data = st.data_editor(
                        st.session_state.price_data,
                        column_config={
                            "Cost": st.column_config.NumberColumn("Цена (KZT)", min_value=0, format="%.0f"),
                            "LastUpdated": st.column_config.DateColumn("Обновлено", format="DD.MM.YYYY", disabled=True)
                        },
                        num_rows="dynamic"
                    )

                    submitted = st.form_submit_button("Сохранить изменения", type="primary")
                    if submitted:
                        if not edited_data.equals(st.session_state.price_data):
                            mask = edited_data.ne(st.session_state.price_data).any(axis=1)
                            edited_data.loc[mask, "LastUpdated"] = datetime.now().date()
                            edited_data.to_csv(MATERIALS_FILE, index=False)
                            st.session_state.price_data = edited_data
                            st.success("Изменения сохранены в файл!")

# ============================================
# APP EXECUTION
# ============================================

if __name__ == "__main__":
    tab1, tab2 = st.tabs(["🧮 Calculator", "📊 Price Management"])

    with tab1:
        main_calculator()

    with tab2:
        price_management()

    # Updated app footer
    st.markdown("---")
    st.markdown(
        f"""
        <div style="text-align: center; color: {text_color}; padding: 20px; font-size: 0.9rem;">
            <p>© 2025 BMC Hydraulic Cylinders | Version 1.2</p>
            <p style="margin-top: 0.5rem;">Technical Support: Andrey Umarov & Almaz Kenzhetaev</p>
        </div>
        """,
        unsafe_allow_html=True
    )