import io
import os
import streamlit as st
import pandas as pd
import time
from datetime import datetime, timedelta, time as dt_time
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# Configure page
st.set_page_config(
    page_title="Control de Proveedores",
    page_icon="🚚",
    layout="wide"
)

# Custom CSS for enhanced tab visibility
st.markdown("""
<style>
/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 20px;
    background-color: #f0f2f6;
    padding: 10px;
    border-radius: 10px;
    margin-bottom: 20px;
}

.stTabs [data-baseweb="tab"] {
    height: 60px;
    background-color: white;
    border-radius: 8px;
    padding: 0 20px;
    border: 2px solid #e1e5e9;
    font-weight: bold;
    font-size: 16px;
}

.stTabs [aria-selected="true"] {
    background-color: #1f77b4 !important;
    color: white !important;
    border-color: #1f77b4 !important;
    box-shadow: 0 4px 8px rgba(31, 119, 180, 0.3);
}

/* Arrival tab content styling */
.arrival-container {
    background: linear-gradient(135deg, #e3f2fd 0%, #f3e5f5 100%);
    border: 3px solid #2196f3;
    border-radius: 15px;
    padding: 25px;
    margin: 15px 0;
    box-shadow: 0 6px 20px rgba(33, 150, 243, 0.15);
}

.arrival-header {
    background-color: #2196f3;
    color: white;
    padding: 15px;
    border-radius: 10px;
    margin-bottom: 20px;
    text-align: center;
    font-weight: bold;
    font-size: 18px;
}

/* Service tab content styling */
.service-container {
    background: linear-gradient(135deg, #e8f5e8 0%, #fff3e0 100%);
    border: 3px solid #4caf50;
    border-radius: 15px;
    padding: 25px;
    margin: 15px 0;
    box-shadow: 0 6px 20px rgba(76, 175, 80, 0.15);
}

.service-header {
    background-color: #4caf50;
    color: white;
    padding: 15px;
    border-radius: 10px;
    margin-bottom: 20px;
    text-align: center;
    font-weight: bold;
    font-size: 18px;
}

/* Button styling */
.arrival-container .stButton > button {
    background-color: #2196f3;
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: bold;
    padding: 10px 20px;
    box-shadow: 0 3px 6px rgba(33, 150, 243, 0.3);
}

.service-container .stButton > button {
    background-color: #4caf50;
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: bold;
    padding: 10px 20px;
    box-shadow: 0 3px 6px rgba(76, 175, 80, 0.3);
}

/* Info boxes */
.arrival-info {
    background-color: rgba(33, 150, 243, 0.1);
    border-left: 5px solid #2196f3;
    padding: 15px;
    border-radius: 0 8px 8px 0;
    margin: 10px 0;
}

.service-info {
    background-color: rgba(76, 175, 80, 0.1);
    border-left: 5px solid #4caf50;
    padding: 15px;
    border-radius: 0 8px 8px 0;
    margin: 10px 0;
}

/* Visual separator */
.tab-separator {
    height: 4px;
    background: linear-gradient(90deg, #2196f3 0%, #4caf50 100%);
    margin: 20px 0;
    border-radius: 2px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 1. Configuration
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    FILE_ID = os.getenv("SP_FILE_ID") or st.secrets["SP_FILE_ID"]
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
except KeyError as e:
    st.error(f"Missing required environment variable or secret: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. Excel Download Functions
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)  # Cache for 5 minutes
def download_excel_to_memory():
    """Download Excel file from SharePoint to memory"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        # Download to memory
        file_content = io.BytesIO()
        
        # Try multiple download methods
        try:
            file.download(file_content)
            ctx.execute_query()
        except TypeError:
            try:
                response = file.download()
                ctx.execute_query()
                file_content = io.BytesIO(response.content)
            except:
                file.download_session(file_content)
                ctx.execute_query()
        
        file_content.seek(0)
        
        # Load all sheets
        credentials_df = pd.read_excel(file_content, sheet_name="proveedor_credencial", dtype=str)
        reservas_df = pd.read_excel(file_content, sheet_name="proveedor_reservas")
        
        # Try to load gestion sheet, create if doesn't exist
        try:
            gestion_df = pd.read_excel(file_content, sheet_name="proveedor_gestion")
        except ValueError:
            # Create empty gestion dataframe with required columns
            gestion_df = pd.DataFrame(columns=[
                'Orden_de_compra', 'Proveedor', 'Numero_de_bultos',
                'Hora_llegada', 'Hora_inicio_atencion', 'Hora_fin_atencion',
                'Tiempo_espera', 'Tiempo_atencion', 'Tiempo_total', 'Tiempo_retraso'
            ])
        
        return credentials_df, reservas_df, gestion_df
        
    except Exception as e:
        st.error(f"Error descargando Excel: {str(e)}")
        return None, None, None

def save_gestion_to_excel(new_record):
    """Save new management record to Excel file"""
    try:
        # Load current data
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if reservas_df is None:
            return False
        
        # Add new record
        new_row = pd.DataFrame([new_record])
        updated_gestion_df = pd.concat([gestion_df, new_row], ignore_index=True)
        
        # Authenticate and upload
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
            updated_gestion_df.to_excel(writer, sheet_name="proveedor_gestion", index=False)
        
        # Get the file info
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        file_name = file.properties['Name']
        server_relative_url = file.properties['ServerRelativeUrl']
        folder_url = server_relative_url.replace('/' + file_name, '')
        
        # Upload the updated file
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        excel_buffer.seek(0)
        folder.files.add(file_name, excel_buffer.getvalue(), True)
        ctx.execute_query()
        
        # Clear cache
        download_excel_to_memory.clear()
        
        return True
        
    except Exception as e:
        st.error(f"Error guardando registro: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 3. Helper Functions
# ─────────────────────────────────────────────────────────────
def get_today_reservations(reservas_df):
    """Get today's reservations"""
    today = datetime.now().strftime('%Y-%m-%d')
    return reservas_df[reservas_df['Fecha'].astype(str).str.contains(today, na=False)]

def parse_time_range(time_range_str):
    """Parse time range string (e.g., '09:00-09:30' or '09:00 - 09:30') and return start time"""
    try:
        # Handle both formats: "12:00-12:30" and "12:00 - 12:30"
        if '-' in time_range_str:
            start_time_str = time_range_str.split('-')[0].strip()
            return datetime.strptime(start_time_str, '%H:%M').time()
        return None
    except:
        return None

def calculate_time_difference(start_datetime, end_datetime):
    """Calculate time difference in minutes"""
    if start_datetime and end_datetime:
        # Ensure both are datetime objects
        if isinstance(start_datetime, str):
            start_datetime = datetime.fromisoformat(start_datetime)
        if isinstance(end_datetime, str):
            end_datetime = datetime.fromisoformat(end_datetime)
            
        diff = end_datetime - start_datetime
        return int(diff.total_seconds() / 60)
    return None

def combine_date_time(date_part, time_part):
    """Combine date and time into datetime"""
    return datetime.combine(date_part, time_part)

# ─────────────────────────────────────────────────────────────
# 4. Helper Functions for Data Management
# ─────────────────────────────────────────────────────────────
def get_existing_arrivals(gestion_df):
    """Get orders that already have arrival registered today but not yet completed"""
    today = datetime.now().strftime('%Y-%m-%d')
    if gestion_df.empty:
        return []
    
    # Filter records with arrival time from today
    today_arrivals = gestion_df[
        gestion_df['Hora_llegada'].astype(str).str.contains(today, na=False)
    ]
    
    # Only return orders that don't have service times completed
    pending_service = today_arrivals[
        today_arrivals['Hora_inicio_atencion'].isna() | 
        today_arrivals['Hora_fin_atencion'].isna()
    ]
    
    return pending_service['Orden_de_compra'].tolist()

def get_completed_orders(gestion_df):
    """Get orders that have both arrival and service registered today"""
    today = datetime.now().strftime('%Y-%m-%d')
    if gestion_df.empty:
        return []
    
    # Filter records with arrival time from today
    today_records = gestion_df[
        gestion_df['Hora_llegada'].astype(str).str.contains(today, na=False)
    ]
    
    # Return orders that have both arrival and service times
    completed = today_records[
        today_records['Hora_inicio_atencion'].notna() & 
        today_records['Hora_fin_atencion'].notna()
    ]
    
    return completed['Orden_de_compra'].tolist()

def get_pending_arrivals(today_reservations, gestion_df):
    """Get orders that haven't registered arrival yet"""
    existing_arrivals = get_existing_arrivals(gestion_df)
    completed_orders = get_completed_orders(gestion_df)
    
    # Combine both lists to exclude from dropdown
    processed_orders = existing_arrivals + completed_orders
    
    # Return orders that haven't been processed at all
    pending = today_reservations[
        ~today_reservations['Orden_de_compra'].isin(processed_orders)
    ]
    
    return pending['Orden_de_compra'].tolist()

def get_arrival_record(gestion_df, orden_compra):
    """Get existing arrival record for an order"""
    if gestion_df.empty:
        return None
    
    record = gestion_df[gestion_df['Orden_de_compra'] == orden_compra]
    return record.iloc[0] if not record.empty else None

def save_arrival_to_excel(arrival_data):
    """Save arrival data to Excel file"""
    try:
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if reservas_df is None:
            return False
        
        # Check if record already exists
        existing_record = get_arrival_record(gestion_df, arrival_data['Orden_de_compra'])
        
        if existing_record is not None:
            # Update existing record
            gestion_df.loc[
                gestion_df['Orden_de_compra'] == arrival_data['Orden_de_compra'], 
                'Hora_llegada'
            ] = arrival_data['Hora_llegada']
            updated_gestion_df = gestion_df
        else:
            # Add new record
            new_row = pd.DataFrame([arrival_data])
            updated_gestion_df = pd.concat([gestion_df, new_row], ignore_index=True)
        
        return upload_excel_file(credentials_df, reservas_df, updated_gestion_df)
        
    except Exception as e:
        st.error(f"Error guardando llegada: {str(e)}")
        return False

def update_service_times(orden_compra, service_data):
    """Update service times for existing arrival record"""
    try:
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if gestion_df.empty:
            return False
        
        # Find the record to update
        mask = gestion_df['Orden_de_compra'] == orden_compra
        if not mask.any():
            st.error("No se encontró registro de llegada para esta orden.")
            return False
        
        # Update service times and calculations
        gestion_df.loc[mask, 'Hora_inicio_atencion'] = service_data['Hora_inicio_atencion']
        gestion_df.loc[mask, 'Hora_fin_atencion'] = service_data['Hora_fin_atencion']
        gestion_df.loc[mask, 'Tiempo_espera'] = service_data['Tiempo_espera']
        gestion_df.loc[mask, 'Tiempo_atencion'] = service_data['Tiempo_atencion']
        gestion_df.loc[mask, 'Tiempo_total'] = service_data['Tiempo_total']
        
        return upload_excel_file(credentials_df, reservas_df, gestion_df)
        
    except Exception as e:
        st.error(f"Error actualizando tiempos de atención: {str(e)}")
        return False

def upload_excel_file(credentials_df, reservas_df, gestion_df):
    """Upload updated Excel file to SharePoint"""
    try:
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
            gestion_df.to_excel(writer, sheet_name="proveedor_gestion", index=False)
        
        # Get the file info and upload
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        file_name = file.properties['Name']
        server_relative_url = file.properties['ServerRelativeUrl']
        folder_url = server_relative_url.replace('/' + file_name, '')
        
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        excel_buffer.seek(0)
        folder.files.add(file_name, excel_buffer.getvalue(), True)
        ctx.execute_query()
        
        # Clear cache
        download_excel_to_memory.clear()
        return True
        
    except Exception as e:
        st.error(f"Error subiendo archivo: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 5. Main App
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🚚 Control de Proveedores")
    st.markdown("---")
    
    # Load data
    with st.spinner("Cargando datos..."):
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
    
    if reservas_df is None:
        st.error("No se pudo cargar los datos. Verifique la conexión.")
        return
    
    # Get today's reservations
    today_reservations = get_today_reservations(reservas_df)
    
    if today_reservations.empty:
        st.warning("No hay reservas programadas para hoy.")
        return
    
    # Get order status
    existing_arrivals = get_existing_arrivals(gestion_df)
    completed_orders = get_completed_orders(gestion_df)
    pending_arrivals = get_pending_arrivals(today_reservations, gestion_df)
    
    # Create tabs with enhanced styling
    tab1, tab2 = st.tabs(["🚚 REGISTRO DE LLEGADA", "⚙️ REGISTRO DE ATENCIÓN"])
    
    # Visual separator
    st.markdown('<div class="tab-separator"></div>', unsafe_allow_html=True)
    
    # ─────────────────────────────────────────────────────────────
    # TAB 1: Arrival Registration
    # ─────────────────────────────────────────────────────────────
    with tab1:
        with st.container():
            st.markdown('<div class="arrival-container">', unsafe_allow_html=True)
            
            st.markdown("*Registre la hora de llegada del proveedor*")
            
            col1, col2 = st.columns(2)
        
        with col1:
            # Order selection - only show orders that haven't been processed
            if not pending_arrivals:
                st.info("✅ Todas las llegadas del día han sido registradas")
                selected_order_tab1 = None
            else:
                selected_order_tab1 = st.selectbox(
                    "Orden de Compra:",
                    options=pending_arrivals,
                    key="order_select_tab1"
                )
            
            if selected_order_tab1:
                # Get order details
                order_details = today_reservations[
                    today_reservations['Orden_de_compra'] == selected_order_tab1
                ].iloc[0]
                
                # Auto-fill fields
                st.text_input(
                    "Proveedor:",
                    value=order_details['Proveedor'],
                    disabled=True
                )
                
                st.text_input(
                    "Número de Bultos:",
                    value=str(order_details['Numero_de_bultos']),
                    disabled=True
                )
        
        with col2:
            # Arrival time input with friendly UI
            st.write("**Hora de Llegada:**")
            today_date = datetime.now().date()
            
            # Get default time from booked hour
            if selected_order_tab1:
                order_details = today_reservations[
                    today_reservations['Orden_de_compra'] == selected_order_tab1
                ].iloc[0]
                booked_start_time = parse_time_range(str(order_details['Hora']))
                if booked_start_time:
                    default_hour = booked_start_time.hour
                    default_minute = booked_start_time.minute
                else:
                    default_hour = datetime.now().hour
                    default_minute = datetime.now().minute
            else:
                default_hour = datetime.now().hour
                default_minute = datetime.now().minute
            
            # Create user-friendly time picker
            time_col1, time_col2 = st.columns(2)
            with time_col1:
                arrival_hour = st.selectbox(
                    "Hora:",
                    options=list(range(0, 24)),
                    index=default_hour,
                    format_func=lambda x: f"{x:02d}",
                    key="arrival_hour_tab1"
                )
            
            with time_col2:
                arrival_minute = st.selectbox(
                    "Minutos:",
                    options=list(range(0, 60, 5)),  # 5-minute intervals
                    index=default_minute // 5,  # Find closest 5-minute interval
                    format_func=lambda x: f"{x:02d}",
                    key="arrival_minute_tab1"
                )
            
            # Combine into time object
            arrival_time = dt_time(arrival_hour, arrival_minute)
            
            st.info(f"Fecha: {today_date.strftime('%Y-%m-%d')}")
        
        # Save arrival button
        if st.button("Guardar Llegada", type="primary", key="save_arrival"):
            if selected_order_tab1 and arrival_time:
                # Get order details for delay calculation
                order_details = today_reservations[
                    today_reservations['Orden_de_compra'] == selected_order_tab1
                ].iloc[0]
                
                arrival_datetime = combine_date_time(today_date, arrival_time)
                
                # Calculate delay
                booked_start_time = parse_time_range(str(order_details['Hora']))
                tiempo_retraso = None
                if booked_start_time:
                    booked_datetime = combine_date_time(today_date, booked_start_time)
                    tiempo_retraso = calculate_time_difference(booked_datetime, arrival_datetime)
                
                # Prepare arrival data
                arrival_data = {
                    'Orden_de_compra': selected_order_tab1,
                    'Proveedor': order_details['Proveedor'],
                    'Numero_de_bultos': order_details['Numero_de_bultos'],
                    'Hora_llegada': arrival_datetime.strftime('%Y-%m-%d %H:%M:%S'),
                    'Hora_inicio_atencion': None,
                    'Hora_fin_atencion': None,
                    'Tiempo_espera': None,
                    'Tiempo_atencion': None,
                    'Tiempo_total': None,
                    'Tiempo_retraso': tiempo_retraso
                }
                
                # Save to Excel
                with st.spinner("Guardando llegada..."):
                    if save_arrival_to_excel(arrival_data):
                        st.success("✅ Llegada registrada exitosamente!")
                        if tiempo_retraso is not None:
                            if tiempo_retraso > 0:
                                st.warning(f"⏰ Retraso: {tiempo_retraso} minutos")
                            elif tiempo_retraso < 0:
                                st.info(f"⚡ Adelanto: {abs(tiempo_retraso)} minutos")
                            else:
                                st.success("🎯 Llegada puntual")
                        
                        # Wait 5 seconds before refreshing
                        with st.spinner("Actualizando datos..."):
                            time.sleep(5)
                        st.rerun()
                    else:
                        st.error("Error al guardar la llegada. Intente nuevamente.")
            else:
                st.error("Por favor complete todos los campos.")
    
    # ─────────────────────────────────────────────────────────────
    # TAB 2: Service Registration
    # ─────────────────────────────────────────────────────────────
    with tab2:
        with st.container():
            st.markdown('<div class="service-container">', unsafe_allow_html=True)
            
            st.markdown("*Registre los tiempos de inicio y fin de atención*")
            
            # Order selection
            selected_order_tab2 = st.selectbox(
                "Orden de Compra:",
                options=existing_arrivals if existing_arrivals else ["No hay llegadas registradas"],
                disabled=not existing_arrivals,
                key="order_select_tab2"
            )
        
            if existing_arrivals and selected_order_tab2:
                # Get arrival record
                arrival_record = get_arrival_record(gestion_df, selected_order_tab2)
                
                if arrival_record is not None:
                    # Show arrival info
                    arrival_time_str = str(arrival_record['Hora_llegada'])
                    st.markdown(f'''
                    <div class="service-info">
                        <strong>Proveedor:</strong> {arrival_record['Proveedor']} | 
                        <strong>Llegada:</strong> {arrival_time_str.split(' ')[1][:5] if ' ' in arrival_time_str else 'N/A'}
                    </div>
                    ''', unsafe_allow_html=True)
                    
                    # Check if service times already registered
                    service_registered = (
                        pd.notna(arrival_record['Hora_inicio_atencion']) and 
                        pd.notna(arrival_record['Hora_fin_atencion'])
                    )
                    
                    if service_registered:
                        st.success("✅ Atención ya registrada")
                        # Show existing times
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Tiempo de Espera", f"{arrival_record['Tiempo_espera']} min")
                            st.metric("Tiempo de Atención", f"{arrival_record['Tiempo_atencion']} min")
                        with col2:
                            st.metric("Tiempo Total", f"{arrival_record['Tiempo_total']} min")
                    else:
                        st.warning("⏳ Pendiente de registrar atención")
                        
                        # Service time inputs - only show when not registered
                        col1, col2 = st.columns(2)
                        
                        # Parse arrival time for defaults
                        arrival_datetime = datetime.fromisoformat(str(arrival_record['Hora_llegada']))
                        default_hour = arrival_datetime.hour
                        default_minute = (arrival_datetime.minute // 5) * 5  # Round to nearest 5-minute interval
                        
                        with col1:
                            st.write("**Hora de Inicio de Atención:**")
                            
                            start_time_col1, start_time_col2 = st.columns(2)
                            with start_time_col1:
                                start_hour = st.selectbox(
                                    "Hora:",
                                    options=list(range(0, 24)),
                                    index=default_hour,
                                    format_func=lambda x: f"{x:02d}",
                                    key="start_hour_tab2"
                                )
                            
                            with start_time_col2:
                                start_minute = st.selectbox(
                                    "Minutos:",
                                    options=list(range(0, 60, 5)),  # 5-minute intervals
                                    index=default_minute // 5,
                                    format_func=lambda x: f"{x:02d}",
                                    key="start_minute_tab2"
                                )
                            
                            start_time = dt_time(start_hour, start_minute)
                        
                        with col2:
                            st.write("**Hora de Fin de Atención:**")
                            
                            end_time_col1, end_time_col2 = st.columns(2)
                            with end_time_col1:
                                end_hour = st.selectbox(
                                    "Hora:",
                                    options=list(range(0, 24)),
                                    index=default_hour,
                                    format_func=lambda x: f"{x:02d}",
                                    key="end_hour_tab2"
                                )
                            
                            with end_time_col2:
                                end_minute = st.selectbox(
                                    "Minutos:",
                                    options=list(range(0, 60, 5)),  # 5-minute intervals
                                    index=default_minute // 5,
                                    format_func=lambda x: f"{x:02d}",
                                    key="end_minute_tab2"
                                )
                            
                            end_time = dt_time(end_hour, end_minute)
                        
                        # Save service times button - only show when not registered
                        if st.button("Guardar Atención", type="primary", key="save_service"):
                            if start_time and end_time:
                                today_date = datetime.now().date()
                                hora_inicio = combine_date_time(today_date, start_time)
                                hora_fin = combine_date_time(today_date, end_time)
                                
                                # Parse arrival time
                                arrival_datetime = datetime.fromisoformat(str(arrival_record['Hora_llegada']))
                                
                                # Validate times
                                if hora_inicio >= hora_fin:
                                    st.error("La hora de fin debe ser posterior a la hora de inicio.")
                                elif hora_inicio < arrival_datetime:
                                    st.error("La hora de inicio de atención no puede ser anterior a la hora de llegada.")
                                else:
                                    # Calculate times
                                    tiempo_espera = calculate_time_difference(arrival_datetime, hora_inicio)
                                    tiempo_atencion = calculate_time_difference(hora_inicio, hora_fin)
                                    tiempo_total = calculate_time_difference(arrival_datetime, hora_fin)
                                    
                                    # Prepare service data
                                    service_data = {
                                        'Hora_inicio_atencion': hora_inicio.strftime('%Y-%m-%d %H:%M:%S'),
                                        'Hora_fin_atencion': hora_fin.strftime('%Y-%m-%d %H:%M:%S'),
                                        'Tiempo_espera': tiempo_espera,
                                        'Tiempo_atencion': tiempo_atencion,
                                        'Tiempo_total': tiempo_total
                                    }
                                    
                                    # Save to Excel
                                    with st.spinner("Guardando atención..."):
                                        if update_service_times(selected_order_tab2, service_data):
                                            st.success("✅ Atención registrada exitosamente!")
                                            
                                            # Show summary
                                            col1, col2 = st.columns(2)
                                            with col1:
                                                st.metric("Tiempo de Espera", f"{tiempo_espera} min")
                                                st.metric("Tiempo de Atención", f"{tiempo_atencion} min")
                                            with col2:
                                                st.metric("Tiempo Total", f"{tiempo_total} min")
                                            
                                            # Wait 5 seconds before refreshing
                                            with st.spinner("Actualizando datos..."):
                                                time.sleep(5)
                                            st.rerun()
                                        else:
                                            st.error("Error al guardar la atención. Intente nuevamente.")
                            else:
                                st.error("Por favor complete todos los campos de tiempo.")
            else:
                st.markdown(
                    '<div class="service-info">⚠️ No hay llegadas registradas hoy. Primero debe registrar la llegada en la pestaña anterior.</div>', 
                    unsafe_allow_html=True
                )
            
            st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()