import io
import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# Configure page
st.set_page_config(
    page_title="Control de Proveedores",
    page_icon="🚚",
    layout="wide"
)

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
# 4. Initialize Session State
# ─────────────────────────────────────────────────────────────
if 'form_step' not in st.session_state:
    st.session_state.form_step = 1
if 'selected_order' not in st.session_state:
    st.session_state.selected_order = None
if 'form_data' not in st.session_state:
    st.session_state.form_data = {}

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
    
    # Display current step
    step_col1, step_col2 = st.columns(2)
    with step_col1:
        if st.session_state.form_step == 1:
            st.info("📍 **Paso 1:** Registro de Llegada")
        else:
            st.success("✅ **Paso 1:** Llegada registrada")
    
    with step_col2:
        if st.session_state.form_step == 2:
            st.info("🔄 **Paso 2:** Registro de Atención")
        elif st.session_state.form_step > 2:
            st.success("✅ **Paso 2:** Atención registrada")
    
    st.markdown("---")
    
    # Form Step 1: Registration of arrival
    if st.session_state.form_step == 1:
        st.subheader("📍 Registro de Llegada del Proveedor")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Order selection
            order_options = today_reservations['Orden_de_compra'].tolist()
            selected_order = st.selectbox(
                "Orden de Compra:",
                options=order_options,
                key="order_select"
            )
            
            if selected_order:
                # Get order details
                order_details = today_reservations[
                    today_reservations['Orden_de_compra'] == selected_order
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
            if selected_order:
                order_details = today_reservations[
                    today_reservations['Orden_de_compra'] == selected_order
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
                    key="arrival_hour"
                )
            
            with time_col2:
                arrival_minute = st.selectbox(
                    "Minutos:",
                    options=list(range(0, 60, 5)),  # 5-minute intervals
                    index=default_minute // 5,  # Find closest 5-minute interval
                    format_func=lambda x: f"{x:02d}",
                    key="arrival_minute"
                )
            
            # Combine into time object
            arrival_time = time(arrival_hour, arrival_minute)
            
            st.info(f"Fecha: {today_date.strftime('%Y-%m-%d')}")
        
        # Next step button
        if st.button("Continuar a Registro de Atención", type="primary"):
            if selected_order and arrival_time:
                # Save step 1 data
                order_details = today_reservations[
                    today_reservations['Orden_de_compra'] == selected_order
                ].iloc[0]
                
                st.session_state.form_data = {
                    'Orden_de_compra': selected_order,
                    'Proveedor': order_details['Proveedor'],
                    'Numero_de_bultos': order_details['Numero_de_bultos'],
                    'Hora_llegada': combine_date_time(today_date, arrival_time),
                    'booked_time_range': order_details['Hora']
                }
                st.session_state.form_step = 2
                st.rerun()
            else:
                st.error("Por favor complete todos los campos.")
    
    # Form Step 2: Service attention registration
    elif st.session_state.form_step == 2:
        st.subheader("🔄 Registro de Atención al Proveedor")
        
        # Display order info
        st.info(f"**Orden de Compra:** {st.session_state.form_data['Orden_de_compra']} | "
                f"**Proveedor:** {st.session_state.form_data['Proveedor']} | "
                f"**Llegada:** {st.session_state.form_data['Hora_llegada'].strftime('%H:%M')}")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Hora de Inicio de Atención:**")
            current_time = datetime.now()
            
            start_time_col1, start_time_col2 = st.columns(2)
            with start_time_col1:
                start_hour = st.selectbox(
                    "Hora:",
                    options=list(range(0, 24)),
                    index=current_time.hour,
                    format_func=lambda x: f"{x:02d}",
                    key="start_hour"
                )
            
            with start_time_col2:
                start_minute = st.selectbox(
                    "Minutos:",
                    options=list(range(0, 60, 5)),  # 5-minute intervals
                    index=current_time.minute // 5,
                    format_func=lambda x: f"{x:02d}",
                    key="start_minute"
                )
            
            start_time = time(start_hour, start_minute)
        
        with col2:
            st.write("**Hora de Fin de Atención:**")
            
            end_time_col1, end_time_col2 = st.columns(2)
            with end_time_col1:
                end_hour = st.selectbox(
                    "Hora:",
                    options=list(range(0, 24)),
                    index=current_time.hour,
                    format_func=lambda x: f"{x:02d}",
                    key="end_hour"
                )
            
            with end_time_col2:
                end_minute = st.selectbox(
                    "Minutos:",
                    options=list(range(0, 60, 5)),  # 5-minute intervals
                    index=current_time.minute // 5,
                    format_func=lambda x: f"{x:02d}",
                    key="end_minute"
                )
            
            end_time = time(end_hour, end_minute)
        
        # Action buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("← Volver a Registro de Llegada"):
                st.session_state.form_step = 1
                st.rerun()
        
        with col2:
            if st.button("Finalizar y Guardar", type="primary"):
                if start_time and end_time:
                    # Complete the form data with step 2 info
                    today_date = datetime.now().date()
                    hora_inicio = combine_date_time(today_date, start_time)
                    hora_fin = combine_date_time(today_date, end_time)
                    
                    # Validate times
                    if hora_inicio >= hora_fin:
                        st.error("La hora de fin debe ser posterior a la hora de inicio.")
                    elif hora_inicio < st.session_state.form_data['Hora_llegada']:
                        st.error("La hora de inicio de atención no puede ser anterior a la hora de llegada.")
                    else:
                        # Calculate times
                        tiempo_espera = calculate_time_difference(
                            st.session_state.form_data['Hora_llegada'], 
                            hora_inicio
                        )
                        tiempo_atencion = calculate_time_difference(hora_inicio, hora_fin)
                        tiempo_total = calculate_time_difference(
                            st.session_state.form_data['Hora_llegada'], 
                            hora_fin
                        )
                        
                        # Calculate delay time
                        booked_start_time = parse_time_range(st.session_state.form_data['booked_time_range'])
                        tiempo_retraso = None
                        if booked_start_time:
                            booked_datetime = combine_date_time(today_date, booked_start_time)
                            tiempo_retraso = calculate_time_difference(
                                booked_datetime, 
                                st.session_state.form_data['Hora_llegada']
                            )
                        
                        # Prepare final record
                        final_record = {
                            'Orden_de_compra': st.session_state.form_data['Orden_de_compra'],
                            'Proveedor': st.session_state.form_data['Proveedor'],
                            'Numero_de_bultos': st.session_state.form_data['Numero_de_bultos'],
                            'Hora_llegada': st.session_state.form_data['Hora_llegada'].strftime('%Y-%m-%d %H:%M:%S'),
                            'Hora_inicio_atencion': hora_inicio.strftime('%Y-%m-%d %H:%M:%S'),
                            'Hora_fin_atencion': hora_fin.strftime('%Y-%m-%d %H:%M:%S'),
                            'Tiempo_espera': tiempo_espera,
                            'Tiempo_atencion': tiempo_atencion,
                            'Tiempo_total': tiempo_total,
                            'Tiempo_retraso': tiempo_retraso
                        }
                        
                        # Save to Excel
                        with st.spinner("Guardando registro..."):
                            if save_gestion_to_excel(final_record):
                                st.success("✅ Registro guardado exitosamente!")
                                
                                # Show summary
                                st.subheader("📊 Resumen del Registro")
                                
                                # Debug information (can be removed in production)
                                with st.expander("🔍 Información de Debug"):
                                    st.write(f"Hora llegada: {st.session_state.form_data['Hora_llegada']}")
                                    st.write(f"Hora inicio: {hora_inicio}")
                                    st.write(f"Hora fin: {hora_fin}")
                                    st.write(f"Tiempo espera calculado: {tiempo_espera}")
                                    st.write(f"Tiempo atención calculado: {tiempo_atencion}")
                                    st.write(f"Tiempo total calculado: {tiempo_total}")
                                    st.write(f"Tiempo retraso calculado: {tiempo_retraso}")
                                
                                summary_col1, summary_col2 = st.columns(2)
                                
                                with summary_col1:
                                    st.metric("Tiempo de Espera", f"{tiempo_espera if tiempo_espera is not None else 'N/A'} min")
                                    st.metric("Tiempo de Atención", f"{tiempo_atencion if tiempo_atencion is not None else 'N/A'} min")
                                
                                with summary_col2:
                                    st.metric("Tiempo Total", f"{tiempo_total if tiempo_total is not None else 'N/A'} min")
                                    if tiempo_retraso is not None:
                                        if tiempo_retraso > 0:
                                            st.metric("Retraso", f"{tiempo_retraso} min", delta=f"+{tiempo_retraso}")
                                        elif tiempo_retraso < 0:
                                            st.metric("Adelanto", f"{abs(tiempo_retraso)} min", delta=tiempo_retraso)
                                        else:
                                            st.metric("Puntualidad", "A tiempo", delta=0)
                                    else:
                                        st.metric("Retraso", "N/A")
                                
                                # Reset form
                                if st.button("Registrar Nuevo Proveedor"):
                                    st.session_state.form_step = 1
                                    st.session_state.form_data = {}
                                    st.rerun()
                            else:
                                st.error("Error al guardar el registro. Intente nuevamente.")
                else:
                    st.error("Por favor complete todos los campos de tiempo.")

if __name__ == "__main__":
    main()