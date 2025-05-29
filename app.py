import streamlit as st
import pandas as pd
import random
from io import BytesIO
import xlsxwriter
import json

# --- App-Konfiguration ---
st.set_page_config(page_title="Praxis-Dienstplanung", layout="wide")
st.title("Praxis-Dienstplanung: Dateninput & Dienstplan")

# --- Statische Daten (nicht sensibel) ---
days = ['Montag','Dienstag','Mittwoch','Donnerstag','Freitag']
schifts = ['Vormittag','Nachmittag']

# --- Config File Loader ---
def load_config(uploaded_file):
    """Lädt und validiert die Konfigurationsdatei"""
    try:
        config = json.load(uploaded_file)
        
        # Basis-Validierung
        required_keys = [
            'bereiche', 'mitarbeiter', 'bereich_schichten', 
            'bereich_mitarbeiter', 'mitarbeiter_verfuegbarkeit',
            'mitarbeiter_bereiche', 'mitarbeiter_max_stunden'
        ]
        
        for key in required_keys:
            if key not in config:
                st.error(f"Fehlender Schlüssel in Config: {key}")
                return None
        
        return config
    except json.JSONDecodeError as e:
        st.error(f"JSON-Fehler: {e}")
        return None
    except Exception as e:
        st.error(f"Fehler beim Laden der Config: {e}")
        return None

def validate_config(config):
    """Erweiterte Validierung der Config-Struktur"""
    errors = []
    
    # Prüfe ob alle Bereiche in bereich_schichten vorhanden sind
    for bereich in config['bereiche']:
        if bereich not in config['bereich_schichten']:
            errors.append(f"Bereich '{bereich}' fehlt in bereich_schichten")
    
    # Prüfe ob alle Mitarbeiter in allen relevanten Mappings vorhanden sind
    for mitarbeiter in config['mitarbeiter']:
        if mitarbeiter not in config['mitarbeiter_verfuegbarkeit']:
            errors.append(f"Mitarbeiter '{mitarbeiter}' fehlt in mitarbeiter_verfuegbarkeit")
        if mitarbeiter not in config['mitarbeiter_bereiche']:
            errors.append(f"Mitarbeiter '{mitarbeiter}' fehlt in mitarbeiter_bereiche")
        if mitarbeiter not in config['mitarbeiter_max_stunden']:
            errors.append(f"Mitarbeiter '{mitarbeiter}' fehlt in mitarbeiter_max_stunden")
    
    return errors

# --- Excel Export Funktion ---
def create_excel_export(df_pivot, plan_data=None):
    """Erstellt eine Excel-Datei mit dem Dienstplan"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Haupttabelle schreiben
        df_pivot.to_excel(writer, sheet_name='Dienstplan', index=True)
        
        # Falls Raw-Daten verfügbar sind, als zweites Sheet hinzufügen
        if plan_data:
            df_raw = pd.DataFrame(plan_data)
            df_raw.to_excel(writer, sheet_name='Raw_Data', index=False)
        
        # Formatierung
        workbook = writer.book
        worksheet = writer.sheets['Dienstplan']
        
        # Header-Format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Standard-Format
        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        })
        
        # Spaltenbreite anpassen
        worksheet.set_column('A:A', 15)  # Slot-Spalte
        for col_num in range(1, len(df_pivot.columns) + 1):
            worksheet.set_column(col_num, col_num, 12)
        
        # Header formatieren
        for col_num, value in enumerate(df_pivot.columns):
            worksheet.write(0, col_num + 1, value, header_format)
        
        # Index-Spalte formatieren
        for row_num, value in enumerate(df_pivot.index):
            worksheet.write(row_num + 1, 0, value, header_format)
    
    processed_data = output.getvalue()
    return processed_data

# --- Session State Initialisierung ---
if 'config_loaded' not in st.session_state:
    st.session_state.config_loaded = False
if 'config_data' not in st.session_state:
    st.session_state.config_data = None
if 'bereiche_cfg' not in st.session_state:
    st.session_state.bereiche_cfg = {}
if 'helpers_cfg' not in st.session_state:
    st.session_state.helpers_cfg = {}
if 'current_plan' not in st.session_state:
    st.session_state.current_plan = None
if 'current_pivot' not in st.session_state:
    st.session_state.current_pivot = None

# --- CONFIG FILE UPLOAD SECTION ---
st.header("🔧 Konfiguration laden")

if not st.session_state.config_loaded:
    st.info("📁 Bitte lade deine Praxis-Konfigurationsdatei (JSON) hoch, um zu beginnen.")
    
    uploaded_file = st.file_uploader(
        "Konfigurationsdatei auswählen", 
        type=['json'],
        help="Lade deine praxis_config.json Datei hoch"
    )
    
    if uploaded_file is not None:
        config = load_config(uploaded_file)
        
        if config:
            # Validierung
            validation_errors = validate_config(config)
            
            if validation_errors:
                st.error("❌ Validierungsfehler in der Config:")
                for error in validation_errors:
                    st.error(f"• {error}")
            else:
                st.session_state.config_data = config
                st.session_state.config_loaded = True
                st.success("✅ Konfiguration erfolgreich geladen!")
                st.rerun()

# --- HAUPT-APP (nur wenn Config geladen) ---
if st.session_state.config_loaded and st.session_state.config_data:
    config = st.session_state.config_data
    
    # Config-Daten extrahieren
    bereiche_list = config['bereiche']
    arzthelfer_list = config['mitarbeiter']
    default_shifts_map = config['bereich_schichten']
    default_helpers_map = config['bereich_mitarbeiter']
    default_helper_shifts_map = config['mitarbeiter_verfuegbarkeit']
    default_areas_map = config['mitarbeiter_bereiche']
    default_max_hours = config['mitarbeiter_max_stunden']
    
    # Spezielle Regeln
    spezial_regeln = config.get('spezial_regeln', {})
    rezeption_prioritaet = spezial_regeln.get('rezeption_prioritaet', None)
    
    # --- Config-Info anzeigen ---
    with st.expander("📋 Geladene Konfiguration"):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Bereiche", len(bereiche_list))
        with col2:
            st.metric("Mitarbeiter", len(arzthelfer_list))
        with col3:
            praxis_name = config.get('meta', {}).get('praxis_name', 'Unbekannt')
            st.metric("Praxis", praxis_name)
    
    # --- 1. Bereiche konfigurieren ---
    st.header("1. Bereiche konfigurieren")
    selected_bereiche = st.multiselect("Bereiche auswählen", bereiche_list, default=bereiche_list)
    
    for bereich in selected_bereiche:
        with st.expander(bereich):
            shifts = {}
            defaults = default_shifts_map.get(bereich, {d: [] for d in days})
            for d in days:
                sel = st.multiselect(
                    f"Schichten am {d}", 
                    schifts, 
                    default=defaults.get(d, []), 
                    key=f"sh_{bereich}_{d}"
                )
                shifts[d] = sel
            
            helpers = st.multiselect(
                f"Arzthelferinnen für {bereich}", 
                arzthelfer_list,
                default=default_helpers_map.get(bereich, []), 
                key=f"h_{bereich}"
            )
            st.session_state.bereiche_cfg[bereich] = {'shifts': shifts, 'helpers': helpers}

    # --- 2. Arzthelferinnen konfigurieren ---
    st.header("2. Arzthelferinnen konfigurieren")
    selected_helpers = st.multiselect("Arzthelferinnen auswählen", arzthelfer_list, default=arzthelfer_list)
    
    for h in selected_helpers:
        with st.expander(h):
            max_h = st.number_input(
                f"Max. Stunden/Woche {h}", 
                0, 60,
                default_max_hours.get(h, 40), 
                key=f"mh_{h}"
            )
            
            times = {}
            h_defaults = default_helper_shifts_map.get(h, {d: [] for d in days})
            for d in days:
                sel = st.multiselect(
                    f"Einsatz am {d}", 
                    schifts, 
                    default=h_defaults.get(d, []), 
                    key=f"ts_{h}_{d}"
                )
                times[d] = sel
            
            areas = st.multiselect(
                f"Einsatzbereiche für {h}", 
                bereiche_list,
                default=default_areas_map.get(h, []), 
                key=f"a_{h}"
            )
            st.session_state.helpers_cfg[h] = {
                'max_hours': max_h, 
                'times': times, 
                'areas': areas
            }

    # --- 3. Dienstplan generieren ---
    st.header("3. Dienstplan generieren")
    if st.button("Plan erstellen"):
        plan = []
        used_helpers = {(d, s): set() for d in days for s in schifts}
        slots_order = [f"{d} {s}" for d in days for s in schifts]

        for d in days:
            for s in schifts:
                for bereich, cfg in st.session_state.bereiche_cfg.items():
                    if s in cfg['shifts'][d]:
                        candidates = [
                            h for h, hcfg in st.session_state.helpers_cfg.items()
                            if hcfg['max_hours'] > 0
                            and s in hcfg['times'][d]
                            and bereich in hcfg['areas']
                            and h in cfg['helpers']
                            and h not in used_helpers[(d, s)]
                        ]
                        
                        # Priorität für Rezeption (aus Config)
                        if bereich.startswith('Rezeption') and rezeption_prioritaet and rezeption_prioritaet in candidates:
                            chosen = rezeption_prioritaet
                        else:
                            random.shuffle(candidates)
                            chosen = candidates[0] if candidates else None

                        if chosen:
                            plan.append({'Slot': f"{d} {s}", 'Bereich': bereich, 'Helferin': chosen})
                            st.session_state.helpers_cfg[chosen]['max_hours'] -= 1
                            used_helpers[(d, s)].add(chosen)

        if plan:
            df_plan = pd.DataFrame(plan)
            df_pivot = (
                df_plan
                .pivot(index='Slot', columns='Bereich', values='Helferin')
                .reindex(index=slots_order, columns=selected_bereiche)
                .fillna('-')
            )
            
            # Plan in Session State speichern
            st.session_state.current_plan = plan
            st.session_state.current_pivot = df_pivot
            
            st.subheader("Wöchentlicher Dienstplan")
            st.table(df_pivot)
            
            # Excel Export Button
            st.subheader("Export")
            excel_data = create_excel_export(df_pivot, plan)
            st.download_button(
                label="📊 Als Excel-Datei herunterladen",
                data=excel_data,
                file_name=f"Dienstplan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Lädt den aktuellen Dienstplan als Excel-Datei herunter"
            )
            
            # Zusätzliche Export-Optionen
            col1, col2 = st.columns(2)
            with col1:
                csv_data = df_pivot.to_csv(index=True).encode('utf-8')
                st.download_button(
                    label="📄 Als CSV herunterladen",
                    data=csv_data,
                    file_name=f"Dienstplan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv"
                )
            
            with col2:
                # HTML Export für Drucken
                html_data = df_pivot.to_html(table_id="dienstplan")
                html_full = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <title>Dienstplan</title>
                    <style>
                        #dienstplan {{ border-collapse: collapse; width: 100%; }}
                        #dienstplan th, #dienstplan td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
                        #dienstplan th {{ background-color: #f2f2f2; font-weight: bold; }}
                        @media print {{ body {{ margin: 0; }} }}
                    </style>
                </head>
                <body>
                    <h1>Wöchentlicher Dienstplan</h1>
                    <p>Erstellt am: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}</p>
                    {html_data}
                </body>
                </html>
                """
                st.download_button(
                    label="🖨️ Als HTML (Drucken)",
                    data=html_full.encode('utf-8'),
                    file_name=f"Dienstplan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.html",
                    mime="text/html"
                )
        else:
            st.warning("Keine Zuweisungen möglich. Überprüfe deine Konfiguration.")

# --- Sidebar Aktionen ---
st.sidebar.header("Aktionen")

if st.session_state.config_loaded:
    if st.sidebar.button("🔄 Neue Config laden"):
        st.session_state.config_loaded = False
        st.session_state.config_data = None
        st.session_state.bereiche_cfg.clear()
        st.session_state.helpers_cfg.clear()
        st.session_state.current_plan = None
        st.session_state.current_pivot = None
        st.rerun()
    
    if st.sidebar.button("🗑️ Reset Konfiguration"):
        st.session_state.bereiche_cfg.clear()
        st.session_state.helpers_cfg.clear()
        st.session_state.current_plan = None
        st.session_state.current_pivot = None
        st.rerun()
    
    # --- Config-Info in Sidebar ---
    if st.session_state.config_data:
        st.sidebar.header("📋 Config-Info")
        config_meta = st.session_state.config_data.get('meta', {})
        st.sidebar.info(f"""
        **Praxis:** {config_meta.get('praxis_name', 'Unbekannt')}
        **Version:** {config_meta.get('version', 'Unbekannt')}
        **Bereiche:** {len(st.session_state.config_data['bereiche'])}
        **Mitarbeiter:** {len(st.session_state.config_data['mitarbeiter'])}
        """)

else:
    st.sidebar.info("🔧 Bitte Config-Datei laden")

# --- Sidebar Info ---
st.sidebar.header("Export-Formate")
st.sidebar.info("""
**Excel (.xlsx)**: Vollständig formatierte Tabelle mit separatem Raw-Data Sheet

**CSV (.csv)**: Einfaches Komma-getrenntes Format für weitere Bearbeitung

**HTML (.html)**: Druckfreundliches Format für Browser
""")

if st.session_state.current_pivot is not None:
    st.sidebar.success("✅ Aktueller Plan verfügbar")
else:
    st.sidebar.info("ℹ️ Noch kein Plan erstellt")

# --- Footer ---
st.sidebar.markdown("---")
st.sidebar.markdown("🔒 **Datenschutz:** Alle Daten bleiben lokal")
