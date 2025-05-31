import streamlit as st
import pandas as pd
import random
from io import BytesIO
import xlsxwriter
import json

# Hilfsfunktion für Excel-Spaltenbuchstaben (wird im Export genutzt)
def xlsx_colname(idx: int) -> str:
    """Wandelt 1-basierte Spaltennummer in Excel-Buchstaben um."""
    name = ""
    while idx > 0:
        idx, rem = divmod(idx-1, 26)
        name = chr(65 + rem) + name
    return name

# --- App-Konfiguration ---
st.set_page_config(page_title="Praxis-Dienstplanung", layout="wide")
st.title("Praxis-Dienstplanung: Dateninput & Dienstplan")

# --- Statische Daten (nicht sensibel) ---
days = ['Montag','Dienstag','Mittwoch','Donnerstag','Freitag']
schifts = ['Vormittag','Nachmittag']

# --- Config File Loader ---
def load_config(uploaded_file):
    """Lädt und validiert die Konfigurationsdatei (JSON)."""
    try:
        config = json.load(uploaded_file)
        
        # Basis-Validierung: Fehlende Schlüssel?
        required_keys = [
            'bereiche', 'mitarbeiter', 'standard_dienstplan',
            'bereich_schichten', 'bereich_mitarbeiter',
            'mitarbeiter_verfuegbarkeit',
            'mitarbeiter_bereiche', 'mitarbeiter_max_stunden',
            'spezial_regeln'
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
    """Erweiterte Validierung der Config-Struktur."""
    errors = []
    # Prüfe, ob alle Bereiche in 'bereich_schichten' gelistet sind
    for bereich in config['bereiche']:
        if bereich not in config['bereich_schichten']:
            errors.append(f"Bereich '{bereich}' fehlt in bereich_schichten")
    # Prüfe, ob alle Mitarbeiter in allen relevanten Mappings vorhanden sind
    for mitarbeiter in config['mitarbeiter']:
        if mitarbeiter not in config['mitarbeiter_verfuegbarkeit']:
            errors.append(f"Mitarbeiter '{mitarbeiter}' fehlt in mitarbeiter_verfuegbarkeit")
        if mitarbeiter not in config['mitarbeiter_bereiche']:
            errors.append(f"Mitarbeiter '{mitarbeiter}' fehlt in mitarbeiter_bereiche")
        if mitarbeiter not in config['mitarbeiter_max_stunden']:
            errors.append(f"Mitarbeiter '{mitarbeiter}' fehlt in mitarbeiter_max_stunden")
    return errors

# --- Excel Export Funktion ---
def create_excel_export(df_pivot, df_demand, plan_data=None):
    """
    Erstellt eine Excel-Datei mit dem Dienstplan und markiert nur unbesetzte, aber geforderte Slots rot.
    df_pivot: Pivot-Tabelle mit den tatsächlich zugewiesenen Helferinnen (Felder mit '-' wenn unbesetzt)
    df_demand: DataFrame (Bool) mit True dort, wo im Standard-/Bedarf eine Schicht vorgesehen war
    plan_data: Rohdaten (Liste von Dicts)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. Haupttabelle schreiben
        df_pivot.to_excel(writer, sheet_name='Dienstplan', index=True)
        # 2. Raw-Data (Plan-Daten) als zweites Sheet
        if plan_data:
            pd.DataFrame(plan_data).to_excel(writer, sheet_name='Raw_Data', index=False)
        
        # Workbook & Worksheet für Formatierung holen
        workbook  = writer.book
        worksheet = writer.sheets['Dienstplan']
        
        # Formate definieren
        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#D7E4BC', 'border': 1
        })
        cell_fmt   = workbook.add_format({
            'text_wrap': True, 'valign': 'top', 'border': 1
        })
        red_fmt    = workbook.add_format({
            'bg_color': '#FFC7CE', 'text_wrap': True, 'valign': 'top', 'border': 1
        })
        
        # Spaltenbreite anpassen
        worksheet.set_column('A:A', 15)  # Slot-Spalte
        for col in range(1, len(df_pivot.columns) + 1):
            worksheet.set_column(col, col, 12)
        
        # Kopfzeilen und Index formatieren
        for c, col_name in enumerate(df_pivot.columns, start=1):
            worksheet.write(0, c, col_name, header_fmt)
        for r, idx in enumerate(df_pivot.index, start=1):
            worksheet.write(r, 0, idx, header_fmt)
        
        # Datenzeilen mit Bedarfsprüfung einfärben
        rows, cols = df_pivot.shape
        for r in range(rows):
            slot = df_pivot.index[r]  # z.B. "Montag Vormittag"
            for c in range(cols):
                bereich = df_pivot.columns[c]  # z.B. "Rezeption 1"
                val = df_pivot.iat[r, c]
                # Wenn im Demand (Standard) True ist und 'val' = '-' (unbesetzt), dann rot, sonst normales Format
                if df_demand.at[slot, bereich] and val == '-':
                    fmt = red_fmt
                else:
                    fmt = cell_fmt
                worksheet.write(r+1, c+1, val, fmt)
    return output.getvalue()

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
        help="Lade deine Klinik_config2.json Datei hoch"
    )
    if uploaded_file is not None:
        config = load_config(uploaded_file)
        if config:
            validation_errors = validate_config(config)
            if validation_errors:
                st.error("❌ Validierungsfehler in der Config:")
                for e in validation_errors:
                    st.error(f"• {e}")
            else:
                st.session_state.config_data = config
                st.session_state.config_loaded = True
                st.success("✅ Konfiguration erfolgreich geladen!")

# --- HAUPT-APP (nur wenn Config geladen) ---
if st.session_state.config_loaded and st.session_state.config_data:
    cfg = st.session_state.config_data
    
    # Config-Daten extrahieren
    bereiche_list               = cfg['bereiche']
    arzthelfer_list             = cfg['mitarbeiter']
    standard_plan_map           = cfg['standard_dienstplan']
    default_shifts_map          = cfg['bereich_schichten']
    default_helpers_map         = cfg['bereich_mitarbeiter']
    default_helper_shifts_map   = cfg['mitarbeiter_verfuegbarkeit']
    default_areas_map           = cfg['mitarbeiter_bereiche']
    default_max_hours           = cfg['mitarbeiter_max_stunden']
    spezial_regeln              = cfg.get('spezial_regeln', {})
    rezeption_prioritaet        = spezial_regeln.get('rezeption_prioritaet', None)
    rezeption2_bedarf_vorher    = spezial_regeln.get('rezeption2_bedarf_erfuellt_vorher', False)
    
    # --- 1. Bereiche konfigurieren ---
    st.header("1. Bereiche konfigurieren")
    selected_bereiche = st.multiselect(
        "Bereiche auswählen", bereiche_list, default=bereiche_list
    )
    for bereich in selected_bereiche:
        with st.expander(bereich):
            shifts = {}
            defaults = default_shifts_map.get(bereich, {d: [] for d in days})
            # Für jeden Wochentag anzeigen, welche Schichten dieser Bereich benötigt
            for d in days:
                sel = st.multiselect(
                    f"Schichten am {d}", schifts,
                    default=defaults.get(d, []),
                    key=f"sh_{bereich}_{d}"
                )
                shifts[d] = sel
            # Auswahl, welche Helferinnen grundsätzlich in diesem Bereich arbeiten können
            helpers = st.multiselect(
                f"Arzthelferinnen für {bereich}", arzthelfer_list,
                default=default_helpers_map.get(bereich, []),
                key=f"h_{bereich}"
            )
            st.session_state.bereiche_cfg[bereich] = {'shifts': shifts, 'helpers': helpers}
    
    # --- 2. Arzthelferinnen konfigurieren ---
    st.header("2. Arzthelferinnen konfigurieren")
    selected_helpers = st.multiselect(
        "Arzthelferinnen auswählen", arzthelfer_list, default=arzthelfer_list
    )
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
            # Zeige für jeden Wochentag, wann die Helferin verfügbar ist
            for d in days:
                # Stelle sicher, dass Default-Werte nur 'Vormittag'/'Nachmittag' enthalten
                valid_default = [s for s in h_defaults.get(d, []) if s in schifts]
                sel = st.multiselect(
                    f"Einsatz am {d}", schifts,
                    default=valid_default,
                    key=f"ts_{h}_{d}"
                )
                times[d] = sel
            areas = st.multiselect(
                f"Einsatzbereiche für {h}", bereiche_list,
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
        plan = []  # Liste von Dicts: {'Slot': 'Montag Vormittag', 'Bereich': '...', 'Helferin': '...'}
        used_helpers = {(d, s): set() for d in days for s in schifts}
        slots_order = [f"{d} {s}" for d in days for s in schifts]
        
        # 3.1 Demand-Matrix (True/False), ob laut Standard/Bedarf überhaupt eine Schicht zu besetzen ist
        df_demand = pd.DataFrame(False, index=slots_order, columns=selected_bereiche)
        for b, cfg_b in st.session_state.bereiche_cfg.items():
            for d in days:
                for s in cfg_b['shifts'][d]:
                    slot = f"{d} {s}"
                    if b in selected_bereiche:
                        df_demand.at[slot, b] = True
        
        # 3.2 Zuweisung: Schritt 1 = STANDARDDIENSTPLAN übernehmen (sofern Helferin verfügbar & Bereich verlangt)
        # Wir bauen ein Mapping aus standard_plan_map: slot -> {Bereich: Helferin}
        # Hinweis: standard_plan_map enthält nur Slots, die der User definiert hat (z.B. "Montag Vormittag": {...})
        # Wir gehen Slot für Slot durch:
        for slot, belegungen in standard_plan_map.items():
            # Slot ist String "Montag Vormittag" o.ä.
            # Falls dieser Slot in unserer Wochenplanung enthalten ist (sonst ignorieren)
            if slot not in slots_order:
                continue
            # Wir nehmen die gewünschte Helferin aus dem Standard-Plan
            for bereich, helferin in belegungen.items():
                # Nur übernehmen, wenn:
                # a) dieser Bereich in selected_bereiche ist
                # b) s (Schicht) gehört momentan in cfg_b['shifts'][d]
                # c) Helferin ist verfügbar (max_hours > 0, Zeit passt, Bereich passt, in Bereichsmapping gelistet)
                
                # 1. Bereich wurde ausgewählt?
                if bereich not in selected_bereiche:
                    continue
                # 2. Zerlege Slot in Tag + Schicht
                parts = slot.split(" ")
                if len(parts) != 2:
                    continue
                d, s = parts[0], parts[1]
                cfg_b = st.session_state.bereiche_cfg.get(bereich, {})
                # 3. Prüfe, ob in cfg_b tatsächlich s eine geplante Schicht ist
                if s not in cfg_b['shifts'].get(d, []):
                    continue
                # 4. Kandidatin ist helferin
                hcfg = st.session_state.helpers_cfg.get(helferin)
                if not hcfg:
                    continue
                # 5. Helferin hat max_hours > 0
                if hcfg['max_hours'] <= 0:
                    continue
                # 6. Helferin hat s in ihren Zeiten an Tag d
                if s not in hcfg['times'].get(d, []):
                    continue
                # 7. Helferin darf in diesem Bereich arbeiten
                if bereich not in hcfg['areas']:
                    continue
                # 8. Helferin ist in Bereichsmapping für diesen Bereich gelistet
                if helferin not in cfg_b['helpers']:
                    continue
                # 9. Noch nicht eingesetzt in diesem Slot (d, s)
                if helferin in used_helpers[(d, s)]:
                    continue
                
                # Alle Prüfungen bestanden => Standardzuweisung übernehmen
                plan.append({'Slot': slot, 'Bereich': bereich, 'Helferin': helferin})
                st.session_state.helpers_cfg[helferin]['max_hours'] -= 1
                used_helpers[(d, s)].add(helferin)
        
        # 3.3 Zuweisung: Schritt 2 = Fehlende Schichten (demand=True, aber noch kein Eintrag) nach Knappheit füllen
        # Knappheitsbegriff: Anzahl möglicher Helferinnen pro Bereich (in diesem Slot) 
        # Je weniger Helferinnen zur Verfügung stehen, desto höher die Priorität, zuerst diesen Bereich zu besetzen.
        # Wir bilden für jeden (Slot, Bereich) die Liste aller Kandidaten und sortieren Bereiche nach Länge dieser Liste.
        
        # Für jeden Slot:
        for d in days:
            for s in schifts:
                slot = f"{d} {s}"
                # Berechne für alle Bereiche, für die df_demand[slot, Bereich] == True und noch keine Helferin zugewiesen wurde
                # Zunächst: Welche Bereiche benötigen an diesem Slot überhaupt noch eine Zuweisung?
                missing_bereiche = []
                for bereich, cfg_b in st.session_state.bereiche_cfg.items():
                    if bereich not in selected_bereiche:
                        continue
                    # Bedarf an diesem Slot?
                    if s not in cfg_b['shifts'].get(d, []):
                        continue
                    # War bereits belegt durch Standardplan?
                    # Wir prüfen, ob plan schon eine Eintragung für (slot, bereich) enthält
                    found = any((p['Slot']==slot and p['Bereich']==bereich) for p in plan)
                    if found:
                        continue
                    # Bedarf also noch offen
                    missing_bereiche.append(bereich)
                
                # Für jeden offen gebliebenen Bereich berechne Kandidaten-Liste
                bereich_kandidaten = []
                for bereich in missing_bereiche:
                    cfg_b = st.session_state.bereiche_cfg[bereich]
                    candidates = []
                    for h, hcfg in st.session_state.helpers_cfg.items():
                        # alle bisherigen Prüfungen nochmal:
                        if hcfg['max_hours'] <= 0:
                            continue
                        if s not in hcfg['times'].get(d, []):
                            continue
                        if bereich not in hcfg['areas']:
                            continue
                        if h not in cfg_b['helpers']:
                            continue
                        if h in used_helpers[(d, s)]:
                            continue
                        candidates.append(h)
                    # Speziell: Rezeption2-Bedarfsvorbedingung
                    if bereich.startswith('Rezeption 2') and rezeption2_bedarf_vorher:
                        # erst, wenn alle anderen Bedarfe in diesem (d, s) bereits erfüllt sind
                        # also: solange missing_bereiche neben 'Rezeption 2' noch andere enthalten sind, 
                        # überspringen wir 'Rezeption 2'
                        other_missing = [b for b in missing_bereiche if b.startswith('Rezeption')==False and b != 'Rezeption 2']
                        if other_missing:
                            # Wir setzen Kandidaten-Liste leer; 'Rezeption 2' wird jetzt noch nicht gefüllt
                            candidates = []
                    bereich_kandidaten.append((bereich, candidates))
                
                # Sortiere missing_bereiche nach Länge der Kandidatenliste (aufsteigend: Knappheit zuerst)
                bereich_kandidaten.sort(key=lambda bc: len(bc[1]))
                
                # Nun belege nacheinander nach dieser Reihenfolge
                for bereich, candidates in bereich_kandidaten:
                    if not candidates:
                        # Keine Kandidaten => kann nicht besetzen, wird später als unbesetzt ('-') gelten
                        continue
                    # Wähle die Helferin mit der geringsten max_hours aktuell ( weiterführende Knappheit innerhalb Helfer)
                    # Dadurch verwenden wir keine random-Zuweisung
                    # Sortiere Kandidaten nach aufsteigender verbleibender Stundenanzahl
                    candidates_sorted = sorted(candidates, key=lambda h: st.session_state.helpers_cfg[h]['max_hours'])
                    chosen = candidates_sorted[0]
                    # Falls Bereich Rezeption und 'rezeption_prioritaet' greift:
                    if bereich.startswith('Rezeption') and rezeption_prioritaet in candidates_sorted:
                        chosen = rezeption_prioritaet
                    
                    # Zuordnung übernehmen
                    plan.append({'Slot': slot, 'Bereich': bereich, 'Helferin': chosen})
                    st.session_state.helpers_cfg[chosen]['max_hours'] -= 1
                    used_helpers[(d, s)].add(chosen)
        
        # 3.4 Erstelle Pivot-Tabelle für Anzeige (leere Felder mit '-' füllen)
        if plan:
            df_plan = pd.DataFrame(plan)
            df_pivot = (
                df_plan
                .pivot(index='Slot', columns='Bereich', values='Helferin')
                .reindex(index=slots_order, columns=selected_bereiche)
                .fillna('-')
            )
            # Speichere im State
            st.session_state.current_plan  = plan
            st.session_state.current_pivot = df_pivot
            
            # 3.5 Anzeige im Streamlit: Nur wirklich nach Bedarf unbesetzte Slots rot markieren
            def highlight_unfilled(data):
                styles = pd.DataFrame("", index=data.index, columns=data.columns)
                for slot in data.index:
                    for b in data.columns:
                        if df_demand.at[slot, b] and data.at[slot, b] == '-':
                            styles.at[slot, b] = 'background-color: #FFC7CE'
                return styles
            
            st.subheader("Wöchentlicher Dienstplan")
            st.dataframe(
                df_pivot.style.apply(highlight_unfilled, axis=None),
                use_container_width=True
            )
            
            # 3.6 Excel-Export
            st.subheader("Export")
            excel_data = create_excel_export(df_pivot, df_demand, plan)
            st.download_button(
                label="📊 Als Excel-Datei herunterladen",
                data=excel_data,
                file_name=f"Dienstplan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # CSV- und HTML-Download
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
        st.session_state.config_loaded  = False
        st.session_state.config_data    = None
        st.session_state.bereiche_cfg.clear()
        st.session_state.helpers_cfg.clear()
        st.session_state.current_plan  = None
        st.session_state.current_pivot = None
    if st.sidebar.button("🗑️ Reset Konfiguration"):
        st.session_state.bereiche_cfg.clear()
        st.session_state.helpers_cfg.clear()
        st.session_state.current_plan  = None
        st.session_state.current_pivot = None
    if st.session_state.config_data:
        st.sidebar.header("📋 Config-Info")
        meta = st.session_state.config_data.get('meta', {})
        st.sidebar.info(f"""
        **Praxis:** {meta.get('praxis_name','Unbekannt')}
        **Version:** {meta.get('version','Unbekannt')}
        **Bereiche:** {len(st.session_state.config_data['bereiche'])}
        **Mitarbeiter:** {len(st.session_state.config_data['mitarbeiter'])}
        """)
else:
    st.sidebar.info("🔧 Bitte Config-Datei laden")

# --- Sidebar Info ---
st.sidebar.header("Export-Formate")
st.sidebar.info("""
**Excel (.xlsx)**: Vollständig formatierte Tabelle mit separatem Raw-Data Sheet  
**CSV (.csv)**: Einfaches Komma-getrenntes Format  
**HTML (.html)**: Druckfreundliches Format für Browser
""")
if st.session_state.current_pivot is not None:
    st.sidebar.success("✅ Aktueller Plan verfügbar")
else:
    st.sidebar.info("ℹ️ Noch kein Plan erstellt")

# --- Footer ---
st.sidebar.markdown("---")
st.sidebar.markdown("🔒 **Datenschutz:** Alle Daten bleiben lokal")
