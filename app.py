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
days = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag']
schifts = ['Vormittag', 'Nachmittag']

# --- Config File Loader ---
def load_config(uploaded_file):
    """Lädt und validiert die Konfigurationsdatei"""
    try:
        config = json.load(uploaded_file)
        # Basis-Validierung
        required_keys = [
            'bereiche', 'mitarbeiter', 'bereich_schichten',
            'bereich_mitarbeiter', 'mitarbeiter_verfuegbarkeit',
            'mitarbeiter_bereiche', 'mitarbeiter_max_stunden',
            'standard_dienstplan'
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
    # Prüfe Standard-Dienstplan Struktur
    std = config.get('standard_dienstplan', {})
    # Jeder Slot-Schlüssel muss "<Wochentag> <Schicht>" sein
    for slot_key, assign_map in std.items():
        # Prüfe Format: "Montag Vormittag" etc.
        parts = slot_key.split()
        if len(parts) != 2 or parts[0] not in days or parts[1] not in schifts:
            errors.append(f"Ungültiges Standard-Slot-Format: '{slot_key}'")
        # Für jeden Bereich in assign_map sollte Bereich existieren und Helfer aus mitarbeiter
        for b, h in assign_map.items():
            if b not in config['bereiche']:
                errors.append(f"Standard-Dienstplan: Bereich '{b}' unbekannt im Slot '{slot_key}'")
            if h not in config['mitarbeiter']:
                errors.append(f"Standard-Dienstplan: Helfer '{h}' unbekannt im Slot '{slot_key}' für Bereich '{b}'")
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
def create_excel_export(df_pivot, df_demand, plan_data=None):
    """Erstellt eine Excel-Datei mit dem Dienstplan und markiert unbesetzte, aber geforderte Slots rot."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Haupttabelle schreiben
        df_pivot.to_excel(writer, sheet_name='Dienstplan', index=True)
        # Falls Raw-Daten verfügbar sind, als zweites Sheet hinzufügen
        if plan_data:
            df_raw = pd.DataFrame(plan_data)
            df_raw.to_excel(writer, sheet_name='Raw_Data', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Dienstplan']

        # Formate
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

        # Spaltenbreite
        worksheet.set_column('A:A', 15)
        for col in range(1, len(df_pivot.columns) + 1):
            worksheet.set_column(col, col, 12)

        # Header + Index
        for c, col in enumerate(df_pivot.columns, start=1):
            worksheet.write(0, c, col, header_fmt)
        for r, idx in enumerate(df_pivot.index, start=1):
            worksheet.write(r, 0, idx, header_fmt)

        # Nur geforderte & unbesetzte Slots rot
        rows, cols = df_pivot.shape
        for r in range(rows):
            slot = df_pivot.index[r]
            for c in range(cols):
                bereich = df_pivot.columns[c]
                val = df_pivot.iat[r, c]
                # Wenn Demand True und kein Helfer ("-"), dann rot, sonst normal
                fmt = red_fmt if (df_demand.at[slot, bereich] and val == '-') else cell_fmt
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
        "Konfigurationsdatei auswählen", type=['json'],
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
    config = st.session_state.config_data

    # Config-Daten extrahieren
    bereiche_list = config['bereiche']
    arzthelfer_list = config['mitarbeiter']
    default_shifts_map = config['bereich_schichten']
    default_helpers_map = config['bereich_mitarbeiter']
    default_helper_shifts_map = config['mitarbeiter_verfuegbarkeit']
    default_areas_map = config['mitarbeiter_bereiche']
    default_max_hours = config['mitarbeiter_max_stunden']
    standard_plan = config['standard_dienstplan']
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
    selected_bereiche = st.multiselect(
        "Bereiche auswählen", bereiche_list, default=bereiche_list
    )
    for bereich in selected_bereiche:
        with st.expander(bereich):
            shifts = {}
            defaults = default_shifts_map.get(bereich, {d: [] for d in days})
            for d in days:
                sel = st.multiselect(
                    f"Schichten am {d}", schifts,
                    default=defaults.get(d, []),
                    key=f"sh_{bereich}_{d}"
                )
                shifts[d] = sel
            helpers = st.multiselect(
                f"Arzthelferinnen für {bereich}", arzthelfer_list,
                default=default_helpers_map.get(bereich, []),
                key=f"h_{bereich}"
            )
            st.session_state.bereiche_cfg[bereich] = {
                'shifts': shifts,
                'helpers': helpers
            }

    # --- 2. Arzthelferinnen konfigurieren ---
    st.header("2. Arzthelferinnen konfigurieren")
    selected_helpers = st.multiselect(
        "Arzthelferinnen auswählen", arzthelfer_list, default=arzthelfer_list
    )
    for h in selected_helpers:
        with st.expander(h):
            max_h = st.number_input(
                f"Max. Stunden/Woche {h}", 0, 60,
                default_max_hours.get(h, 40),
                key=f"mh_{h}"
            )
            times = {}
            h_defaults = default_helper_shifts_map.get(h, {d: [] for d in days})
            for d in days:
                sel = st.multiselect(
                    f"Einsatz am {d}", schifts,
                    default=h_defaults.get(d, []),
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
        plan = []
        # Tracke, wie oft jede Helferin schon eingeteilt wurde (Stunden)
        used_helpers = {(d, s): set() for d in days for s in schifts}
        slots_order = [f"{d} {s}" for d in days for s in schifts]

        # **Demand-Matrix aufbauen** (welche Bereiche an welchen Slots Bedarf haben)
        df_demand = pd.DataFrame(False, index=slots_order, columns=selected_bereiche)
        for b, cfg in st.session_state.bereiche_cfg.items():
            for d in days:
                for s in cfg['shifts'][d]:
                    df_demand.at[f"{d} {s}", b] = True

        # 1) Zuweisungen gemäß Standard-Dienstplan (so strikt wie möglich)
        # Wir gehen über alle Slots; pro Slot über alle Bereiche
        assignments = {slot: {} for slot in slots_order}  # temporär: slot -> {bereich: helferin}

        # Zähle verfügbare Stunden pro Helferin
        helper_hours_left = {h: st.session_state.helpers_cfg[h]['max_hours']
                             for h in st.session_state.helpers_cfg}

        # Behalte Hilfsstruktur: Helfer-Verfügbarkeit & -Einsatzbereiche aus Session State
        helpers_cfg = st.session_state.helpers_cfg

        for slot in slots_order:
            # Zerlegen in Wochentag und Schicht
            d, s = slot.split()
            used_this_slot = set()  # Helfer, die schon in diesem Slot belegt wurden
            # 1a) Versuche, Standardzuweisung durchzuführen
            std_map = standard_plan.get(slot, {})
            for b in selected_bereiche:
                # Nur wenn Bereich für diesen Slot überhaupt Bedarf hat
                if df_demand.at[slot, b]:
                    std_h = std_map.get(b, None)
                    if std_h:
                        # Prüfe, ob diese Helferin noch frei und verfügbar ist
                        cfg_h = helpers_cfg.get(std_h, None)
                        if cfg_h:
                            # Voraussetzungen: noch Stunden verfügbar, Schicht wählbar, Bereich passt, Helfer in Bereichsliste
                            if (helper_hours_left[std_h] > 0
                                and s in cfg_h['times'][d]
                                and b in cfg_h['areas']
                                and std_h in st.session_state.bereiche_cfg[b]['helpers']
                                and std_h not in used_this_slot):
                                # Zuweisung fest übernehmen
                                assignments[slot][b] = std_h
                                used_this_slot.add(std_h)
                                helper_hours_left[std_h] -= 1

            # 1b) Standardzuweisungen hinzugefügt - markiere als verwendet
            for b, h in assignments[slot].items():
                used_helpers[(d, s)].add(h)

        # 2) Auffüllen der übrigen Bedarfslücken nach Knappheit
        for slot in slots_order:
            d, s = slot.split()
            used_this_slot = used_helpers[(d, s)]
            for b in selected_bereiche:
                if df_demand.at[slot, b]:
                    # Wenn noch keine Standardzuweisung existiert
                    if b not in assignments[slot]:
                        # Kandidatenliste: Helfer, die noch Stunden haben, verfügbar sind, im Bereich arbeiten, nicht schon in Slot, und in Bereichshelfer-Liste
                        candidates = [
                            h for h, cfg_h in helpers_cfg.items()
                            if helper_hours_left[h] > 0
                               and s in cfg_h['times'][d]
                               and b in cfg_h['areas']
                               and h in st.session_state.bereiche_cfg[b]['helpers']
                               and h not in used_this_slot
                        ]
                        # Sortiere nach Knappheit: Helfer mit wenigen Einsatzbereichen zuerst
                        candidates_sorted = sorted(
                            candidates,
                            key=lambda h: len(helpers_cfg[h]['areas'])
                        )

                        chosen = None
                        if candidates_sorted:
                            # Falls Bereich Rezeption und Priorität gegeben, z.B. "Eva"
                            if b.startswith('Rezeption') and rezeption_prioritaet in candidates_sorted:
                                chosen = rezeption_prioritaet
                            else:
                                chosen = candidates_sorted[0]

                        if chosen:
                            assignments[slot][b] = chosen
                            used_this_slot.add(chosen)
                            helper_hours_left[chosen] -= 1

            # Setze im used_helpers für diesen Slot die finalen Helfer ein
            for h in used_this_slot:
                used_helpers[(d, s)].add(h)

        # 3) Baue plan-Liste für DataFrame (Slot, Bereich, Helferin)
        for slot, asg_map in assignments.items():
            for b, h in asg_map.items():
                plan.append({'Slot': slot, 'Bereich': b, 'Helferin': h})

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

            # Anzeige: Nur wirklich geforderte & unbesetzte Slots rot markieren
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

            # Excel-Export
            st.subheader("Export")
            excel_data = create_excel_export(df_pivot, df_demand, plan)
            st.download_button(
                label="📊 Als Excel-Datei herunterladen",
                data=excel_data,
                file_name=f"Dienstplan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # CSV-/HTML-Export
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
                <!DOCTYPE html><html><head><title>Dienstplan</title>
                <style>#dienstplan{{border-collapse:collapse;width:100%;}}
                #dienstplan th,#dienstplan td{{border:1px solid #ddd;padding:8px;text-align:center;}}
                #dienstplan th{{background-color:#f2f2f2;font-weight:bold;}}
                @media print{{body{{margin:0;}}}}</style></head><body>
                <h1>Wöchentlicher Dienstplan</h1>
                <p>Erstellt am: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}</p>
                {html_data}
                </body></html>
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
        st.experimental_rerun()  # Neuer Start, damit Session State zurückgesetzt wird
    if st.sidebar.button("🗑️ Reset Konfiguration"):
        st.session_state.bereiche_cfg.clear()
        st.session_state.helpers_cfg.clear()
        st.session_state.current_plan = None
        st.session_state.current_pivot = None
        st.experimental_rerun()

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
**HTML (.html)**: Druckfreundliches Format
""")
if st.session_state.current_pivot is not None:
    st.sidebar.success("✅ Aktueller Plan verfügbar")
else:
    st.sidebar.info("ℹ️ Noch kein Plan erstellt")

# --- Footer ---
st.sidebar.markdown("---")
st.sidebar.markdown("🔒 **Datenschutz:** Alle Daten bleiben lokal")
