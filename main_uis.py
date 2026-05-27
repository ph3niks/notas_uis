import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from datetime import datetime
import os

# 1. CONFIGURACIÓN INICIAL
st.set_page_config(page_title="UIS | Portal de Notas", layout="wide")

# 2. DICCIONARIO DE MATERIAS Y PESOS POR GRUPO
MAPA_CURSOS = {
    "E1":  "Cálculo II",
    "PE9": "Cálculo I",
    "PF1": "Álgebra Lineal",
    "PF3": "Álgebra Lineal",
}

PESOS = {
    "E1":  {"P1": 0.20, "P2": 0.20, "P3": 0.25, "P4": 0.20, "PQT": 0.15},
    "PE9": {"P1": 0.20, "P2": 0.20, "P3": 0.25, "P4": 0.20, "PQT": 0.15},
    "PF1": {"P1": 0.15, "P2": 0.25, "P3": 0.20, "P4": 0.20, "PQT": 0.20},
    "PF3": {"P1": 0.15, "P2": 0.25, "P3": 0.20, "P4": 0.20, "PQT": 0.20},
}

# 3. ESTILO CSS
st.markdown("""
    <style>
    .main { background-color: #0E1117; color: #FFFFFF; }
    .user-welcome {
        color: #0755F2;
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        font-size: 1.8rem;
    }
    [data-testid="stMetric"] {
        background-color: #161B22 !important;
        border: 1px solid #30363D !important;
        border-radius: 12px !important;
        padding: 15px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.5) !important;
    }
    [data-testid="stMetricLabel"] p {
        color: #E0E0E0 !important;
        font-size: 0.9rem !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
    }
    [data-testid="stMetricValue"] { color: #00F2FF !important; }
    .taller-card {
        background-color: #1c2128;
        border: 1px solid #444c56;
        border-radius: 8px;
        padding: 10px;
        text-align: center;
        margin-bottom: 10px;
    }
    .taller-label {
        color: #8b949e;
        font-size: 0.75rem;
        font-weight: bold;
        display: block;
        margin-bottom: 3px;
    }
    .taller-value {
        color: #00F2FF;
        font-size: 1.1rem;
        font-weight: bold;
    }
    .update-badge {
        background-color: #161B22;
        border: 1px solid #30363D;
        border-radius: 8px;
        padding: 4px 12px;
        color: #8b949e;
        font-size: 0.78rem;
        display: inline-block;
    }
    .stButton>button {
        width: 100%;
        background-color: #00F2FF;
        color: #0E1117;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        padding: 0.5rem;
    }
    .stButton>button:hover { background-color: #00D1DB; color: #000; }
    </style>
    """, unsafe_allow_html=True)


@st.cache_data(ttl=60)
def load_data():
    file_path = "app_notas_uis.xlsx"
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        data = {}
        mod_time = "No disponible"
        try:
            ts = os.path.getmtime(file_path)
            mod_time = datetime.fromtimestamp(ts).strftime("%-d de %B de %Y, %I:%M %p")
        except Exception:
            pass

        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            df.columns = [str(c).strip().upper() for c in df.columns]

            # Normalizar columna de nombre
            for col in df.columns:
                if col.startswith("ESTUDIANTE"):
                    df = df.rename(columns={col: "NOMBRE"})
                    break

            # Normalizar columna COD
            if "COD" in df.columns:
                df["COD"] = df["COD"].astype(str).str.strip().str.split(".").str[0]

            # Convertir a numérico SIN rellenar NaN — así distinguimos vacío de cero
            for col in df.columns:
                if col not in ["NOMBRE", "COD", "NO"]:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            data[sheet] = df
        return data, mod_time
    except Exception as e:
        st.error(f"Error cargando archivo: {e}")
        return None, None


def round_nota(val):
    """Convierte a float redondeado; devuelve 0.0 para NaN/None."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    return float(round(float(val) + 0.0000001, 1))


def celda_tiene_valor(row, col):
    """True si la celda existe y NO es NaN (incluye el caso valor = 0.0)."""
    val = row.get(col, None)
    if val is None:
        return False
    try:
        return not pd.isna(val)
    except Exception:
        return False


# --- CARGA ---
dict_cursos, ultima_actualizacion = load_data()

# --- CABECERA PRINCIPAL ---
col_titulo, col_fecha = st.columns([3, 1])
with col_titulo:
    st.title("🎓 Portal de Notas — UIS")
with col_fecha:
    st.markdown(f"""
        <div style="margin-top:18px; text-align:right;">
            <span class="update-badge">🕒 Actualizado: {ultima_actualizacion}</span>
        </div>
    """, unsafe_allow_html=True)

if st.button("🔄 Actualizar datos"):
    st.cache_data.clear()
    st.rerun()

if dict_cursos:
    col_input1, col_input2 = st.columns([1, 1])
    with col_input1:
        grupo_sel = st.selectbox(
            "Seleccione su Grupo", list(dict_cursos.keys()),
            format_func=lambda g: f"{g} — {MAPA_CURSOS.get(g, g)}")
    with col_input2:
        cod_estudiante = st.text_input(
            "Ingrese su Código de Estudiante", placeholder="Ej: 22...").strip()

    consultar = st.button("Consultar mis Notas")

    # ── Guardar en session_state al consultar ──────────────────────────────────
    if cod_estudiante and consultar:
        df_actual = dict_cursos[grupo_sel]
        if "COD" in df_actual.columns:
            est = df_actual[df_actual["COD"] == cod_estudiante]
            if not est.empty:
                st.session_state["resultado"] = {
                    "row": est.iloc[0].to_dict(),
                    "todas_cols": list(df_actual.columns),
                    "grupo": grupo_sel,
                }
                st.session_state["error"] = None
            else:
                st.session_state["resultado"] = None
                st.session_state["error"] = "not_found"
        else:
            st.session_state["resultado"] = None
            st.session_state["error"] = "no_col"

    # ── Mensajes de error ──────────────────────────────────────────────────────
    err = st.session_state.get("error")
    if err == "not_found":
        st.warning("Código no encontrado en este grupo. Verifica el grupo y tu código.")
    elif err == "no_col":
        st.error("Error: Columna COD no detectada en la hoja.")

    # ── Mostrar resultados (persisten aunque se muevan los sliders) ────────────
    res = st.session_state.get("resultado")
    if res:
        row        = res["row"]
        todas_cols = res["todas_cols"]
        grupo_res  = res["grupo"]
        pesos      = PESOS[grupo_res]

        # Cabecera estudiante
        nombre = row.get("NOMBRE", "Estudiante")
        st.markdown(f'<p class="user-welcome">Bienvenid@, {nombre}</p>', unsafe_allow_html=True)
        st.write(f"📖 **{MAPA_CURSOS.get(grupo_res, grupo_res)}** | Grupo: {grupo_res}")

        # Notas
        p1  = round_nota(row.get("P1"))
        p2  = round_nota(row.get("P2"))
        p3  = round_nota(row.get("P3"))
        p4  = round_nota(row.get("P4"))
        pqt = round_nota(row.get("PQT"))

        total = round(
            p1  * pesos["P1"] +
            p2  * pesos["P2"] +
            p3  * pesos["P3"] +
            p4  * pesos["P4"] +
            pqt * pesos["PQT"] + 0.0000001, 2)

        # Semáforo
        if total >= 3.0:
            color_b    = "#00FF41"
            status_txt = "¡FELICITACIONES, HAS APROBADO LA MATERIA! 🎉"
        elif total >= 2.5:
            color_b    = "#F7B707"
            status_txt = "ADVERTENCIA ⚠️: No bajes la guardia, estás cerca"
        elif total >= 1.8:
            color_b    = "#FF3131"
            status_txt = "RIESGO ALTO 🚨: Necesitas esforzarte al máximo"
        else:
            color_b    = "#FF3131"
            status_txt = "SITUACIÓN CRÍTICA 😵: Habla con tu profesor"

        st.markdown(f"""
            <div style="margin-bottom:5px; display:flex; justify-content:space-between; align-items:flex-end;">
                <span style="color:{color_b}; font-weight:bold; font-size:1.1rem;">{status_txt}</span>
                <span style="color:#8b949e; font-size:0.8rem; font-weight:bold;">Meta mínima: 3.0</span>
            </div>
            <div style="width:100%; background-color:#333; border-radius:20px; height:24px; position:relative; overflow:hidden;">
                <div style="width:{min((total/5)*100,100)}%;
                            background-color:{color_b};
                            height:100%; border-radius:20px;
                            box-shadow:0 0 15px {color_b};
                            transition:width 1.5s ease-in-out;
                            position:absolute; z-index:1;"></div>
                <div style="position:absolute; left:60%; top:0; width:2px; height:100%;
                            background-color:rgba(255,255,255,0.4); z-index:2;"></div>
            </div>
        """, unsafe_allow_html=True)

        # Confeti
        if total >= 3.0:
            import time
            components.html(f"""
                <!-- cache-bust: {time.time()} -->
                <script>
                    (function() {{
                        var script = parent.document.createElement('script');
                        script.src = 'https://cdn.jsdelivr.net/npm/canvas-confetti@1.9.2/dist/confetti.browser.min.js';
                        script.onload = function() {{
                            var myConfetti = parent.confetti;
                            myConfetti({{
                                particleCount: 250,
                                spread: 160,
                                origin: {{ x: 0.5, y: 0.4 }},
                                colors: ['#00FF41', '#00F2FF', '#FFD700', '#FF69B4', '#ffffff']
                            }});
                            setTimeout(() => myConfetti({{
                                particleCount: 120,
                                angle: 60,
                                spread: 70,
                                origin: {{ x: 0, y: 0.6 }},
                                colors: ['#00FF41', '#00F2FF', '#FFD700']
                            }}), 300);
                            setTimeout(() => myConfetti({{
                                particleCount: 120,
                                angle: 120,
                                spread: 70,
                                origin: {{ x: 1, y: 0.6 }},
                                colors: ['#00FF41', '#00F2FF', '#FFD700']
                            }}), 600);
                        }};
                        parent.document.head.appendChild(script);
                    }})();
                </script>
                <div style="height:1px"></div>
            """, height=80)

        #estado de notas actual
        # Reemplaza la línea 286 por esto:
        segundo_corte_cerrado = celda_tiene_valor(row, "P4") and p4 > 0
        
        if segundo_corte_cerrado:
            if total >= 3.0:
                color_final = "#00FF41"
                mensaje_final = "🎉 ¡Felicitaciones! Aprobaste la materia."
            else:
                color_final = "#FF3131"
                mensaje_final = "😔 Lo siento, no pasaste. Debes habilitar la materia."
        
            st.markdown(f"""
                <div style="margin:20px 0; padding:24px; background-color:#161B22;
                            border:2px solid {color_final}; border-radius:16px; text-align:center;
                            box-shadow:0 0 20px {color_final}44;">
                    <div style="color:#8b949e; font-size:0.95rem; font-weight:bold;
                                text-transform:uppercase; margin-bottom:6px;">
                        NOTA DEFINITIVA FINAL
                    </div>
                    <div style="color:{color_final}; font-size:3.5rem; font-weight:bold; line-height:1.1;">
                        {total:.2f}
                        <span style="color:#8b949e; font-size:1.2rem;"> / 5.0</span>
                    </div>
                    <div style="color:{color_final}; font-size:1.2rem; font-weight:bold; margin-top:10px;">
                        {mensaje_final}
                    </div>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.write(f"Nota definitiva actual: **{total:.2f}** / 5.0")
        #st.write(f"Nota definitiva actual: **{total:.2f}** / 5.0")
        st.divider()

        # ── PESTAÑAS ──────────────────────────────────────────────────────────
        t1, t2, t3 = st.tabs(["📊 Detalle de Notas", "🎯 Aporte por Componente", "🔮 Simulador de Proyección"])

        with t1:
                tuto = round_nota(row.get("TUTO")) if celda_tiene_valor(row, "TUTO") else None

                # Porcentaje de tutorias según grupo
                TUTO_PCT = {"PE9": 8, "PF1": 10, "PF3": 10}
                tuto_pct = TUTO_PCT.get(grupo_res, 0)
                mostrar_tuto = grupo_res in TUTO_PCT

                tuto = round_nota(row.get("TUTO")) if mostrar_tuto else None

                num_cols = 6 if tuto is not None else 5
                c = st.columns(num_cols)
                c[0].metric(f"Parcial 1 ({int(pesos['P1']*100)}%)",       f"{p1:.1f}")
                c[1].metric(f"Parcial 2 ({int(pesos['P2']*100)}%)",       f"{p2:.1f}")
                c[2].metric(f"Parcial 3 ({int(pesos['P3']*100)}%)",       f"{p3:.1f}")
                c[3].metric(f"Parcial 4 ({int(pesos['P4']*100)}%)",       f"{p4:.1f}")
                c[4].metric(f"Prom. Talleres ({int(pesos['PQT']*100)}%)", f"{pqt:.1f}")
                if mostrar_tuto:
                    c[5].metric(f"Tutorías ({tuto_pct}%)", f"{tuto:.1f}")

                # Quices — mostrar si la celda tiene valor (incluyendo 0), omitir solo vacías
                q_cols = [col for col in todas_cols
                          if col.upper().startswith("Q")
                          and col.upper() not in ("QT",)
                          and celda_tiene_valor(row, col)]
                if q_cols:
                    st.markdown("#### 🧩 Quices")
                    cols_q = st.columns(min(len(q_cols), 8))
                    for i, col_name in enumerate(q_cols):
                        with cols_q[i % 8]:
                            st.markdown(f"""
                                <div class="taller-card">
                                    <span class="taller-label">{col_name}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""", unsafe_allow_html=True)

                # Talleres
                ta_cols = [col for col in todas_cols
                           if col.upper().startswith("TA")
                           and celda_tiene_valor(row, col)]
                if ta_cols:
                    st.markdown("#### 📝 Talleres")
                    cols_ta = st.columns(min(len(ta_cols), 8))
                    for i, col_name in enumerate(ta_cols):
                        with cols_ta[i % 8]:
                            st.markdown(f"""
                                <div class="taller-card">
                                    <span class="taller-label">T{col_name[2:]}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""", unsafe_allow_html=True)

                # Trabajos
                tr_cols = [col for col in todas_cols
                           if col.upper().startswith("TR")
                           and celda_tiene_valor(row, col)]
                if tr_cols:
                    st.markdown("#### 📁 Trabajos")
                    cols_tr = st.columns(min(len(tr_cols), 8))
                    for i, col_name in enumerate(tr_cols):
                        with cols_tr[i % 8]:
                            st.markdown(f"""
                                <div class="taller-card">
                                    <span class="taller-label">Tr{col_name[2:]}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""", unsafe_allow_html=True)

        with t2:
            st.subheader("🎯 Aporte real de cada componente a la nota definitiva")
            componentes = {
                f"Parcial 1 ({int(pesos['P1']*100)}%)": p1  * pesos["P1"],
                f"Parcial 2 ({int(pesos['P2']*100)}%)": p2  * pesos["P2"],
                f"Parcial 3 ({int(pesos['P3']*100)}%)": p3  * pesos["P3"],
                f"Parcial 4 ({int(pesos['P4']*100)}%)": p4  * pesos["P4"],
                f"Talleres ({int(pesos['PQT']*100)}%)":  pqt * pesos["PQT"],
            }
            cols_d = st.columns(len(componentes))
            for i, (label, aporte) in enumerate(componentes.items()):
                cols_d[i].metric(label, f"{aporte:.3f}")
            st.info(f"Suma de aportes = **{sum(componentes.values()):.2f}** (nota definitiva actual)")

        with t3:
            st.subheader("🔮 ¿Qué necesito para aprobar?")

            acum_sin_p4  = p1*pesos["P1"] + p2*pesos["P2"] + p3*pesos["P3"] + pqt*pesos["PQT"]
            p4_necesario = (3.0 - acum_sin_p4) / pesos["P4"]

            st.markdown("##### Con el promedio de talleres actual:")
            if celda_tiene_valor(row, "P4") and p4 > 0:
                st.info(f"Ya tienes P4 registrado: **{p4:.1f}**. Tu nota actual es **{total:.2f}**.")
            elif p4_necesario <= 0:
                st.success("¡Ya aprobaste sin necesitar P4! 🎉")
            elif p4_necesario > 5.0:
                st.error(f"Con el PQT actual ({pqt:.1f}), necesitarías **{p4_necesario:.2f}** en P4, superando el máximo de 5.0.")
                st.warning("Necesitas mejorar el promedio de talleres. Mira la simulación abajo 👇")
            else:
                st.warning(f"Necesitas mínimo **{p4_necesario:.2f} / 5.0** en P4 para aprobar (con PQT = {pqt:.1f}).")

            st.divider()
            st.markdown("##### 🎛️ Simula tus escenarios")

            col_s1, col_s2 = st.columns(2)
            with col_s1:
                p4_sim = st.slider(
                    "Nota esperada en P4", 0.0, 5.0,
                    value=float(p4) if (celda_tiene_valor(row, "P4") and p4 > 0) else 3.0,
                    step=0.1, key="slider_p4")
            with col_s2:
                pqt_sim = st.slider(
                    "Promedio de talleres proyectado+(TUTORIA si aplica) (PQT)", 0.0, 5.0,
                    value=float(pqt), step=0.1, key="slider_pqt")

            total_sim = round(
                p1*pesos["P1"] + p2*pesos["P2"] + p3*pesos["P3"] +
                p4_sim*pesos["P4"] + pqt_sim*pesos["PQT"] + 0.0000001, 2)

            color_sim = "#00FF41" if total_sim >= 3.0 else ("#F7B707" if total_sim >= 2.5 else "#FF3131")
            st.markdown(f"""
                <div style="margin-top:10px; padding:16px; background-color:#161B22;
                            border-radius:12px; border:1px solid #30363D; text-align:center;">
                    <span style="color:#8b949e; font-size:0.9rem;">Nota definitiva proyectada</span><br>
                    <span style="color:{color_sim}; font-size:2.5rem; font-weight:bold;">{total_sim:.2f}</span>
                    <span style="color:#8b949e; font-size:1rem;"> / 5.0</span><br>
                    <span style="color:{color_sim}; font-weight:bold;">
                        {"✅ APRUEBA" if total_sim >= 3.0 else "❌ NO APRUEBA"}
                    </span>
                </div>
            """, unsafe_allow_html=True)

            p4_min_sim = (3.0 - (p1*pesos["P1"] + p2*pesos["P2"] +
                                  p3*pesos["P3"] + pqt_sim*pesos["PQT"])) / pesos["P4"]
            if p4_min_sim <= 0:
                st.success(f"Con PQT = {pqt_sim:.1f}, ya apruebas sin importar P4. 🎉")
            elif p4_min_sim > 5.0:
                st.error(f"Con PQT = {pqt_sim:.1f}, no es posible aprobar ni con 5.0 en P4.")
            else:
                st.info(f"Con PQT = {pqt_sim:.1f}, necesitas mínimo **{p4_min_sim:.2f}** en P4 para aprobar.")
