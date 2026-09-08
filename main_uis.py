import os
from datetime import datetime
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# 1. CONFIGURACIÓN INICIAL
st.set_page_config(page_title="UIS | Portal de Notas", layout="wide")

# 2. DICCIONARIO DE MATERIAS Y PESOS RESTRUCTURADOS POR CURSO
MAPA_CURSOS = {
    "E1": "Cálculo II",
    "PE9": "Cálculo I",
    "PF1": "Álgebra Lineal",
    "PF3": "Álgebra Lineal",
}

PESOS = {
    # Álgebras (PF1, PF3): P1 (15%), P2 (25%), P3 (20%), P4 (20%), PQT (10%), TUTOR (10%)
    "PF1": {
        "P1": 0.15,
        "P2": 0.25,
        "P3": 0.20,
        "P4": 0.20,
        "PQT": 0.10,
        "ALEKS": 0.00,
        "TUTOR": 0.10,
    },
    "PF3": {
        "P1": 0.15,
        "P2": 0.25,
        "P3": 0.20,
        "P4": 0.20,
        "PQT": 0.10,
        "ALEKS": 0.00,
        "TUTOR": 0.10,
    },
    # Cálculo I (PE9): P1 (15%), P2 (20%), P3 (25%), P4 (20%), PQT (5%), ALEKS (10%), TUTOR (10%)
    "PE9": {
        "P1": 0.15,
        "P2": 0.20,
        "P3": 0.25,
        "P4": 0.20,
        "PQT": 0.05,
        "ALEKS": 0.10,
        "TUTOR": 0.10,
    },
    # Cálculo II (E1): P1 (20%), P2 (20%), P3 (25%), P4 (20%), PQT (5%), ALEKS (10%)
    "E1": {
        "P1": 0.20,
        "P2": 0.20,
        "P3": 0.25,
        "P4": 0.20,
        "PQT": 0.05,
        "ALEKS": 0.10,
        "TUTOR": 0.00,
    },
}

# 3. ESTILO CSS
st.markdown(
    """
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
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=60)
def load_data():
  file_path = "app_notas_uis.xlsx"
  try:
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    data = {}
    mod_time = "No disponible"
    try:
      ts = os.path.getmtime(file_path)
      mod_time = datetime.fromtimestamp(ts).strftime(
          "%-d de %B de %Y, %I:%M %p"
      )
    except Exception:
      pass

    for sheet in xls.sheet_names:
      df = xls.parse(sheet)
      df.columns = [str(c).strip().upper() for c in df.columns]

      for col in df.columns:
        if col.startswith("ESTUDIANTE"):
          df = df.rename(columns={col: "NOMBRE"})
          break

      if "COD" in df.columns:
        df["COD"] = df["COD"].astype(str).str.strip().str.split(".").str[0]

      for col in df.columns:
        if col not in ["NOMBRE", "COD", "NO"]:
          df[col] = pd.to_numeric(df[col], errors="coerce")

      data[sheet] = df
    return data, mod_time
  except Exception as e:
    st.error(f"Error cargando archivo: {e}")
    return None, None


def round_nota(val):
  if val is None or (isinstance(val, float) and pd.isna(val)):
    return 0.0
  return float(round(float(val) + 0.0000001, 1))


def celda_tiene_valor(row, col):
  val = row.get(col, None)
  if val is None:
    return False
  try:
    return not pd.isna(val)
  except Exception:
    return False


def obtener_pqt(row, todas_cols):
  """Obtiene el PQT de Excel o lo calcula desde Quices (Q), Trabajos (Tr) y Talleres (Ta)."""
  if celda_tiene_valor(row, "PQT"):
    return round_nota(row.get("PQT"))

  q_tr_ta_vals = []
  for col in todas_cols:
    c_upper = col.upper()
    is_q_tr_ta = (
        c_upper.startswith("Q")
        or c_upper.startswith("TR")
        or c_upper.startswith("TA")
    )
    if is_q_tr_ta and c_upper not in ["PQT", "QT"]:
      if celda_tiene_valor(row, col):
        q_tr_ta_vals.append(round_nota(row.get(col)))

  if q_tr_ta_vals:
    return float(round(sum(q_tr_ta_vals) / len(q_tr_ta_vals) + 0.0000001, 1))
  return 0.0


def obtener_columna_alternativa(row, nombres_posibles):
  """Busca coincidencias flexibles de nombres de columna (p. ej., ALEKS o TUTOR)."""
  for nombre in nombres_posibles:
    for key in row.keys():
      if nombre in str(key).upper() and celda_tiene_valor(row, key):
        return round_nota(row.get(key))
  return 0.0


# --- CARGA ---
dict_cursos, ultima_actualizacion = load_data()

# --- CABECERA PRINCIPAL ---
col_titulo, col_fecha = st.columns([3, 1])
with col_titulo:
  st.title("🎓 Portal de Notas — UIS")
with col_fecha:
  st.markdown(
      f"""
        <div style="margin-top:18px; text-align:right;">
            <span class="update-badge">🕒 Actualizado: {ultima_actualizacion}</span>
        </div>
    """,
      unsafe_allow_html=True,
  )

if st.button("🔄 Actualizar datos"):
  st.cache_data.clear()
  st.rerun()

if dict_cursos:
  col_input1, col_input2 = st.columns([1, 1])
  with col_input1:
    grupo_sel = st.selectbox(
        "Seleccione su Grupo",
        list(dict_cursos.keys()),
        format_func=lambda g: f"{g} — {MAPA_CURSOS.get(g, g)}",
    )
  with col_input2:
    cod_estudiante = st.text_input(
        "Ingrese su Código de Estudiante", placeholder="Ej: 22..."
    ).strip()

  consultar = st.button("Consultar mis Notas")

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

  err = st.session_state.get("error")
  if err == "not_found":
    st.warning(
        "Código no encontrado en este grupo. Verifica el grupo y tu código."
    )
  elif err == "no_col":
    st.error("Error: Columna COD no detectada en la hoja.")

  res = st.session_state.get("resultado")
  if res:
    row = res["row"]
    todas_cols = res["todas_cols"]
    grupo_res = res["grupo"]
    pesos = PESOS.get(
        grupo_res,
        {
            "P1": 0.20,
            "P2": 0.20,
            "P3": 0.20,
            "P4": 0.20,
            "PQT": 0.20,
            "ALEKS": 0.0,
            "TUTOR": 0.0,
        },
    )

    nombre = row.get("NOMBRE", "Estudiante")
    st.markdown(
        f'<p class="user-welcome">Bienvenid@, {nombre}</p>',
        unsafe_allow_html=True,
    )
    st.write(
        f"📖 **{MAPA_CURSOS.get(grupo_res, grupo_res)}** | Grupo: {grupo_res}"
    )

    # Extraer componentes
    p1 = round_nota(row.get("P1"))
    p2 = round_nota(row.get("P2"))
    p3 = round_nota(row.get("P3"))
    p4 = round_nota(row.get("P4"))
    pqt = obtener_pqt(row, todas_cols)
    aleks = obtener_columna_alternativa(row, ["ALEKS"])
    tutor = obtener_columna_alternativa(row, ["TUTOR", "TUTO", "TUTORIA"])

    # Nota total calculada dinámicamente con ponderación oficial
    total = round(
        p1 * pesos["P1"]
        + p2 * pesos["P2"]
        + p3 * pesos["P3"]
        + p4 * pesos["P4"]
        + pqt * pesos["PQT"]
        + aleks * pesos.get("ALEKS", 0.0)
        + tutor * pesos.get("TUTOR", 0.0)
        + 0.0000001,
        2,
    )

    # Semáforo
    if total >= 3.0:
      color_b = "#00FF41"
      status_txt = "¡FELICITACIONES, HAS APROBADO LA MATERIA! 🎉"
    elif total >= 2.5:
      color_b = "#F7B707"
      status_txt = "ADVERTENCIA ⚠️: No bajes la guardia, estás cerca"
    elif total >= 1.8:
      color_b = "#FF3131"
      status_txt = "RIESGO ALTO 🚨: Necesitas esforzarte al máximo"
    else:
      color_b = "#FF3131"
      status_txt = "SITUACIÓN CRÍTICA 😵: Habla con tu profesor"

    st.markdown(
        f"""
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
        """,
        unsafe_allow_html=True,
    )

    if total >= 3.0:
      import time

      components.html(
          f"""
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
            """,
          height=80,
      )

    segundo_corte_cerrado = celda_tiene_valor(row, "P4") and p4 > 0

    if segundo_corte_cerrado:
      if total >= 3.0:
        color_final = "#00FF41"
        mensaje_final = "🎉 ¡Felicitaciones! Aprobaste la materia."
      else:
        color_final = "#FF3131"
        mensaje_final = (
            "😔 Lo siento, no pasaste. Debes habilitar la materia."
        )

      st.markdown(
          f"""
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
            """,
          unsafe_allow_html=True,
      )
    else:
      st.write(f"Nota definitiva actual: **{total:.2f}** / 5.0")

    st.divider()

    # --- PESTAÑAS DE DETALLE Y SIMULADOR ---
    t1, t2, t3 = st.tabs([
        "📊 Detalle de Notas",
        "🎯 Aporte por Componente",
        "🔮 Simulador de Proyección",
    ])

    with t1:
      cards_principales = [
          (f"Parcial 1 ({int(pesos['P1']*100)}%)", f"{p1:.1f}"),
          (f"Parcial 2 ({int(pesos['P2']*100)}%)", f"{p2:.1f}"),
          (f"Parcial 3 ({int(pesos['P3']*100)}%)", f"{p3:.1f}"),
          (f"Parcial 4 ({int(pesos['P4']*100)}%)", f"{p4:.1f}"),
          (f"Prom. PQT ({int(pesos['PQT']*100)}%)", f"{pqt:.1f}"),
      ]
      if pesos.get("ALEKS", 0) > 0:
        cards_principales.append(
            (f"ALEKS ({int(pesos['ALEKS']*100)}%)", f"{aleks:.1f}")
        )
      if pesos.get("TUTOR", 0) > 0:
        cards_principales.append(
            (f"TUTOR ({int(pesos['TUTOR']*100)}%)", f"{tutor:.1f}")
        )

      cols_cards = st.columns(len(cards_principales))
      for idx, (label, val) in enumerate(cards_principales):
        cols_cards[idx].metric(label, val)

      # Módulos (Ma)
      ma_cols = [
          col
          for col in todas_cols
          if (col.upper().startswith("MA") or col.upper().startswith("MOD"))
          and celda_tiene_valor(row, col)
      ]
      if ma_cols:
        st.markdown("#### 📚 Módulos ALEKS (Ma)")
        cols_ma = st.columns(min(len(ma_cols), 8))
        for i, col_name in enumerate(ma_cols):
          with cols_ma[i % 8]:
            st.markdown(
                f"""
                                <div class="taller-card">
                                    <span class="taller-label">{col_name}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""",
                unsafe_allow_html=True,
            )

      # Quices
      q_cols = [
          col
          for col in todas_cols
          if col.upper().startswith("Q")
          and col.upper() not in ("QT", "PQT")
          and celda_tiene_valor(row, col)
      ]
      if q_cols:
        st.markdown("#### 🧩 Quices")
        cols_q = st.columns(min(len(q_cols), 8))
        for i, col_name in enumerate(q_cols):
          with cols_q[i % 8]:
            st.markdown(
                f"""
                                <div class="taller-card">
                                    <span class="taller-label">{col_name}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""",
                unsafe_allow_html=True,
            )

      # Talleres
      ta_cols = [
          col
          for col in todas_cols
          if col.upper().startswith("TA") and celda_tiene_valor(row, col)
      ]
      if ta_cols:
        st.markdown("#### 📝 Talleres")
        cols_ta = st.columns(min(len(ta_cols), 8))
        for i, col_name in enumerate(ta_cols):
          with cols_ta[i % 8]:
            st.markdown(
                f"""
                                <div class="taller-card">
                                    <span class="taller-label">{col_name}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""",
                unsafe_allow_html=True,
            )

      # Trabajos
      tr_cols = [
          col
          for col in todas_cols
          if col.upper().startswith("TR") and celda_tiene_valor(row, col)
      ]
      if tr_cols:
        st.markdown("#### 📁 Trabajos")
        cols_tr = st.columns(min(len(tr_cols), 8))
        for i, col_name in enumerate(tr_cols):
          with cols_tr[i % 8]:
            st.markdown(
                f"""
                                <div class="taller-card">
                                    <span class="taller-label">{col_name}</span>
                                    <span class="taller-value">{round_nota(row.get(col_name)):.1f}</span>
                                </div>""",
                unsafe_allow_html=True,
            )

    with t2:
      st.subheader("🎯 Aporte real de cada componente a la nota definitiva")
      componentes = {
          f"Parcial 1 ({int(pesos['P1']*100)}%)": p1 * pesos["P1"],
          f"Parcial 2 ({int(pesos['P2']*100)}%)": p2 * pesos["P2"],
          f"Parcial 3 ({int(pesos['P3']*100)}%)": p3 * pesos["P3"],
          f"Parcial 4 ({int(pesos['P4']*100)}%)": p4 * pesos["P4"],
          f"PQT ({int(pesos['PQT']*100)}%)": pqt * pesos["PQT"],
      }
      if pesos.get("ALEKS", 0) > 0:
        componentes[f"ALEKS ({int(pesos['ALEKS']*100)}%)"] = (
            aleks * pesos["ALEKS"]
        )
      if pesos.get("TUTOR", 0) > 0:
        componentes[f"TUTOR ({int(pesos['TUTOR']*100)}%)"] = (
            tutor * pesos["TUTOR"]
        )

      cols_d = st.columns(len(componentes))
      for i, (label, aporte) in enumerate(componentes.items()):
        cols_d[i].metric(label, f"{aporte:.3f}")
      st.info(
          f"Suma de aportes = **{sum(componentes.values()):.2f}** (nota"
          " definitiva actual)"
      )

    with t3:
      st.subheader("🔮 ¿Qué necesito para aprobar?")

      acum_fijo_sin_p4 = (
          p1 * pesos["P1"]
          + p2 * pesos["P2"]
          + p3 * pesos["P3"]
          + pqt * pesos["PQT"]
          + aleks * pesos.get("ALEKS", 0.0)
          + tutor * pesos.get("TUTOR", 0.0)
      )
      p4_necesario = (3.0 - acum_fijo_sin_p4) / pesos["P4"]

      st.markdown("##### Estado actual con componentes registrados:")
      if celda_tiene_valor(row, "P4") and p4 > 0:
        st.info(
            f"Ya tienes P4 registrado: **{p4:.1f}**. Tu nota final es"
            f" **{total:.2f}**."
        )
      elif p4_necesario <= 0:
        st.success("¡Ya aprobaste sin necesitar P4! 🎉")
      elif p4_necesario > 5.0:
        st.error(
            f"Con el acumulado actual, necesitarías **{p4_necesario:.2f}** en"
            " P4 para llegar a 3.0."
        )
        st.warning("Usa los controles de abajo para simular mejores notas. 👇")
      else:
        st.warning(
            f"Necesitas mínimo **{p4_necesario:.2f} / 5.0** en P4 para"
            " aprobar."
        )

      st.divider()
      st.markdown("##### 🎛️ Simula tus escenarios")

      num_sim_cols = (
          2
          + (1 if pesos.get("ALEKS", 0) > 0 else 0)
          + (1 if pesos.get("TUTOR", 0) > 0 else 0)
      )
      cols_sim = st.columns(num_sim_cols)

      with cols_sim[0]:
        p4_sim = st.slider(
            "Nota esperada P4",
            0.0,
            5.0,
            value=float(p4)
            if (celda_tiene_valor(row, "P4") and p4 > 0)
            else 3.0,
            step=0.1,
            key="slider_p4",
        )
      with cols_sim[1]:
        pqt_sim = st.slider(
            "Promedio PQT proyectado",
            0.0,
            5.0,
            value=float(pqt),
            step=0.1,
            key="slider_pqt",
        )

      curr_col = 2
      if pesos.get("ALEKS", 0) > 0:
        with cols_sim[curr_col]:
          aleks_sim = st.slider(
              "Nota ALEKS proyectada",
              0.0,
              5.0,
              value=float(aleks),
              step=0.1,
              key="slider_aleks",
          )
        curr_col += 1
      else:
        aleks_sim = 0.0

      if pesos.get("TUTOR", 0) > 0:
        with cols_sim[curr_col]:
          tutor_sim = st.slider(
              "Nota TUTOR proyectada",
              0.0,
              5.0,
              value=float(tutor),
              step=0.1,
              key="slider_tutor",
          )
      else:
        tutor_sim = 0.0

      total_sim = round(
          p1 * pesos["P1"]
          + p2 * pesos["P2"]
          + p3 * pesos["P3"]
          + p4_sim * pesos["P4"]
          + pqt_sim * pesos["PQT"]
          + aleks_sim * pesos.get("ALEKS", 0.0)
          + tutor_sim * pesos.get("TUTOR", 0.0)
          + 0.0000001,
          2,
      )

      color_sim = (
          "#00FF41"
          if total_sim >= 3.0
          else ("#F7B707" if total_sim >= 2.5 else "#FF3131")
      )
      st.markdown(
          f"""
                <div style="margin-top:10px; padding:16px; background-color:#161B22;
                            border-radius:12px; border:1px solid #30363D; text-align:center;">
                    <span style="color:#8b949e; font-size:0.9rem;">Nota definitiva proyectada</span><br>
                    <span style="color:{color_sim}; font-size:2.5rem; font-weight:bold;">{total_sim:.2f}</span>
                    <span style="color:#8b949e; font-size:1rem;"> / 5.0</span><br>
                    <span style="color:{color_sim}; font-weight:bold;">
                        {"✅ APRUEBA" if total_sim >= 3.0 else "❌ NO APRUEBA"}
                    </span>
                </div>
            """,
          unsafe_allow_html=True,
      )
