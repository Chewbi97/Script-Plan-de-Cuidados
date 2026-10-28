"""
=============================================================
  BOT - CARACTERIZACIÓN SI-APS RES 212 VIGENCIA 2026
=============================================================
Requisitos:
  pip install selenium openpyxl webdriver-manager

Estructura Excel (hoja "Caracterizacion"):
  FECHA | MICROTERRITORIO | NOMBRE | (vacía) | TIPO |
  IDENTIFICACION | FEHA DE NACIMIENTO | TELEFONO | SEXO |
  EDAD | ... | REGIMEN | EAPB

Agrupación: misma familia = mismo TELEFONO
Sin teléfono/"-"/"NO TIENE" → familia unipersonal independiente
=============================================================
"""

import time, re, openpyxl
from datetime import datetime, date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ─────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────
ARCHIVO_EXCEL   = "base_datos.xlsx"
HOJA_NOMBRE     = "Caracterizacion"
URL_FORMULARIO  = "https://ee.kobotoolbox.org/x/iK2YkkfP"


ID_EQUIPO       = "7"
PRESTADOR       = "E.S.E Hospital Emiro Quintero Cañizares"
DOC_MARBEY      = "1004967093"
PERFIL_MARBEY   = "Auxiliar de enfermería"
CODIGO_FICHA    = "7"

DEPARTAMENTO    = "Norte de Santander"
UNIDAD_ZONAL    = "E.S.E Hospital Emiro Quintero Cañizares"
MUNICIPIO       = "Ocaña"

TIMEOUT_RESET   = 45
PAUSA           = 0.5

# ─────────────────────────────────────────────
#  PREFIJOS DE DATA-REF (del HTML real)
# ─────────────────────────────────────────────
BASE = "/aFMwUEBXQ7pGHdGqVEvLfk"
G1   = f"{BASE}/group_au6ls20"   # Sección 1 política
G2   = f"{BASE}/group_kb9cg47"   # Sección 2 info general
G3   = f"{BASE}/group_up21j08"   # Sección 3 entorno/vivienda
G4   = f"{BASE}/group_zx2jz62"   # Sección 4 familia
G5   = f"{BASE}/group_vt8bt74"   # Sección 5 integrantes (repeat)

# ─────────────────────────────────────────────
#  MAPEOS
# ─────────────────────────────────────────────
TIPOS_DOC = {
    "AS": "as__adulto_sin_identificaci_n",
    "CC": "cc__c_dula_de_ciudadan_a",
    "CD": "cd__carn__diplom_tico",
    "CE": "ce__c_dula_de_extranjer_a",
    "MS": "ms__menor_sin_identificaci_n",
    "NV": "nv__certificado_de_nacido_vivo",
    "PE": "pe__permiso_especial_de_permanencia",
    "PT": "pt__permiso_de_protecci_n_temporal",
    "RC": "rc__registro_civil",
    "TI": "ti__tarjeta_de_identidad",
}
REGIMEN_VAL = {
    "SUB": "subsdiado", "SUBSIDIADO": "subsdiado",
    "CON": "contributivo", "CONTRIBUTIVO": "contributivo",
    "ESPECIAL": "especial", "EXCEPCION": "excepci_n",
    "NO AFILIADO": "no_afiliado",
}

# ─────────────────────────────────────────────
#  UTILIDADES
# ─────────────────────────────────────────────
def limpiar(v):
    return re.sub(r"[.,\s]", "", str(v)).strip()

def tel_valido(t):
    s = str(t).strip().upper()
    return s not in ("", "-", "NO TIENE", "NONE", "NAN", "NONE")

def edad_calc(fn):
    hoy = date.today()
    if isinstance(fn, datetime): fn = fn.date()
    return hoy.year - fn.year - ((hoy.month, hoy.day) < (fn.month, fn.day))

def edu(edad):
    if edad < 6:  return "preescolar"
    if edad < 18: return "b_sica_primaria"
    if edad < 30: return "b_sica_secundaria"
    if edad < 60: return "media_acad_mica_o_cl_sica"
    return "b_sica_primaria"

def peso_talla(edad):
    if edad < 1:  return 8, 65
    if edad < 6:  return 18, 110
    if edad < 13: return 32, 140
    if edad < 18: return 55, 160
    if edad < 60: return 68, 165
    return 62, 158

def imc_val(peso, talla):
    imc = peso / ((talla/100)**2)
    if imc < 18.5: return "riesgo_de_delgadez"
    if imc < 25:   return "peso_adecuado_para_la_talla_o_imc_adecua"
    if imc < 30:   return "sobrepeso"
    return "obesidad"

def dividir_nombre(n):
    p = n.strip().split()
    if len(p) == 1: return p[0], "", "", ""
    if len(p) == 2: return p[0], "", p[1], p[1]   # repite apellido si solo hay 1
    if len(p) == 3: return p[0], "", p[1], p[2]
    return p[0], p[1], p[2], " ".join(p[3:])

def rol(pos, edad, sexo):
    if pos == 0: return "responsable_econ_mico_de_la_familia"
    if edad < 18: return "hijo_a"
    if sexo == "FEMENINO": return "c_nyuge_o_compa_ero_a"
    return "otro"

def tipo_familia(n):
    return "nuclear_biparental" if n > 1 else "nuclear_monoparenta"

def fecha_str(f):
    if isinstance(f, datetime): return f.strftime("%Y-%m-%d")
    if isinstance(f, date):     return f.strftime("%Y-%m-%d")
    return str(f)[:10]

# ─────────────────────────────────────────────
#  LEER EXCEL
# ─────────────────────────────────────────────
def leer_familias(archivo, hoja):
    wb = openpyxl.load_workbook(archivo, read_only=True)
    ws = wb[hoja]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c).strip().upper() if c else "" for c in rows[0]]
    def idx(n):
        for i,h in enumerate(hdr):
            if n in h: return i
        raise ValueError(f"No encontré columna '{n}'")
    iF=idx("FECHA"); iM=idx("MICROTERRITORIO"); iN=idx("NOMBRE")
    iT=idx("TIPO"); iI=idx("IDENTIFICACION"); iNac=idx("NACIMIENTO")
    iTel=idx("TELEFONO"); iS=idx("SEXO"); iE=idx("EDAD")
    iR=idx("REGIMEN"); iEapb=idx("EAPB")

    personas=[]
    for row in rows[1:]:
        if not row[iN]: continue
        tel_raw = row[iTel]
        tel = str(tel_raw).strip() if tel_valido(tel_raw) else None
        fn = row[iNac]
        fn_date = fn.date() if isinstance(fn,datetime) else (fn if isinstance(fn,date) else None)
        edad_r = row[iE]
        edad = int(edad_r) if edad_r else (edad_calc(fn_date) if fn_date else 30)
        personas.append({
            "fecha": row[iF], "micro": str(row[iM]).strip() if row[iM] else "",
            "nombre": str(row[iN]).strip(),
            "tipo": str(row[iT]).strip().upper() if row[iT] else "CC",
            "id": limpiar(row[iI]),
            "fecha_nac": fn_date, "telefono": tel,
            "tel_raw": str(tel_raw).strip() if tel_raw else "",
            "sexo": str(row[iS]).strip().upper() if row[iS] else "MASCULINO",
            "edad": edad,
            "regimen": str(row[iR]).strip().upper() if row[iR] else "SUB",
            "eapb": str(row[iEapb]).strip() if row[iEapb] else "",
        })

    familias = []
    tel_idx = {}
    for p in personas:
        tel = p["telefono"]
        if tel and tel in tel_idx:
            familias[tel_idx[tel]].append(p)
        elif tel:
            tel_idx[tel] = len(familias)
            familias.append([p])
        else:
            familias.append([p])
    return familias

# ─────────────────────────────────────────────
#  HELPERS SELENIUM
# ─────────────────────────────────────────────
def W(driver, t=12):
    return WebDriverWait(driver, t)

class CampoError(Exception):
    """Error al llenar un campo específico del formulario.
    Al lanzarse, detiene todo el proceso (no sigue con la siguiente familia)."""
    pass

def set_radio(driver, data_ref, value, context=None, intentos=3):
    """Selecciona un radio usando data-ref + value exactos.
    Reintenta y VERIFICA que el click realmente marcó la opción
    (Enketo a veces tarda en procesar el evento)."""
    base = context or driver
    ultimo_error = None
    for intento in range(1, intentos + 1):
        try:
            el = base.find_element(By.CSS_SELECTOR,
                f'input[type="radio"][data-ref="{data_ref}"][value="{value}"]')
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.2)
            # Click "real" simulando interacción de usuario (no solo .click() via JS)
            driver.execute_script("""
                const el = arguments[0];
                el.checked = true;
                el.dispatchEvent(new Event('click', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
            """, el)
            time.sleep(PAUSA)
            # Verificación: ¿quedó realmente marcado?
            if el.is_selected():
                return True
            # Si no quedó marcado, reintenta con click normal de Selenium
            time.sleep(0.3)
            el.click()
            time.sleep(PAUSA)
            if el.is_selected():
                return True
            ultimo_error = "el campo no quedó marcado tras el click"
        except Exception as e:
            ultimo_error = str(e)
        time.sleep(0.4)
    raise CampoError(f"radio {data_ref}={value} (tras {intentos} intentos): {ultimo_error}")

def set_checkbox(driver, data_ref, value, context=None, intentos=3):
    """Marca un checkbox usando data-ref + value exactos.
    Reintenta y VERIFICA que quedó marcado."""
    base = context or driver
    ultimo_error = None
    for intento in range(1, intentos + 1):
        try:
            el = base.find_element(By.CSS_SELECTOR,
                f'input[type="checkbox"][data-ref="{data_ref}"][value="{value}"]')
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.2)
            if el.is_selected():
                return True
            driver.execute_script("""
                const el = arguments[0];
                el.checked = true;
                el.dispatchEvent(new Event('click', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
            """, el)
            time.sleep(PAUSA)
            if el.is_selected():
                return True
            time.sleep(0.3)
            el.click()
            time.sleep(PAUSA)
            if el.is_selected():
                return True
            ultimo_error = "el checkbox no quedó marcado tras el click"
        except Exception as e:
            ultimo_error = str(e)
        time.sleep(0.4)
    raise CampoError(f"checkbox {data_ref}={value} (tras {intentos} intentos): {ultimo_error}")

def set_text(driver, data_ref, value, context=None):
    """Escribe en un input/text usando data-ref."""
    base = context or driver
    try:
        el = base.find_element(By.CSS_SELECTOR,
            f'input[type="text"][data-ref="{data_ref}"], '
            f'textarea[data-ref="{data_ref}"]')
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.15)
        el.clear()
        el.send_keys(str(value))
        time.sleep(PAUSA)
        return True
    except Exception as e:
        raise CampoError(f"text {data_ref}={value}: {e}")

def set_date(driver, data_ref, value, context=None):
    """Escribe en el input visible de fecha (el widget, no el hidden)."""
    base = context or driver
    try:
        # El widget date tiene un input visible de clase 'input-small'
        # que es hermano del input hidden con data-ref
        parent = base.find_element(By.CSS_SELECTOR,
            f'input[data-ref="{data_ref}"]').find_element(By.XPATH, "./..")
        inp = parent.find_element(By.CSS_SELECTOR, "input.input-small")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
        time.sleep(0.15)
        inp.clear()
        inp.send_keys(str(value))
        time.sleep(PAUSA)
        # Dispara blur para que Enketo procese la fecha
        driver.execute_script("arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", inp)
        time.sleep(0.3)
        return True
    except Exception as e:
        raise CampoError(f"date {data_ref}={value}: {e}")

def set_number(driver, data_ref, value, context=None):
    """Escribe en input type=number usando data-ref."""
    base = context or driver
    try:
        el = base.find_element(By.CSS_SELECTOR,
            f'input[type="number"][data-ref="{data_ref}"]')
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.15)
        el.clear()
        el.send_keys(str(value))
        time.sleep(PAUSA)
        return True
    except Exception as e:
        raise CampoError(f"number {data_ref}={value}: {e}")

def contar_repeats(driver):
    """Cuenta cuántos or-repeat de integrantes hay actualmente en el DOM."""
    return len(driver.find_elements(By.CSS_SELECTOR, f'section.or-repeat[name="{G5}"]'))

def get_repeat(driver, idx, timeout=8):
    """Retorna el elemento or-repeat con ese índice (0-based).
    Espera activamente a que exista por si Kobo aún no terminó de renderizarlo."""
    try:
        W(driver, timeout).until(lambda d: contar_repeats(d) > idx)
    except Exception:
        pass  # si no aparece, el error de abajo lo reporta con el conteo real
    repeats = driver.find_elements(By.CSS_SELECTOR,
        f'section.or-repeat[name="{G5}"]')
    if idx < len(repeats):
        return repeats[idx]
    raise CampoError(f"No se encontró el or-repeat #{idx} (solo hay {len(repeats)})")

def cerrar_popups_vex(driver):
    """Cierra cualquier popup/modal tipo 'vex' (alertas, confirmaciones)
    que pueda estar flotando y bloqueando clicks."""
    try:
        # Botones típicos de cierre/confirmación en popups vex de Kobo/Enketo
        selectores = [
            "div.vex button.vex-dialog-button-primary",
            "div.vex button.vex-dialog-button-secondary",
            "div.vex-overlay",
            "div.vex .vex-dialog-button",
        ]
        cerrado = False
        for sel in selectores:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                try:
                    driver.execute_script("arguments[0].click();", el)
                    cerrado = True
                except Exception:
                    pass
        # Si sigue habiendo overlay vex, lo removemos del DOM directamente
        driver.execute_script("""
            document.querySelectorAll('.vex, .vex-overlay').forEach(e => e.remove());
        """)
        if cerrado:
            time.sleep(0.5)
        return cerrado
    except Exception:
        return False

def add_repeat(driver, intentos=4):
    """Click en el botón add-repeat-btn para agregar integrante.
    Cierra popups vex que puedan estar bloqueando el click, y VERIFICA
    que el número de repeats realmente aumentó antes de darse por exitoso."""
    ultimo_error = None
    n_antes = contar_repeats(driver)
    for intento in range(1, intentos + 1):
        try:
            cerrar_popups_vex(driver)
            btn = W(driver,10).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, f'div.or-repeat-info[data-name="{G5}"] button.add-repeat-btn')
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.3)
            cerrar_popups_vex(driver)
            try:
                btn.click()
            except Exception:
                # Si el click normal es interceptado, forzamos vía JS
                driver.execute_script("arguments[0].click();", btn)
            # Espera activamente a que el repeat realmente aparezca
            try:
                W(driver, 6).until(lambda d: contar_repeats(d) > n_antes)
                time.sleep(0.5)
                return True
            except Exception:
                ultimo_error = (f"el click no creó un nuevo repeat "
                                 f"(antes={n_antes}, después={contar_repeats(driver)})")
        except Exception as e:
            ultimo_error = str(e)
        cerrar_popups_vex(driver)
        time.sleep(0.7)
    raise CampoError(f"add_repeat (tras {intentos} intentos): {ultimo_error}")

def descartar_borrador(driver, t=5):
    try:
        btn = W(driver,t).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(normalize-space(),'DESCARTAR')]")))
        btn.click(); time.sleep(2)
        print("  OK borrador descartado")
    except: pass

def esperar_reset(driver, timeout=TIMEOUT_RESET):
    try:
        W(driver, timeout).until(lambda d:
            (d.find_element(By.CSS_SELECTOR,
                f'input[data-ref="{G2}/_3_Departamento"]').get_attribute("value") or "") == ""
        )
        time.sleep(3)
    except: time.sleep(5)

# ─────────────────────────────────────────────
#  LLENAR FORMULARIO
# ─────────────────────────────────────────────
def llenar(driver, familia, consec):
    n = len(familia)
    cod = str(consec).zfill(3)
    p0 = familia[0]
    nombre0 = p0["nombre"]
    # apellido para nombre institución
    partes0 = nombre0.strip().split()
    apellido_fam = partes0[-2] if len(partes0) >= 3 else partes0[-1]
    nombre_inst = f"FAMILIA {apellido_fam}"
    micro = p0["micro"]
    hab = max(1, round(n/2 + 0.5))

    print(f"\n{'='*55}")
    print(f"  FAMILIA #{cod} | {nombre_inst} | {n} integrante(s)")
    print(f"{'='*55}")

    # ── S1 Política ──────────────────────────────────────────
    print("[S1] Autorización")
    set_radio(driver, f"{G1}/_1_1_TRATAMIENTO_Y_USO_DE_DATOS", "AUTOMATIC")

    # ── S2 Info general ──────────────────────────────────────
    print("[S2] Info general")
    set_radio(driver, f"{G2}/_2_1_Identificaci_n_d_vida_de_las_personas", "no_aplica")
    set_radio(driver, f"{G2}/_2_2_Datos_generales_del_escenario", "_")
    set_text(driver,  f"{G2}/_3_Departamento", DEPARTAMENTO)
    set_text(driver,  f"{G2}/_4_Unidad_Zonal_de_p_neaci_n_y_evaluaci_n", UNIDAD_ZONAL)
    set_text(driver,  f"{G2}/_5_Municipio_rea_no_municipalizada", MUNICIPIO)
    set_radio(driver, f"{G2}/_6_rea_de_ubicaci_n_de_la_vivienda", "_rea_urbana")
    set_text(driver,  f"{G2}/_7_Territorio", micro)
    set_text(driver,  f"{G2}/_8_Micro_territorio", micro)
    set_text(driver,  f"{G2}/_9_Corregimiento_C_Resguardos_Ind_genas", "Barrio")
    set_radio(driver, f"{G2}/_2_3_Detalles_del_Equipo_de_Salud", "_")
    set_number(driver,f"{G2}/_10_N_mero_de_identi_del_Equipo_de_Salud", ID_EQUIPO)
    set_text(driver,  f"{G2}/_11_Prestador_primar_del_Equipo_de_Salud", PRESTADOR)
    set_radio(driver, f"{G2}/_2_4_Detalle_de_perso_del_Equipo_de_salud", "_")
    set_radio(driver, f"{G2}/_12_Tipo_de_identifi_condiciones_de_salud", "cc__c_dula_de_ciudadan_a")
    set_number(driver,f"{G2}/_13_N_mero_de_identi_condiciones_de_salud", DOC_MARBEY)
    set_text(driver,  f"{G2}/_14_Perfil_profesion_condiciones_de_salud", PERFIL_MARBEY)
    set_number(driver,f"{G2}/_15_C_digo_de_ficha", CODIGO_FICHA)
    if p0["fecha"]:
        set_date(driver, f"{G2}/_16_Fecha_de_diligen_iamiento_de_la_ficha", fecha_str(p0["fecha"]))
    set_radio(driver, f"{G2}/_2_4_1_Datos_de_abordaje", "_")
    set_radio(driver, f"{G2}/_17_Entorno_donde_se_apacidades_y_riesgos", "hogar")
    set_text(driver,  f"{G2}/_18_Nombre_de_la_ins_d_organizaci_n_etc", nombre_inst)
    set_text(driver,  f"{G2}/_19_Cabeza_de_famili_Depende_del_entorno", nombre0)
    set_radio(driver, f"{G2}/_20_Alguno_de_los_ha_grama_j_venes_en_Paz", "no")
    # Dirección obligatoria
    set_text(driver,  f"{G2}/_21_Direcci_n", micro)
    set_text(driver,  f"{G2}/_24_Ubicaci_n_del_ho_Punto_de_referencia", micro)
    set_number(driver,f"{G2}/_25_N_de_identificaci_n_del_Hogar", int(cod))
    set_number(driver,f"{G2}/_26_N_de_identificaci_n_de_la_familia", int(cod))

    # ── S3 Entorno/vivienda ───────────────────────────────────
    print("[S3] Vivienda")
    set_radio(driver, f"{G2}/_27_Estrato_socio_ec_mico_de_la_vivienda", "medio")
    set_number(driver,f"{G2}/_28_N_de_hogares_fa_iares_en_la_vivienda", 1)
    set_number(driver,f"{G2}/_29_N_de_personas_en_la_vivienda", n)
    set_number(driver,f"{G2}/_30_N_de_habitaciones_de_vivienda", hab)
    set_number(driver,f"{G2}/_31_Cu_ntos_element_ualmente_en_su_hogar", n)
    set_number(driver,f"{G2}/_32_N_de_personas_por_habitaci_n", max(1,round(n/hab)))
    set_radio(driver, f"{G2}/_33_Hacinamiento", "no")

    set_radio(driver,    f"{G3}/_34_Tipo_de_Vivienda", "casa")
    set_radio(driver,    f"{G3}/_35_Cu_l_es_el_mate_techo_de_la_vivienda", "concreto")
    set_checkbox(driver, f"{G3}/_36_Se_identifica_a_dente_en_la_vivienda", "ninguno")
    set_radio(driver,    f"{G3}/_37_Se_observa_cerc_ores_de_enfermedades", "no")
    set_checkbox(driver, f"{G3}/_38_Observe_e_indagu_no_de_los_siguientes", "ninguno")
    set_radio(driver,    f"{G3}/_39_Al_interior_de_actividad_econ_mica", "no")
    set_checkbox(driver, f"{G3}/_40_Se_ale_los_anima_su_entorno_inmediato", "ninguno")
    # Sin animales seleccionados, pero estos 4 campos son obligatorios igual
    set_number(driver,   f"{G3}/_41_N_de_perros_en_la_vivienda", 0)
    set_number(driver,   f"{G3}/_42_Cantidad_de_perr_bica_en_la_vivienda", 0)
    set_number(driver,   f"{G3}/_43_N_de_gatos_en_la_vivienda", 0)
    set_number(driver,   f"{G3}/_44_Cantidad_de_gato_bica_en_la_vivienda", 0)
    # Sin animales → preguntas 41-44 no aplican, pero 45 sí
    set_radio(driver,    f"{G3}/_45_Los_perros_y_ga_ca_inferior_a_3_a_os", "no_aplica")
    set_checkbox(driver, f"{G3}/_46_Cu_l_es_la_prin_umano_en_la_vivienda", "acueducto_administrado_por_empresa_prest")
    set_checkbox(driver, f"{G3}/_47_Cu_l_es_el_sist_retas_en_la_vivienda", "sanitario_conectado_a_alcantarillado")
    set_checkbox(driver, f"{G3}/_48_Cu_l_es_el_sist_ticas_de_la_vivienda", "alcantarillado")
    set_checkbox(driver, f"{G3}/_49_C_mo_realiza_la_arios_de_la_vivienda", "recolecci_n_por_parte_del_servicio_de_as")

    # ── S4 Familia ────────────────────────────────────────────
    print("[S4] Familia")
    set_radio(driver,    f"{G4}/_50_Tipo_de_familia", tipo_familia(n))
    set_number(driver,   f"{G4}/_51_N_de_personas_q_conforman_la_familia", n)
    set_radio(driver,    f"{G4}/_52_Se_identifica_u_o_mayor_o_enfermedad", "no")
    set_checkbox(driver, f"{G4}/_54_Familia_en_la_qu_lud_F_sica_o_mental", "ninguna")
    set_checkbox(driver, f"{G4}/_55_Pr_cticas_que_fa_v_nculos_familiares", "cuando_hay_tensi_n_o_estr_s_en_la_famili")
    set_checkbox(driver, f"{G4}/_55_Pr_cticas_que_fa_v_nculos_familiares", "la_familia_resuelve_los_conflictos_busca")
    set_checkbox(driver, f"{G4}/_55_Pr_cticas_que_fa_v_nculos_familiares", "la_comunicaci_n_en_la_familia_est__basad")
    set_radio(driver,    f"{G4}/_56_Redes_de_apoyo_social", "cuentan_con_redes_de_apoyo_sociales_prot")
    set_checkbox(driver, f"{G4}/_57_Pr_cticas_para_e_en_el_entorno_hogar", "facilita_la_circulaci_n_del_aire_en_la_v")
    set_checkbox(driver, f"{G4}/_57_Pr_cticas_para_e_en_el_entorno_hogar", "los_residuos_s_lidos_se_almacenan_en_rec")
    set_checkbox(driver, f"{G4}/_57_Pr_cticas_para_e_en_el_entorno_hogar", "los_pisos_y_paredes_de_la_vivienda_est_n")
    set_checkbox(driver, f"{G4}/_57_Pr_cticas_para_e_en_el_entorno_hogar", "al_interior_de_la_vivienda_o_el_peridomi")

    # ── S5 Integrantes ────────────────────────────────────────
    for i, p in enumerate(familia):
        if i > 0:
            print(f"  → Agregando integrante {i+1}...")
            add_repeat(driver)
        print(f"[S5.{i+1}] {p['nombre']}")
        llenar_integrante(driver, p, i)

    # ── SUBMIT ────────────────────────────────────────────────
    print("[✓] Enviando...")
    try:
        cerrar_popups_vex(driver)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        cerrar_popups_vex(driver)
        btn = W(driver,15).until(EC.element_to_be_clickable((By.ID,"submit-form")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(3)
        for _ in range(3):
            try:
                ok = driver.find_element(By.XPATH,
                    "//button[normalize-space()='OK'] | //a[normalize-space()='OK']")
                ok.click(); time.sleep(2)
                btn = W(driver,10).until(EC.element_to_be_clickable((By.ID,"submit-form")))
                driver.execute_script("arguments[0].click();", btn); time.sleep(3)
            except: break
        esperar_reset(driver)
        print(f"  ✔ Familia #{cod} enviada")
        return True
    except Exception as e:
        print(f"  ✘ Submit error: {e}")
        return False


def llenar_integrante(driver, p, idx):
    """Llena la sección 5 de un integrante en el repeat con índice idx."""
    ctx = get_repeat(driver, idx)
    if ctx is None:
        print(f"    ✘ No encontré repeat #{idx}")
        return

    n1, n2, ap1, ap2 = dividir_nombre(p["nombre"])
    edad = p.get("edad", 30) or 30
    sexo = p.get("sexo","MASCULINO").upper()
    tipo_doc_val = TIPOS_DOC.get(p["tipo"], "cc__c_dula_de_ciudadan_a")
    reg_val = REGIMEN_VAL.get(p.get("regimen","SUB"), "subsdiado")
    edu_val = edu(edad)
    peso, talla = peso_talla(edad)
    imc_v = imc_val(peso, talla)
    rol_v = rol(idx, edad, sexo)
    es_menor = edad < 18

    # Sexo → género → identidad
    if sexo == "FEMENINO":
        sexo_v, genero_v, id_v = "mujer", "femenino", "femenino"
    else:
        sexo_v, genero_v, id_v = "hombre", "masculino", "masculino"

    tel = p.get("tel_raw","") or ""
    tel_num = limpiar(tel) if tel_valido(tel) else "0000000000"

    # 5.1 Identificación
    set_text(driver,   f"{G5}/_58_Primer_nombre", n1, ctx)
    if n2: set_text(driver, f"{G5}/_59_Segundo_nombre", n2, ctx)
    set_text(driver,   f"{G5}/_60_Primer_apellido", ap1, ctx)
    if ap2: set_text(driver, f"{G5}/_61_Segundo_apellido", ap2, ctx)
    set_radio(driver,  f"{G5}/_62_Tipo_de_identifi_grante_de_la_familia", tipo_doc_val, ctx)
    set_number(driver, f"{G5}/_63_N_mero_de_identificaci_n", p["id"], ctx)
    if p.get("fecha_nac"):
        set_date(driver, f"{G5}/_64_Fecha_de_nacimiento", fecha_str(p["fecha_nac"]), ctx)
    set_text(driver,   f"{G5}/_65_Nacionalidad", "Colombiana", ctx)
    set_radio(driver,  f"{G5}/_66_Sexo", sexo_v, ctx)
    set_radio(driver,  f"{G5}/_67_G_nero", genero_v, ctx)
    set_radio(driver,  f"{G5}/_68_Usted_se_identifica_como", id_v, ctx)
    set_radio(driver,  f"{G5}/_69_Cu_l_es_su_orientaci_n_sexual", "heterosexual", ctx)
    set_number(driver, f"{G5}/_70_Tel_fono_1", tel_num, ctx)
    set_radio(driver,  f"{G5}/_72_Rol_dentro_de_la_familia", rol_v, ctx)

    # 5.2 Socioeconómico
    set_radio(driver,  f"{G5}/_74_Nivel_educativo", edu_val, ctx)
    set_radio(driver,  f"{G5}/_75_R_gimen_de_afiliaci_n", reg_val, ctx)
    set_text(driver,   f"{G5}/_76_EAPB", p.get("eapb",""), ctx)

    # 77 Protección especial
    if es_menor:
        set_checkbox(driver, f"{G5}/_77_Se_reconoce_como_ciones_o_situaciones", "ni_as__ni_os_y_adolescentes", ctx)
    else:
        set_checkbox(driver, f"{G5}/_77_Se_reconoce_como_ciones_o_situaciones", "ninguna", ctx)

    # 79 Etnia
    set_radio(driver, f"{G5}/_79_Pertenencia_tnica", "ninguna", ctx)

    # 5.3 Salud
    set_checkbox(driver, f"{G5}/_81_Pr_cticas_de_cui_ales_tradicionales", "ninguna", ctx)
    set_radio(driver,    f"{G5}/_82_Se_reconoce_com_ona_con_discapacidad", "sin_discapacidad", ctx)
    set_radio(driver,    f"{G5}/_83_Tiene_certificac_discapacidad_RLCPD", "no_aplica", ctx)
    set_radio(driver,    f"{G5}/_84_Intenci_n_reproductiva_a_corto_plazo", "no", ctx)
    set_radio(driver,    f"{G5}/_85_Gestaci_n_actual_confirmada", "no", ctx)

    # 86 Prácticas salud
    set_checkbox(driver, f"{G5}/_86_Realiza_rutinar_cuidado_de_la_salud", "consume_alimentos_en_cantidad_y_calidad_", ctx)
    set_checkbox(driver, f"{G5}/_86_Realiza_rutinar_cuidado_de_la_salud", "higiene_oral_diaria__cepillado_m_nimo_2_", ctx)
    set_checkbox(driver, f"{G5}/_86_Realiza_rutinar_cuidado_de_la_salud", "lavado_de_manos_antes_de_consumir_alimen", ctx)
    set_checkbox(driver, f"{G5}/_86_Realiza_rutinar_cuidado_de_la_salud", "duerme_lo_suficiente_para_estar_con_ener", ctx)

    # 87 Atenciones pendientes
    set_checkbox(driver, f"{G5}/_87_Atenciones_pendi_ci_n_y_mantenimiento", "ninguna", ctx)
    # 88 Ruta materno perinatal
    set_checkbox(driver, f"{G5}/_88_Atenciones_pendi_ta_materno_perinatal", "ninguna", ctx)
    # 89 Motivo no atención
    set_checkbox(driver, f"{G5}/_89_Motivo_por_el_cu_o_materno_perinatal", "ninguna", ctx)

    # 90 Derechos
    set_checkbox(driver, f"{G5}/_90_Pr_cticas_para_e_l_derecho_a_la_salud", "conoce_los_derechos_y_deberes_en_salud", ctx)
    set_checkbox(driver, f"{G5}/_90_Pr_cticas_para_e_l_derecho_a_la_salud", "tiene_informaci_n_sobre_las_atenciones_y", ctx)
    set_checkbox(driver, f"{G5}/_90_Pr_cticas_para_e_l_derecho_a_la_salud", "conoce_los_lugares_donde_pueden_prestar_", ctx)
    set_checkbox(driver, f"{G5}/_90_Pr_cticas_para_e_l_derecho_a_la_salud", "conoce_c_mo_resolver_las_dificultades_qu", ctx)

    # 91 Lactancia
    set_radio(driver, f"{G5}/_91_Si_es_menor_de_6_ia_materna_exclusiva", "no_aplica", ctx)

    # 92-93 Peso/Talla
    set_number(driver, f"{G5}/_92_Peso_en_Kilogramos", peso, ctx)
    set_number(driver, f"{G5}/_93_Talla_en_cent_metros", talla, ctx)

    # 96 Clasificación nutricional
    set_radio(driver, f"{G5}/_96_Clasificaci_n_an_dad_o_si_es_gestante", imc_v, ctx)

    # 97 Signos desnutrición
    set_checkbox(driver, f"{G5}/_97_Se_identifican_s_de_3_meses_a_5_a_os", "no_aplica", ctx)

    # 98-99 Tensión arterial
    if not es_menor:
        set_number(driver, f"{G5}/Sist_lica_Num_rico", 110, ctx)
        set_number(driver, f"{G5}/Diast_lica_Num_rico", 70, ctx)
        set_radio(driver,  f"{G5}/_99_Clasificaci_n_de_rt_Associattion_2024",
                  "normal__sist_lica_menos_de_120_mmhg_y_di", ctx)
    else:
        set_radio(driver, f"{G5}/_99_Clasificaci_n_de_rt_Associattion_2024", "no_aplica", ctx)

    # 100 ENT
    set_checkbox(driver, f"{G5}/_100_Alg_n_profesio_edad_no_transmisible", "ninguna", ctx)
    # 101 Transmisibles
    set_checkbox(driver, f"{G5}/_101_Presenta_algun_e_salud_transmisible", "ninguna", ctx)
    # 102 Endémica
    set_checkbox(driver, f"{G5}/_102_Vive_en_zona_e_ta_signos_y_s_ntomas", "ninguna", ctx)
    # 103 Tratamiento
    set_radio(driver,    f"{G5}/_103_Si_en_las_tres_la_enfermedad_actual", "no_aplica", ctx)
    # 104 Motivo no atención
    set_checkbox(driver, f"{G5}/_104_Motivo_s_por_e_ha_recibido_atenci_n", "no_aplica", ctx)
    # 105 Situaciones 14-28
    set_checkbox(driver, f"{G5}/_105_Identifica_alg_tre_los_14_y_28_a_os", "no_aplica", ctx)
    # 106 Dos semanas
    set_radio(driver,    f"{G5}/_106_Durante_m_s_de_s_mayores_de_14_a_os", "no_aplica", ctx)
    # 107 8 días
    set_radio(driver,    f"{G5}/_107_Durante_los_lt_s_mayores_de_14_a_os", "no_aplica", ctx)
    # 108 Sustancias
    set_radio(driver,    f"{G5}/_108_En_los_ltimos_ndo_alcohol_y_tabaco", "no_aplica", ctx)
    # 110 Situación de salud
    set_radio(driver,    f"{G5}/_110_Ha_presentado_en_la_ltima_semana", "no", ctx)


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
def main():
    print("="*55)
    print("  BOT CARACTERIZACIÓN SI-APS RES 212 VIGENCIA 2026")
    print("="*55)
    try:
        familias = leer_familias(ARCHIVO_EXCEL, HOJA_NOMBRE)
        print(f"\n✔ {len(familias)} familia(s) cargadas en este Excel:")
        for i,f in enumerate(familias,1):
            print(f"  {i:03d}: {' | '.join(p['nombre'] for p in f)}")
    except Exception as e:
        print(f"\n✘ Error Excel: {e}"); return

    # Pregunta el consecutivo de familia con el que se debe empezar
    while True:
        resp = input("\n¿En qué número de familia/hogar empezamos? "
                      "(ej: 1 para 001, 15 para 015) [default 1]: ").strip()
        if resp == "":
            consec_inicial = 1
            break
        try:
            consec_inicial = int(resp)
            if consec_inicial < 1:
                print("  Debe ser un número mayor o igual a 1.")
                continue
            break
        except ValueError:
            print("  Por favor escribe solo un número (ej: 15).")

    print(f"\n✔ Esta corrida usará los consecutivos "
          f"{str(consec_inicial).zfill(3)} a "
          f"{str(consec_inicial + len(familias) - 1).zfill(3)}")

    input("\nPresiona ENTER para comenzar (Ctrl+C para cancelar)...")

    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

    ok = []
    fallida = None       # (numero_familia, nombre_jefe, motivo)
    detenido_por_error = False

    try:
        driver.get(URL_FORMULARIO)
        time.sleep(6)
        descartar_borrador(driver)

        for i, fam in enumerate(familias):
            consec = consec_inicial + i
            cod = str(consec).zfill(3)
            jefe = fam[0]["nombre"]
            print(f"\n[{i+1}/{len(familias)}] Familia {cod}")
            try:
                exito = llenar(driver, fam, consec)
                if exito:
                    ok.append(f"Familia {cod}: {jefe}")
                else:
                    # llenar() solo retorna False si el error ocurrió
                    # en el submit (ya reportado dentro de la función)
                    fallida = (cod, jefe, "Error al enviar el formulario (ver detalle arriba)")
                    detenido_por_error = True
                    break
            except CampoError as e:
                print(f"\n  ✘✘✘ ERROR llenando un campo: {e}")
                fallida = (cod, jefe, str(e))
                detenido_por_error = True
                break
            except Exception as e:
                print(f"\n  ✘✘✘ ERROR INESPERADO: {e}")
                fallida = (cod, jefe, str(e))
                detenido_por_error = True
                break
    finally:
        driver.quit()

    # ── INFORME FINAL ──────────────────────────────────────────
    print("\n"+"="*55)
    print("  INFORME FINAL")
    print("="*55)
    print(f"  ✔ Familias enviadas exitosamente: {len(ok)}")
    if ok:
        for o in ok:
            print(f"    • {o}")

    if detenido_por_error and fallida:
        cod_f, jefe_f, motivo = fallida
        print(f"\n  ⛔ PROCESO DETENIDO en la Familia {cod_f} ({jefe_f})")
        print(f"     Motivo: {motivo}")
        if ok:
            ultimo_cod = ok[-1].split(":")[0].replace("Familia ", "").strip()
            print(f"\n  ➜ La última familia que SÍ se cargó correctamente fue: {ultimo_cod}")
            print(f"  ➜ Para la próxima corrida, empieza en el consecutivo: "
                  f"{int(ultimo_cod) + 1}")
        else:
            print(f"\n  ➜ No se cargó ninguna familia en esta corrida.")
            print(f"  ➜ Revisa el error y vuelve a intentar desde el consecutivo: {cod_f}")
    elif len(ok) == len(familias):
        print(f"\n  🎉 Todas las familias del Excel se cargaron correctamente.")

    print("="*55)

if __name__ == "__main__":
    main()