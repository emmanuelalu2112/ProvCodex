# Requisitos:
#   pip install nicegui pandas requests gspread google-auth

from nicegui import ui, app
import pandas as pd
import requests
from io import StringIO
import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
import re, math

# ---------- CONFIG ----------
USERS = {"admin": "1234"}

CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRt0PJnBjc9mN3bKL48WXo-GPJHE5Kz0yD8FbEwFb8PnrlKuIKfIJG5A8mfqufT8irUDI2NEblFbe9X/pub?output=csv'
SERVICE_ACCOUNT_FILE = 'activacion-466414-5ac0b0de7fad.json'
TARGET_SPREADSHEET_ID = '14gsDElUZWUcAVa63-oZ_zMWMYbZbGv6FfomaNiH_y9o'
TARGET_SHEET_NAME = 'Hoja 1'

DEST_COL_FECHA = 'Fecha Compromiso'
DEST_COL_CANT  = 'Cantidad Compromiso'

# ---------- UTIL ----------
def is_logged_in() -> bool:
    return bool(app.storage.user.get('logged_in', False))

def login(username: str, password: str):
    if USERS.get(username) == password:
        app.storage.user['logged_in'] = True
        app.storage.user['username'] = username
        ui.notify(f'Bienvenido {username}')
        ui.navigate.to('/')
    else:
        ui.notify('Credenciales inválidas', type='negative')

def logout():
    app.storage.user.clear()
    ui.navigate.to('/login')

def cargar_df_desde_csv(url: str) -> pd.DataFrame:
    r = requests.get(url, timeout=12); r.raise_for_status()
    df = pd.read_csv(StringIO(r.text))
    df.columns = [c.strip() for c in df.columns]
    return df

def detectar_columna(df: pd.DataFrame, candidatos: set[str], contiene_todos: tuple[str, ...] = ()) -> str | None:
    normal = {c.lower().strip(): c for c in df.columns}
    for clave in candidatos:
        if clave in normal:
            return normal[clave]
    for k, v in normal.items():
        if all(p in k for p in contiene_todos):
            return v
    return None

def _s(x):
    if x is None: return ''
    if isinstance(x, float) and math.isnan(x): return ''
    return str(x)

def _has(val) -> bool:
    return val is not None and str(val).strip() != ''

def _norm(val) -> str:
    return '' if val is None else str(val).strip()

def normalizar_fecha(fecha: str) -> str:
    s = _norm(fecha)
    if not s: return ''
    m = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', s)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2: y = '20' + y
        return f'{y.zfill(4)}-{mo.zfill(2)}-{d.zfill(2)}'
    m = re.fullmatch(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m:
        y, mo, d = m.groups()
        return f'{y.zfill(4)}-{mo.zfill(2)}-{d.zfill(2)}'
    return s

def conectar_gspread():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds)

def get_dest_worksheet(dest_headers: list[str]):
    gc = conectar_gspread()
    sh = gc.open_by_key(TARGET_SPREADSHEET_ID)
    try:
        ws = sh.worksheet(TARGET_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=TARGET_SHEET_NAME, rows=2000, cols=max(10, len(dest_headers)))
        ws.append_row(dest_headers)
    ensure_dest_headers(ws, dest_headers)
    return ws

def ensure_dest_headers(ws, desired_headers: list[str]) -> list[str]:
    values = ws.get_all_values() or []
    current = values[0] if values else []
    if current != desired_headers:
        if len(current) < len(desired_headers):
            ws.resize(cols=len(desired_headers))
        ws.update('1:1', [desired_headers])
        return desired_headers
    return current

def append_fila_destino(dest_headers: list[str], row_final: dict):
    ws = get_dest_worksheet(dest_headers)
    ws_headers = ensure_dest_headers(ws, dest_headers)
    fila = [_s(row_final.get(h, '')) for h in ws_headers]
    try:
        ws.append_row(fila, value_input_option='USER_ENTERED')
    except TypeError:
        ws.append_row(fila)

def append_filas_destino_batch(dest_headers: list[str], rows_finales: list[dict]):
    if not rows_finales:
        return
    ws = get_dest_worksheet(dest_headers)
    ws_headers = ensure_dest_headers(ws, dest_headers)
    valores = [[_s(r.get(h, '')) for h in ws_headers] for r in rows_finales]
    # gspread >= 5.7: append_rows; fallback a append_row por compatibilidad
    try:
        ws.append_rows(valores, value_input_option='USER_ENTERED')
    except Exception:
        for fila in valores:
            try:
                ws.append_row(fila, value_input_option='USER_ENTERED')
            except TypeError:
                ws.append_row(fila)

def leer_df_destino(headers_default: list[str]) -> pd.DataFrame:
    gc = conectar_gspread()
    sh = gc.open_by_key(TARGET_SPREADSHEET_ID)
    try:
        ws = sh.worksheet(TARGET_SHEET_NAME)
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame(columns=headers_default + [DEST_COL_FECHA, DEST_COL_CANT])
        header = values[0]
        rows = [row + [''] * (len(header) - len(row)) for row in values[1:]]
        df = pd.DataFrame(rows, columns=header)
        df.insert(0, '__rownum', range(2, 2 + len(df)))  # fila real en Sheet (comienza en 2)
        return df
    except gspread.WorksheetNotFound:
        return pd.DataFrame(columns=headers_default + [DEST_COL_FECHA, DEST_COL_CANT])

def parsear_numeros_robusto(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    n1 = pd.to_numeric(s, errors='coerce')
    s2 = s.str.replace(r'[^0-9,\.\-]', '', regex=True)
    def _convert(v: str) -> str:
        if v.count(',') and v.count('.'):
            last = max(v.rfind(','), v.rfind('.'))
            import re as _re
            int_part = _re.sub(r'[^0-9\-]', '', v[:last])
            dec_part = _re.sub(r'[^0-9]', '', v[last+1:])
            return f'{int_part}.{dec_part}' if dec_part != '' else int_part
        elif v.count(',') and not v.count('.'):
            return v.replace(',', '.')
        else:
            return v.replace(',', '')
    s3 = s2.apply(_convert)
    n2 = pd.to_numeric(s3, errors='coerce')
    return n1.combine_first(n2)

# ---------- LOGIN ----------
@ui.page('/login')
def login_page():
    ui.page_title('Login')
    with ui.card().classes('p-6 w-96 mx-auto mt-20'):
        ui.label('Iniciar sesión').classes('text-lg mb-4')
        user_inp = ui.input('Usuario')
        pass_inp = ui.input('Contraseña').props('type=password')
        ui.button('Ingresar', on_click=lambda: login(user_inp.value, pass_inp.value)).classes('mt-4 w-full')

# ---------- APP ----------
@ui.page('/')
def main_page():
    if not is_logged_in():
        ui.navigate.to('/login'); return

    ui.page_title('Proveedores')
    with ui.header().classes('justify-between'):
        with ui.row().classes('items-center gap-2'):
            menu_btn = ui.button(icon='menu').props('flat round dense')
            ui.label('Gestión de Proveedores')
        with ui.row().classes('items-center gap-3'):
            ui.label(f"Usuario: {app.storage.user.get('username', '')}")
            ui.button('Salir', on_click=logout)

    # Carga ORIGEN
    try:
        df_original = cargar_df_desde_csv(CSV_URL)
    except Exception as e:
        ui.label(f'No se pudo cargar el origen: {e}').classes('text-negative')
        df_original = pd.DataFrame()

    # Detectar columnas
    col_pendiente    = detectar_columna(df_original, {'cantidad pendiente','cant. pendiente','pendiente','pendientes'}, ('pend',))
    col_desc_prov    = detectar_columna(df_original, {'descripcion proveedor','descripción proveedor'}, ('descrip','prove'))
    col_cod_art_prov = detectar_columna(df_original, {'cod. art. proveedor','cod art proveedor','codigo art proveedor','cód. art. proveedor'}, ('cod','art','prove'))

    # Filtrar pendientes > 0
    if not df_original.empty and col_pendiente in df_original.columns:
        try:
            nums = parsear_numeros_robusto(df_original[col_pendiente])
            df_pend = df_original[nums > 0].copy()
            df_pend[col_pendiente] = nums[nums > 0]
            aviso = f'Se muestran SOLO filas con "{col_pendiente} > 0".'
        except Exception as e:
            df_pend = df_original.iloc[0:0].copy()
            aviso = f'Error al filtrar por {col_pendiente}: {e}'
    else:
        df_pend = df_original.iloc[0:0].copy()
        aviso = 'No encontré la columna de pendiente en el ORIGEN.'

    # Modelo base
    df_pend['__id']         = range(1, len(df_pend)+1)
    df_pend['__checked']    = False
    df_pend['__fecha_fila'] = ''
    df_pend['__cant_fila']  = ''

    filas        = df_pend.to_dict(orient='records')
    vista_actual = list(filas)
    by_id        = {r['__id']: r for r in filas}
    selected_ids: set[int] = set()

    # ---------- LAYOUT: Left Drawer (sidebar) + contenido ----------
    with ui.left_drawer(value=True, bordered=False).classes('bg-grey-2') as drawer:
        with ui.column().classes('w-full items-start p-3 gap-2'):
            # LOGO / título
            with ui.row().classes('items-center gap-2'):
                ui.icon('inventory_2').classes('text-primary')
                ui.label('Proveedores').classes('text-primary text-lg')
            ui.separator()

            # Tabs verticales de navegación
            with ui.tabs().props('vertical dense').classes('w-full') as tabs:
                ui.label('Órdenes').classes('text-xs text-grey-7 q-ml-sm q-mt-sm')
                t1 = ui.tab('PENDIENTES (> 0)', icon='assignment_late').classes('text-left w-full')
                ui.separator()
                ui.label('Consultas').classes('text-xs text-grey-7 q-ml-sm q-mt-sm')
                t2 = ui.tab('GRABADOS (DESTINO)', icon='done_all').classes('text-left w-full')
                t3 = ui.tab('MODIFICAR EXISTENTES', icon='edit_note').classes('text-left w-full')

    # Toggle del drawer desde el botón hamburguesa
    def toggle_drawer():
        drawer.value = not drawer.value
        drawer.update()
    menu_btn.on('click', lambda e: toggle_drawer())

    # Paneles a la derecha (contenido)
    with ui.tab_panels(tabs, value=t1).props('animated').classes('w-full'):
        # =================== TAB 1 ===================
        with ui.tab_panel(t1):

            # Filtros (COLAPSADOS por defecto)
            with ui.expansion('Filtros por columna', icon='filter_list', value=False).classes('w-full'):
                columnas_filtrables = [c for c in [col_desc_prov, col_cod_art_prov] if c]
                filtros_inputs = {}
                with ui.grid(columns=4).classes('gap-2'):
                    for col in columnas_filtrables:
                        filtros_inputs[col] = ui.input(label=f'Filtrar {col}', placeholder='contiene...')
                def aplicar_filtros():
                    textos  = {c:(filtros_inputs[c].value or '').strip().lower() for c in columnas_filtrables}
                    activos = {c:t for c,t in textos.items() if t}
                    if not activos:
                        vista_actual[:] = filas
                    else:
                        vista_actual[:] = [
                            r for r in filas
                            if all((str(r.get(c,'')).lower().find(needle) >= 0)
                                   for c,needle in activos.items())
                        ]
                    tabla.rows = vista_actual
                    tabla.update()
                def limpiar_filtros():
                    for inp in filtros_inputs.values(): inp.value = ''
                    aplicar_filtros()
                with ui.row().classes('gap-2 q-mt-sm'):
                    ui.button('APLICAR FILTROS', on_click=aplicar_filtros, color='primary')
                    ui.button('LIMPIAR FILTROS', on_click=limpiar_filtros).props('flat')
                for inp in filtros_inputs.values():
                    inp.props('debounce=300')
                    inp.on('update:model-value', lambda e: aplicar_filtros())
                    inp.on('keydown.enter',     lambda e: aplicar_filtros())

            # Cabecera (fecha/cantidad)
            with ui.row().classes('items-center gap-3 q-mt-md'):
                fecha_global = ui.input(label='Fecha (CABECERA, opcional)').props('type=date')
                cant_global  = ui.input(label='Cantidad (CABECERA, opcional)').props('type=number')

            ui.separator()
            ui.label(aviso)

            # Acciones de selección
            contador = ui.label('Seleccionados: 0').classes('text-primary')
            def actualizar_contador():
                cuenta = sum(1 for r in filas if bool(r.get('__checked')))
                contador.text = f'Seleccionados: {cuenta}'
                contador.update()

            def seleccionar_todo_vista():
                for r in vista_actual:
                    r['__checked'] = True
                tabla.update(); actualizar_contador()

            def seleccionar_tres_vista():
                for r in vista_actual:
                    r['__checked'] = False
                for r in vista_actual[:3]:
                    r['__checked'] = True
                tabla.update(); actualizar_contador()

            def limpiar_todas():
                for r in filas:
                    r['__checked'] = False
                tabla.update(); actualizar_contador()

            with ui.row().classes('items-center gap-3'):
                ui.button('Seleccionar 3 primeras (vista)', on_click=seleccionar_tres_vista)
                ui.button('Seleccionar todo (vista)', on_click=seleccionar_todo_vista)
                ui.button('Limpiar selección (todas)', on_click=limpiar_todas).props('flat')
                # Vista rápida alternativa con AG Grid
                def abrir_ultraliviana():
                    dlg = ui.dialog().props('maximized')
                    headers_origen = list(df_original.columns)
                    # elegir columnas clave si existen
                    col1 = col_desc_prov or (headers_origen[0] if headers_origen else '')
                    col2 = col_cod_art_prov or (headers_origen[1] if len(headers_origen) > 1 else '')
                    # construir HTML estático
                    filas_html = []
                    for r in vista_actual:
                        rid = int(r.get('__id'))
                        v1 = _s(r.get(col1, '')) if col1 else ''
                        v2 = _s(r.get(col2, '')) if col2 else ''
                        f0 = _s(r.get('__fecha_fila', ''))
                        c0 = _s(r.get('__cant_fila', ''))
                        filas_html.append(
                            f'''<tr class="ultra-row" data-id="{rid}">
                                  <td><input type="checkbox" class="sel" /></td>
                                  <td class="t">{v1}</td>
                                  <td class="t">{v2}</td>
                                  <td><input type="date" class="f" value="{f0}"/></td>
                                  <td><input type="number" class="c" value="{c0}" style="width:110px"/></td>
                                </tr>'''
                        )
                    tabla_html = f'''
                    <div class="q-pa-sm" style="max-height:75vh; overflow:auto;">
                      <table class="q-table" style="min-width:720px;">
                        <thead>
                          <tr>
                            <th>Sel</th>
                            <th>{col1 or ''}</th>
                            <th>{col2 or ''}</th>
                            <th>Fecha (fila)</th>
                            <th>Cant. (fila)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {''.join(filas_html)}
                        </tbody>
                      </table>
                    </div>
                    <div class="q-gutter-sm q-mt-sm">
                      <q-btn color="primary" label="Grabar SELECCIONADOS (HTML)" @click="
                        (() => {{
                          const out = [];
                          const rows = Array.from($el.querySelectorAll('.ultra-row'));
                          for (const row of rows) {{
                            const chk = row.querySelector('.sel');
                            if (chk && chk.checked) {{
                              const id = Number(row.getAttribute('data-id') || '0');
                              const f  = (row.querySelector('.f')?.value || '').trim();
                              const c  = (row.querySelector('.c')?.value || '').trim();
                              out.push({{ id, fecha: f, cant: c }});
                            }}
                          }}
                          $emit('ultra-submit', out);
                        }})()
                      "/>
                      <q-btn flat label="Cerrar" @click="$parent.hide()"/>
                    </div>
                    '''
                    with dlg, ui.card().classes('q-pa-md fit'):
                        html = ui.html(tabla_html)
                        def on_ultra_submit(e):
                            try:
                                payload = e.args if isinstance(e.args, list) else []
                                if not payload:
                                    ui.notify('No hay filas seleccionadas (HTML)', type='warning'); return
                                dest_headers = headers_origen + [DEST_COL_FECHA, DEST_COL_CANT]
                                nuevas = []
                                f_cab = normalizar_fecha(fecha_global.value)
                                c_cab = _norm(cant_global.value)
                                def toma(pos, cab, es_fecha=False):
                                    if _has(pos):
                                        return normalizar_fecha(pos) if es_fecha else _norm(pos)
                                    if _has(cab):
                                        return normalizar_fecha(cab) if es_fecha else _norm(cab)
                                    return None
                                for item in payload:
                                    rid = int(item.get('id', 0))
                                    r = by_id.get(rid)
                                    if not r: continue
                                    f_final = toma(item.get('fecha','') or r.get('__fecha_fila',''), f_cab, es_fecha=True)
                                    c_final = toma(item.get('cant','')  or r.get('__cant_fila',''),  c_cab, es_fecha=False)
                                    row_final = {h: _s(r.get(h, '')) for h in headers_origen}
                                    row_final[DEST_COL_FECHA] = '' if not f_final else f_final
                                    row_final[DEST_COL_CANT]  = '' if not c_final else c_final
                                    nuevas.append(row_final)
                                append_filas_destino_batch(dest_headers, nuevas)
                                ui.notify(f'Se grabaron {len(nuevas)} fila(s) desde HTML')
                            except Exception as ex:
                                ui.notify(f'Error al grabar (HTML): {ex}', type='negative')
                        html.on('ultra-submit', on_ultra_submit)
                    dlg.open()
                ui.button('Vista ultraliviana (HTML)', on_click=abrir_ultraliviana)
                selected_ids_ag: set[int] = set()
                def abrir_aggrid():
                    dlg = ui.dialog().props('maximized')
                    with dlg, ui.card().classes('q-pa-md fit'):
                        ui.label('Vista rápida (AG Grid) - selección local, alto rendimiento')
                        # preparar datos y columnas
                        headers_origen = list(df_original.columns)
                        base_cols = [{'headerName': c, 'field': c, 'sortable': True, 'filter': True, 'resizable': True} for c in headers_origen]
                        ag_columns = (
                            [{'headerName': 'Sel', 'checkboxSelection': True, 'headerCheckboxSelection': True, 'width': 60, 'pinned': 'left'}] +
                            [{'headerName': 'ID', 'field': '__id', 'width': 90, 'pinned': 'left'}] +
                            base_cols +
                            [{'headerName': 'Fecha (fila)', 'field': '__fecha_fila', 'editable': True, 'width': 140},
                             {'headerName': 'Cant. (fila)',  'field': '__cant_fila',  'editable': True, 'width': 120}]
                        )
                        ag_options = {
                            'rowSelection': 'multiple',
                            'suppressRowClickSelection': True,
                            'animateRows': False,
                            'pagination': True,
                            'paginationPageSize': 50,
                            'defaultColDef': {'sortable': True, 'filter': True, 'resizable': True},
                            'columnDefs': ag_columns,
                            'rowData': vista_actual,
                        }
                        def on_selection_changed(e):
                            nonlocal selected_ids_ag
                            try:
                                rows = e.args.get('selected', []) if isinstance(e.args, dict) else []
                                selected_ids_ag.clear()
                                for r in rows:
                                    rid = int(r.get('__id'))
                                    selected_ids_ag.add(rid)
                            except Exception:
                                pass
                        grid = ui.aggrid(ag_options).on('selectionChanged', on_selection_changed)
                        with ui.row().classes('q-mt-md gap-2'):
                            def grabar_aggrid():
                                sel_ids = list(selected_ids_ag) if selected_ids_ag else [int(r['__id']) for r in grid.options.get('rowData', []) if bool(r.get('__checked'))]
                                if not sel_ids:
                                    ui.notify('No hay filas seleccionadas en AG Grid', type='warning'); return
                                try:
                                    headers_origen = list(df_original.columns)
                                    dest_headers   = headers_origen + [DEST_COL_FECHA, DEST_COL_CANT]
                                    nuevas: list[dict] = []
                                    f_cab = normalizar_fecha(fecha_global.value)
                                    c_cab = _norm(cant_global.value)
                                    def toma(pos, cab, es_fecha=False):
                                        if _has(pos):
                                            return normalizar_fecha(pos) if es_fecha else _norm(pos)
                                        if _has(cab):
                                            return normalizar_fecha(cab) if es_fecha else _norm(cab)
                                        return None
                                    for rid in sel_ids:
                                        r = by_id.get(int(rid))
                                        if not r: continue
                                        f_pos = r.get('__fecha_fila')
                                        c_pos = r.get('__cant_fila')
                                        f_final = toma(f_pos, f_cab, es_fecha=True)
                                        c_final = toma(c_pos, c_cab, es_fecha=False)
                                        row_final = {h: _s(r.get(h, '')) for h in headers_origen}
                                        row_final[DEST_COL_FECHA] = '' if not f_final else f_final
                                        row_final[DEST_COL_CANT]  = '' if not c_final else c_final
                                        nuevas.append(row_final)
                                    append_filas_destino_batch(dest_headers, nuevas)
                                    ui.notify(f'Se grabaron {len(nuevas)} fila(s) desde AG Grid')
                                except Exception as ex:
                                    ui.notify(f'Error al grabar (AG Grid): {ex}', type='negative')
                            ui.button('Grabar SELECCIONADOS (AG Grid)', on_click=grabar_aggrid, color='primary')
                            ui.button('Cerrar', on_click=dlg.close)
                    dlg.open()
                ui.button('Vista rápida (AG Grid)', on_click=abrir_aggrid, color='secondary')

            # Tabla principal (scroll interno + horizontal)
            row_h, header_h, visible = 44, 56, 10
            table_h = header_h + row_h*visible

            cols = [{'name': c, 'label': c, 'field': c, 'sortable': True} for c in df_original.columns]
            cols += [
                {'name': 'Seleccionar', 'label': 'Seleccionar', 'field': '__checked', 'sortable': False},
                {'name': 'Fecha (fila)', 'label': 'Fecha (fila)', 'field': '__fecha_fila', 'sortable': False},
                {'name': 'Cant. (fila)', 'label': 'Cant. (fila)', 'field': '__cant_fila', 'sortable': False},
            ]
            total_cols = len(df_original.columns) + 3
            min_w = 120 * total_cols

            with ui.element('div').style(f'height:{table_h}px; max-width:100%; overflow:auto;'):
                tabla = (
                    ui.table(columns=cols, rows=vista_actual, row_key='__id')
                      .props('wrap-cells=false')
                      .props('virtual-scroll')
                      .props('virtual-scroll-item-size=44')
                      .props('virtual-scroll-slice-size=20')
                      .props('no-animation')
                      .props("rows-per-page-options='[25,50,100]'")
                      .props(":pagination='{rowsPerPage:50}'")
                      .props('hide-bottom')
                      .style(f'min-width:{min_w}px')
                      .classes('w-full')
                )

            # Slots
            tabla.add_slot(
                'body-cell-Seleccionar',
                '''
                <q-td :props="props">
                  <q-icon
                    :name="props.row.__checked ? 'check_box' : 'check_box_outline_blank'"
                    size="md"
                    class="cursor-pointer"
                    @click="() => {
                      const nv = !props.row.__checked;
                      props.row.__checked = nv;
                      $parent.$emit('toggle-select-lite', { id: props.row.__id, checked: nv })
                    }"
                  />
                </q-td>
                '''
            )
            tabla.add_slot(
                'body-cell-Fecha (fila)',
                '''
                <q-td :props="props">
                  <q-input dense type="date"
                    :model-value="props.row.__fecha_fila"
                    @blur="e => {
                      const val = e.target.value || '';
                      props.row.__fecha_fila = val || '';
                      $parent.$emit('update-cell', { id: props.row.__id, field: '__fecha_fila', value: val || '' })
                    }"
                  />
                </q-td>
                '''
            )
            tabla.add_slot(
                'body-cell-Cant. (fila)',
                '''
                <q-td :props="props">
                  <q-input dense type="number" style="max-width:110px"
                    debounce="500"
                    :model-value="props.row.__cant_fila"
                    @keyup.enter="e => {
                      const val = e.target.value || '';
                      props.row.__cant_fila = val;
                      $parent.$emit('update-cell', { id: props.row.__id, field: '__cant_fila', value: val })
                    }"
                    @blur="e => {
                      const val = e.target.value || '';
                      props.row.__cant_fila = val;
                      $parent.$emit('update-cell', { id: props.row.__id, field: '__cant_fila', value: val })
                    }"
                  />
                </q-td>
                '''
            )

            def on_update_cell(e):
                args = e.args if isinstance(e.args, dict) else {}
                rid   = args.get('id')
                field = args.get('field')
                value = args.get('value', '')
                if rid is None: return
                rid = int(rid)
                fila = by_id.get(rid)
                if not fila: return
                if field == '__fecha_fila':
                    fila['__fecha_fila'] = normalizar_fecha(value)
                elif field == '__cant_fila':
                    fila['__cant_fila'] = value
                # Evitar re-render innecesario en cada tecleo para mejorar rendimiento

            def on_toggle_select_lite(e):
                args = e.args if isinstance(e.args, dict) else {}
                rid = args.get('id')
                checked = bool(args.get('checked', False))
                if rid is None: return
                rid = int(rid)
                fila = by_id.get(rid)
                if not fila: return
                fila['__checked'] = checked
                actualizar_contador()

            tabla.on('update-cell',  on_update_cell)
            tabla.on('toggle-select-lite', on_toggle_select_lite)

            # Botón Grabar + resultados
            async def grabar_seleccionados():
                # Priorizar selección de AG Grid si existe y no está vacía
                try:
                    sel_ids = list(selected_ids_ag) if selected_ids_ag else [int(r['__id']) for r in filas if bool(r.get('__checked'))]
                except NameError:
                    sel_ids = [int(r['__id']) for r in filas if bool(r.get('__checked'))]
                if not sel_ids:
                    ui.notify('No hay filas seleccionadas', type='warning'); return

                btn_grabar.disable()
                try:
                    headers_origen = list(df_original.columns)
                    dest_headers   = headers_origen + [DEST_COL_FECHA, DEST_COL_CANT]
                    nuevas: list[dict] = []

                    f_cab = normalizar_fecha(fecha_global.value)
                    c_cab = _norm(cant_global.value)

                    def toma(pos, cab, es_fecha=False):
                        if _has(pos):
                            return normalizar_fecha(pos) if es_fecha else _norm(pos)
                        if _has(cab):
                            return normalizar_fecha(cab) if es_fecha else _norm(cab)
                        return None

                    for rid in sel_ids:
                        r = by_id.get(int(rid))
                        if not r: continue

                        f_pos = r.get('__fecha_fila')
                        c_pos = r.get('__cant_fila')

                        f_final = toma(f_pos, f_cab, es_fecha=True)    # POS > CAB > vacío
                        c_final = toma(c_pos, c_cab, es_fecha=False)   # POS > CAB > vacío

                        row_final = {h: _s(r.get(h, '')) for h in headers_origen}
                        row_final[DEST_COL_FECHA] = '' if not f_final else f_final
                        row_final[DEST_COL_CANT]  = '' if not c_final else c_final

                        nuevas.append(row_final)
                    append_filas_destino_batch(dest_headers, nuevas)

                    ui.notify(f'Se grabaron {len(nuevas)} fila(s)')
                    tabla_temp.rows = nuevas; tabla_temp.update()
                except Exception as ex:
                    ui.notify(f'Error al grabar: {ex}', type='negative')
                finally:
                    btn_grabar.enable()

            with ui.row().classes('items-center gap-3 q-mt-md'):
                btn_grabar = ui.button('Grabar SELECCIONADOS en otro Sheet', on_click=grabar_seleccionados, color='primary')

            # Tabla de resultados
            temp_headers = list(df_original.columns) + [DEST_COL_FECHA, DEST_COL_CANT]
            cols_temp = [{'name': c,'label': c,'field': c,'sortable': True} for c in temp_headers]
            tabla_temp = ui.table(columns=cols_temp, rows=[]).classes('w-full')

        # =================== TAB 2 ===================
        with ui.tab_panel(t2):
            destino_container = ui.column().classes('w-full')
            df_dest0 = leer_df_destino(headers_default=list(df_original.columns))
            cols_dest0 = [{'name': c,'label': c,'field': c,'sortable': True} for c in df_dest0.columns]
            rows_dest0 = df_dest0.to_dict(orient='records')
            with destino_container:
                ui.label('Contenido total del Sheet destino (inicial):')
                ui.table(columns=cols_dest0, rows=rows_dest0).props("rows-per-page-options='[25,50,100]'").props(":pagination='{rowsPerPage:50}'").classes('w-full')

        # =================== TAB 3: MODIFICAR EXISTENTES ===================
        with ui.tab_panel(t3):
            ui.label('Editar fechas y cantidades ya grabadas').classes('q-mb-sm')

            df_dest = leer_df_destino(headers_default=list(df_original.columns))
            if DEST_COL_FECHA not in df_dest.columns or DEST_COL_CANT not in df_dest.columns:
                ui.label(f'El Sheet destino no tiene columnas "{DEST_COL_FECHA}" y "{DEST_COL_CANT}".').classes('text-negative')
            else:
                # Tabla editable enfocada en Fecha/Cantidad
                cols_mod = [
                    {'name': '__rownum', 'label': 'Fila', 'field': '__rownum', 'sortable': True},
                ]
                for c in df_original.columns:
                    if c in df_dest.columns:
                        cols_mod.append({'name': c, 'label': c, 'field': c, 'sortable': True})
                cols_mod += [
                    {'name': 'Seleccionar', 'label': 'Seleccionar', 'field': '__checked', 'sortable': False},
                    {'name': DEST_COL_FECHA, 'label': DEST_COL_FECHA, 'field': DEST_COL_FECHA, 'sortable': False},
                    {'name': DEST_COL_CANT,  'label': DEST_COL_CANT,  'field': DEST_COL_CANT,  'sortable': False},
                ]

                rows_mod = df_dest.copy()
                if '__checked' not in rows_mod.columns:
                    rows_mod['__checked'] = False
                rows_mod = rows_mod.to_dict(orient='records')
                selected_rows_mod: set[int] = set()

                tabla_mod = ui.table(columns=cols_mod, rows=rows_mod, row_key='__rownum').classes('w-full')
                tabla_mod.props('wrap-cells=false')
                tabla_mod.props('virtual-scroll')
                tabla_mod.props('virtual-scroll-item-size=44')
                tabla_mod.props('virtual-scroll-slice-size=20')
                tabla_mod.props('no-animation')
                tabla_mod.props("rows-per-page-options='[0]'")
                tabla_mod.props(":pagination='{rowsPerPage:0}'")
                tabla_mod.props('hide-bottom')
                tabla_mod.props("rows-per-page-options='[25,50,100]'")
                tabla_mod.props(":pagination='{rowsPerPage:50}'")
                # quitamos row-class reactivo para evitar re-cálculo global por click

                # Slots de selección y edición
                tabla_mod.add_slot(
                    'body-cell-Seleccionar',
                    '''
                    <q-td :props="props">
                      <q-icon
                        :name="props.row.__checked ? 'check_box' : 'check_box_outline_blank'"
                        size="md"
                        class="cursor-pointer"
                        @click="() => {
                          const nv = !props.row.__checked;
                          props.row.__checked = nv;
                          $parent.$emit('toggle-select-mod', { rownum: props.row.__rownum, checked: nv })
                        }"
                      />
                    </q-td>
                    '''
                )
                tabla_mod.add_slot(
                    f'body-cell-{DEST_COL_FECHA}',
                    f'''
                    <q-td :props="props">
                      <q-input dense type="date"
                        :model-value="props.row['{DEST_COL_FECHA}']"
                        @blur="e => {{
                          const val = e.target.value || '';
                          props.row['{DEST_COL_FECHA}'] = val;
                          $parent.$emit('update-cell-mod', {{ rownum: props.row.__rownum, field: '{DEST_COL_FECHA}', value: val }});
                        }}"
                      />
                    </q-td>
                    '''
                )
                tabla_mod.add_slot(
                    f'body-cell-{DEST_COL_CANT}',
                    f'''
                    <q-td :props="props">
                      <q-input dense type="number" style="max-width:110px"
                        debounce="500"
                        :model-value="props.row['{DEST_COL_CANT}']"
                        @keyup.enter="e => {{
                          const val = e.target.value || '';
                          props.row['{DEST_COL_CANT}'] = val;
                          $parent.$emit('update-cell-mod', {{ rownum: props.row.__rownum, field: '{DEST_COL_CANT}', value: val }});
                        }}"
                        @blur="e => {{
                          const val = e.target.value || '';
                          props.row['{DEST_COL_CANT}'] = val;
                          $parent.$emit('update-cell-mod', {{ rownum: props.row.__rownum, field: '{DEST_COL_CANT}', value: val }});
                        }}"
                      />
                    </q-td>
                    '''
                )

                def on_toggle_select_mod(e):
                    args = e.args if isinstance(e.args, dict) else {}
                    rownum = args.get('rownum')
                    checked = bool(args.get('checked', False))
                    if rownum is None:
                        return
                    rownum = int(rownum)
                    if checked:
                        selected_rows_mod.add(rownum)
                    else:
                        selected_rows_mod.discard(rownum)

                def on_update_cell_mod(e):
                    args = e.args if isinstance(e.args, dict) else {}
                    rownum = args.get('rownum')
                    field = args.get('field')
                    value = args.get('value', '')
                    if rownum is None or field not in {DEST_COL_FECHA, DEST_COL_CANT}:
                        return
                    # normalizar fecha si aplica
                    if field == DEST_COL_FECHA:
                        value = normalizar_fecha(value)
                    # actualizar en memoria
                    for r in tabla_mod.rows:
                        if int(r.get('__rownum', -1)) == int(rownum):
                            r[field] = value
                            break
                    tabla_mod.update()

                tabla_mod.on('toggle-select-mod', on_toggle_select_mod)
                tabla_mod.on('update-cell-mod',  on_update_cell_mod)

                def batch_update_destino(updates: list[dict]):
                    if not updates:
                        return
                    headers_origen = list(df_original.columns)
                    dest_headers = headers_origen + [DEST_COL_FECHA, DEST_COL_CANT]
                    ws = get_dest_worksheet(dest_headers)
                    header = ws.row_values(1)
                    try:
                        col_fecha = header.index(DEST_COL_FECHA) + 1
                        col_cant  = header.index(DEST_COL_CANT) + 1
                    except ValueError:
                        raise RuntimeError('No encontré columnas destino para actualizar')
                    requests_payload = []
                    values_payload = []
                    for u in updates:
                        row = int(u['row'])
                        f = _s(u.get('fecha', ''))
                        c = _s(u.get('cant', ''))
                        start = rowcol_to_a1(row, col_fecha)
                        end   = rowcol_to_a1(row, col_cant)
                        a1 = f'{start}:{end}'
                        requests_payload.append({'range': a1, 'values': [[f, c]]})
                    ws.batch_update(requests_payload, value_input_option='USER_ENTERED')

                def guardar_modificaciones():
                    if not selected_rows_mod:
                        ui.notify('No hay filas seleccionadas para modificar', type='warning'); return
                    btn_guardar.disable()
                    try:
                        # construir updates desde tabla_mod.rows filtrando por selección
                        pendientes = []
                        for r in tabla_mod.rows:
                            rownum = int(r.get('__rownum', -1))
                            if rownum in selected_rows_mod:
                                pendientes.append({
                                    'row': rownum,
                                    'fecha': normalizar_fecha(r.get(DEST_COL_FECHA, '')),
                                    'cant': _s(r.get(DEST_COL_CANT, '')),
                                })
                        batch_update_destino(pendientes)
                        ui.notify(f'Se actualizaron {len(pendientes)} fila(s)')
                    except Exception as ex:
                        ui.notify(f'Error al actualizar: {ex}', type='negative')
                    finally:
                        btn_guardar.enable()

                with ui.row().classes('items-center gap-3 q-mt-md'):
                    btn_guardar = ui.button('Guardar modificaciones seleccionadas', on_click=guardar_modificaciones, color='primary')

# ---------- RUN ----------
ui.run(storage_secret='cambia-esta-clave-super-larga-y-aleatoria-para-sesiones-1234567890', reload=False)
