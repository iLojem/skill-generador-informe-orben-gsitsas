---
name: generar-acta-parcial-desde-chat
description: Procesa un chat exportado de WhatsApp (.txt o .zip) y la plantilla oficial MNN-REG-048 (.xlsx) para generar el Excel de consumo de materiales y el Acta Parcial de Entrega diligenciada, preguntando primero las fechas de corte.
---

# Generador de Acta Parcial de Entrega (MNN-REG-048) desde Chat de WhatsApp

## Regla de Inicio (OBLIGATORIA)
Antes de procesar o generar archivos, **pregunta al usuario**:
1. **Rango de fechas de corte** (`Fecha desde` y `Fecha hasta` en formato `YYYY-MM-DD`) para filtrar los registros del chat.
2. **Regla de Tubería y Coraza:** Si desea calcular `TUBERIA IMC 3/4 Y CORAZAS 3/4` con la fórmula de obra (`TUBERIA 3/4 × 3 + CORAZA`) o como suma directa de unidades (`TUBERIA 3/4 + CORAZA + LIQUID TIGHT` y equivalentes de `1"`).
3. **Valores manuales del resumen (si aplican para la quincena):**
   - `REVIONES DE PUNTO Y REPARACIONES` (UND)
   - `DISPONIBILIDAD INMEDIATA` (UND)
   - `PERSONAL HSEQ` (DIAS)
   - Datos opcionales de encabezado si cambian (Consecutivo, Fecha Inicio, Fecha Finalización, Avance %, Orden de Trabajo, # RA, Solicitante, Jefe del Área).

No ejecutes la generación hasta recibir las fechas de corte.

---

## Paso 1: Extracción y limpieza desde el Chat (.txt o .zip)
1. Si el archivo es `.zip`, extrae el `.txt` interno.
2. Limpia prefijos de WhatsApp (`[dd/mm/yy, hh:mm:ss] Remitente: [Forwarded]` o `dd/mm/yyyy, hh:mm - Remitente:`), caracteres invisibles Unicode (categoría `Cf` como `\u2060`, `\u200b`, `\ufeff`), espacios de no separación (`\u202f`, `\xa0`) y espacios múltiples.
3. Divide el texto por cada bloque que inicie con `REGISTRO DE ACTIVIDAD`.
4. Para cada bloque:
   - Extrae `Fecha: YYYY-MM-DD` y **filtra** únicamente los que estén dentro del rango `[Fecha desde, Fecha hasta]`.
   - Extrae la línea siguiente a `MARQUILLA:` como texto principal (`DATOS`). Si no existe `MARQUILLA:`, usa `Actividad` + `Área / Máquina`.
   - Extrae los ítems bajo `MATERIAL UTILIZADO:` (viñetas `*`, `•` o `-`) separando cantidad inicial y descripción, normalizando tildes, mayúsculas y errores comunes de digitación:
     - `cable`, `utp`, `cat6` → `CABLE UTP `
     - `jack`, `jacks`, `jac`, `jacs`, `yak` → `JACKS`
     - `zplug`, `z-plug`, `z plug`, `zplu`, `splug` → `ZPLUG`
     - `patch`, `passcord`, `passcor`, `pachcor`, `pascor`, `patchcore` → `PATCHCORE`
     - `fibra` → `FIBRA ARMADA`
     - `caja` + (`4x4`, `4×4`, `4s`, `rawelt`) → `CAJA 4S` (si aparece `Tapa 4x4` junto con `Caja 4x4`, no duplicar; si solo aparece `Tapa` sin `Caja 4x4`, asignar a `CAJA 4S`)
     - `face` → `FACE PLATE`
     - `tuberia`, `tubo`, `imc` → `TUBERIA 2"` (si dice 2"), `TUBERIA  1'` (si dice 1", 1.5" o 1'), si no `TUBERIA  3/4`
     - `coraza`, `corasa` → `CORAZA 1'` (si dice 1", 1.5" o 1'), si no `CORAZA`
     - `liquid`, `likid`, `licuid` → `LIQUID TIGHT 1'` (si dice 1", 1.5" o 1'), si no `LIQUID TIGHT`
     - `curva`, `grapa`, `condu lr`, `condu t`, `canaleta` → sus columnas respectivas.

---

## Paso 2: Generar Excel de Consumo de Materiales
Crea `CONSUMO_MATERIALES_<fecha_corte>.xlsx` (hoja `Hoja1`) con las 26 columnas estándar (`DATOS`, `CABLE UTP `, `JACKS`, `ZPLUG`, `PATCHCORE`, `FIBRA ARMADA`, `TUBERIA  3/4`, `CURVAS  3/4`, `CONDU LR`, `CONDU T`, `GRAPAS 3/4`, `CAJA 4S`, `FACE PLATE`, `CORAZA`, `LIQUID TIGHT`, `TUBERIA 2"`, `CURVAS 2"`, `UNIONES 2"`, `CANALETA PLST`, `GRAPA CH`, `TUBERIA  1'`, `CURVAS  1'`, `GRAPAS 1'`, `CORAZA 1'`, `LIQUID TIGHT 1'`, `UBICACION DE LOS PUNTOS`), filas desde la 2, fila `TOTAL` con `=SUM(...)` y fila inferior con `=B*1.5`, `=C+D`, `=G*3+N`, `=U*3+X`.

---

## Paso 3: Diligenciar Plantilla Oficial (`MNN-REG-048`)
1. Abre la plantilla (hoja `MNN-REG-048`) sin alterar logos, encabezados, colores, bordes ni estilos.
2. En **DESCRIPCION DE LAS ACTIVIDADES REALIZADAS** (a partir de la fila `20`):
   - **Si tiene materiales (> 0):**
     `[DATOS]: CABLE UTP: [X] - JACK: [X] - ZPLUG: [X] - PATCH CORD: [X] - CAJA 4X4 RAWELT: [X] - TAPA CIEGA CON EMPAQUE: [X] - TUBERIA IMC 3/4": [X] - CORAZA 3/4": [X] - LIQUID TIGHT: [X]`
     *(Nota: `CAJA 4X4 RAWELT` y `TAPA CIEGA CON EMPAQUE` toman el valor de `CAJA 4S`. Si hay otros materiales > 0 como `FIBRA ARMADA`, `TUBERIA 1'`, `CORAZA 1'` o `LIQUID TIGHT 1'`, agregarlos al final con formato `NOMBRE: CANTIDAD` separados por ` - `).*
   - **Si es solo texto sin materiales** (urgencias, retiros, fusiones, montajes): colocar únicamente `[DATOS]`.
3. **Desplazamiento dinámico de filas:** La plantilla base tiene 31 filas de actividad (`20:50`). Si hay más o menos registros, descombina y desplaza exactamente la tabla resumen, `ACTIVIDADES PENDIENTES` y el bloque completo de firmas, manteniendo todas las celdas combinadas (`A:K` en actividades, `A:I` en resumen), bordes, estilos y ajustando el alto de cada fila según el largo del texto.
4. **Tabla resumen (columna `K`):**
   - `Puntos de red`: cantidad de filas con `CABLE UTP > 0` (o valor manual indicado).
   - `CABLE F/UTP COMMSCOPE`: suma total de `CABLE UTP`.
   - `TUBERIA IMC 3/4 Y CORAZAS 3/4`: según la regla elegida por el usuario al inicio.
   - `REVIONES`, `DISPONIBILIDAD INMEDIATA`, `PERSONAL HSEQ`: valores confirmados por el usuario.

---

## Entrega
- No muestres código ni pidas pasos manuales.
- Entrega los enlaces descargables de ambos archivos:
  1. `CONSUMO_MATERIALES_<fecha_corte>.xlsx`
  2. `ACTA_PARCIAL_ENTREGA_<fecha_corte>.xlsx`
- Incluye un resumen breve de los registros filtrados por fecha y los totales calculados en la tabla resumen.
