# ayuda-colegio

Ordena los archivos de casos en **inicial / primaria / secundaria** y mantiene un único Excel (`casos-niños.xlsx`).
Funciona solo en tu Mac. **No envía ninguna información a internet.**

## Cómo usarla (sin saber nada de computación)
1. Descarga `ayuda-colegio-Mac.zip` desde la página de *Releases*, haz doble clic para abrirlo.
2. **Primera vez:** clic derecho (o Control + clic) sobre **ayuda-colegio** → **Abrir** → **Abrir**. Solo se hace una vez.
   - Si en macOS 15 (Sequoia) o más nuevo eso no basta: abre **Ajustes del Sistema** → **Privacidad y seguridad**, baja hasta el aviso de ayuda-colegio y pulsa **Abrir de todos modos**.
3. **Paso 1:** pulsa *Examinar* y elige la carpeta donde están los archivos.
4. **Paso 2:**
   - **Ordenar archivos**: mueve cada archivo a su carpeta. Lo que no se pueda clasificar va a `REVISAR_A_MANO`.
   - **Actualizar Excel**: crea el Excel la primera vez y luego lo actualiza con lo nuevo.
5. Si el Excel está abierto, la app avisará: ciérralo y vuelve a intentar.

## Reglas
- Nunca se borra ni se sobrescribe ningún archivo.
- El **Estado** (Derivado / No derivado) y lo que escribas a mano en el Excel no se pisa.
- El Estado **no cambia solo** si después cambia el contenido del archivo: si quieres otro Estado, cámbialo tú en el Excel.
- Si no se pudo leer un archivo, el Estado queda **vacío** (ni Derivado ni No derivado): revísalo a mano.
- Las columnas o colores que agregues a la derecha o en las celdas de las hojas de casos se conservan; aun así, evita cambiar los títulos de la fila 1.
- Los archivos cuyo nombre empieza con «Informe» no se leen ni entran al Excel.
- Si un dato no aparece en el documento, queda vacío. No se inventa nada.
- Documentos escaneados: la app avisa y quedan anotados como «LEER MANUALMENTE».

## Para quien mantiene el código
Este repositorio contiene **solo código**: el `.gitignore` bloquea Excel, Word, PDF y carpetas de casos.
Para publicar una versión: crear una etiqueta `vX.Y` (ej. `v1.0`) y GitHub compila la app Mac y la adjunta al *Release*.
