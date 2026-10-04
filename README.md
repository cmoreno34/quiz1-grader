# Corrector del Quiz 1 (BTM-2500)

Página web estática (GitHub Pages) que corrige el Quiz 1 de Excel con un agente de Claude.
**Todo se procesa en el navegador**: los Excel de los alumnos no se suben a ningún servidor;
a la API de Anthropic solo se envía, con la clave del profesor, el informe de evidencias de
cada alumno.

## Uso

1. Canvas ▸ tarea ▸ **Download Submissions** → un `.zip`.
2. Abre la página, arrastra el zip, pega tu clave `sk-ant-…` (se guarda **cifrada** con tu
   contraseña, AES-256 + PBKDF2, solo en tu navegador) y pulsa **Corregir con el agente**.
3. Revisa la tabla (comentario de cada pregunta al pasar el ratón, justificación completa al
   hacer clic).
4. Descarga **el zip con todos los Excel corregidos** → Canvas ▸ tarea ▸ *Re-Upload
   Submissions*; y el **CSV de notas** → Canvas ▸ Grades ▸ *Import*.

Cada Excel devuelto es el original del alumno: solo cambia la tabla de corrección (nota 0–10
por pregunta, comentario con el porqué justo al lado y GRADE /100).

## Qué hay dentro

| Archivo | Qué hace |
|---|---|
| `js/xlsx.js` | Lee los .xlsx (fórmulas, valores, formatos) |
| `js/engine.js` | Motor de evidencias: recalcula cada respuesta desde los datos, clasifica cada celda (fórmula / a mano / error arrastrado / fila desplazada / error típico…) |
| `js/agent.js` | Agente Claude (`claude-opus-5-5`): nota y justificación por pregunta, calibración de la clase |
| `js/writer.js` | Escribe la tabla de corrección en el Excel original |
| `js/rubric.js` | Rúbrica (sin respuestas: se calculan al vuelo con los datos de cada clase) |

Ponle a la clave un **límite de gasto** (console.anthropic.com ▸ Workspaces ▸ Limits).
