# Sistema de Compras

Sistema web para cruzar Requisiciones de Compra con Órdenes de Compra y generar el archivo final consolidado.

## Requisitos

- Python 3.10 o superior (recomendado: Python 3.12)
- Windows 10/11

## Cómo instalar y ejecutar

### Opción 1 — Doble clic (más fácil)
1. Extrae la carpeta `sistema_compras` donde quieras
2. Haz doble clic en **`iniciar.bat`**
3. Se abre el navegador automáticamente en `http://localhost:8000`

### Opción 2 — Terminal / PowerShell
```
cd sistema_compras
python -m venv venv312
venv312\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
Luego abre `http://localhost:8000` en el navegador.

---

## Cómo usar el sistema

1. **Cargar archivos**: Arrastra o selecciona los dos Excel:
   - `BASE_DE_DATOS_COMPRAS.xlsx` → zona izquierda (Requisiciones)
   - `COMPRAS_POR_FACTURAR__ISP.xlsx` → zona derecha (Órdenes de Compra)

2. **Procesar**: Haz clic en **"Procesar y Cruzar"**. El sistema:
   - Lee las requisiciones
   - Lee las OC de todas las hojas (OC, OC 2, OC ANABEL)
   - Normaliza los estatus (unifica SURTIDO / surtido / Surtido / etc.)
   - Hace el match por número de Orden de Compra
   - Muestra el resultado en la tabla

3. **Filtrar**: Usa la barra de búsqueda y los selectores para filtrar por:
   - Texto libre (req, usuario, descripción, proveedor)
   - Match (Con OC / Sin OC / No encontrada)
   - Estatus OC (Surtido / Parcial / Por surtir)
   - Mes

4. **Exportar**: Haz clic en **"Descargar Excel Consolidado"**. Se genera:
   - Hoja `CONSOLIDADO` con todas las filas y colores por estatus
   - Hoja `RESUMEN` con totales
   - El archivo queda guardado en la carpeta `historial/` con fecha y hora

5. **Historial**: En el panel izquierdo puedes ver los últimos 5 archivos generados.

---

## Lógica del cruce (match)

| Resultado | Significado |
|-----------|-------------|
| ✓ CON OC | La requisición tiene OC asignada y se encontró en el archivo de facturación |
| SIN OC | La requisición no tiene número de OC registrado |
| ⚠ NO ENCONTRADA | La requisición tiene OC pero no se encontró en las hojas de facturación |

## Normalización de estatus

El sistema unifica automáticamente estas variantes en el archivo de OC:
- SURTIDO, surtido, Surtido, SURTIDO (con espacios) → **SURTIDO**
- PARCIAL, Parcial, parcial → **PARCIAL**
- POR SURTIR, Por Surtir, poR SURTIR → **POR SURTIR**
- CANCELADA → **CANCELADA**

---

## Estructura del proyecto

```
sistema_compras/
├── main.py              ← Servidor FastAPI (toda la lógica)
├── requirements.txt     ← Dependencias Python
├── iniciar.bat          ← Script de inicio Windows
├── templates/
│   └── index.html       ← Interfaz web completa
├── static/              ← Archivos estáticos (CSS/JS si se agregan)
├── uploads/             ← Archivos Excel cargados (temporal)
└── historial/           ← Exportaciones guardadas + log JSON
```
