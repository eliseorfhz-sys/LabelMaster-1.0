# LabelMaster-1.0

Aplicación web construida con **Streamlit** para administrar un catálogo de productos y generar archivos PDF listos para impresión:

- **Etiquetas de precio** con nombre, precio y código de barras (Code128).
- **Fichas/Cupones de internet** por grupos en plantilla de impresión.

La app importa datos desde Excel, guarda información en una base local SQLite y permite búsquedas rápidas por código (incluyendo lector USB).

## Características principales

- Importación de catálogo desde archivos `.xlsx` y `.xls`.
- Normalización de columnas y precios al momento de importar.
- Almacenamiento local en `catalogo.db`.
- Consulta y filtrado de productos por departamento.
- Búsqueda por código de barras (entrada manual o escáner USB).
- Generación de PDF de etiquetas en formato carta.
- Gestión de grupos de cupones y generación de PDF por grupo.

## Requisitos

- Python 3.10 o superior recomendado.
- pip actualizado.

## Instalación

1. Clona este repositorio:

```bash
git clone https://github.com/eliseorfhz-sys/LabelMaster-1.0.git
cd LabelMaster-1.0
```

2. (Opcional, recomendado) Crea y activa un entorno virtual:

```bash
python -m venv .venv
```

En Windows (PowerShell):

```bash
.venv\Scripts\Activate.ps1
```

3. Instala dependencias:

```bash
pip install -r requirements.txt
```

## Ejecución

Inicia la aplicación con:

```bash
streamlit run app.py
```

Luego abre en tu navegador la URL que te muestre Streamlit (por ejemplo `http://localhost:8501`).

## Uso rápido

### 1) Catálogo

- En la barra lateral, carga un Excel para importar o actualizar el catálogo.
- La app espera columnas como: `Código`, `Producto`, `P. Costo`, `P. Venta`, `P. Mayoreo`, `Departamento`, `Existencia`, `Inv. Mínimo`, `Inv. Máximo`, `Tipo de Venta`, `Proveedor`.
- Puedes filtrar por departamento y buscar por código.

### 2) Etiquetas

- Escanea o escribe códigos para agregar productos a la selección.
- Define cuántas copias por producto necesitas.
- Genera y descarga `etiquetas.pdf`.

### 3) Cupones Internet

- Carga un Excel con la estructura esperada para cupones.
- Guarda grupos y genera el PDF correspondiente.
- Descarga el archivo final desde la misma interfaz.

## Estructura básica del proyecto

- `app.py`: lógica principal de la aplicación Streamlit.
- `requirements.txt`: dependencias de Python.
- `catalogo.db`: base de datos local SQLite (se crea/actualiza en ejecución).

## Notas de versión de archivos locales

Este proyecto está configurado para **no subir** al repositorio:

- Base de datos local (`*.db`).
- Archivos de Docker (`Dockerfile`, `docker-compose.yml`, `.dockerignore`).

## Próximas mejoras sugeridas

- Validaciones avanzadas de formato para archivos Excel.
- Exportación de reportes adicionales (CSV/PDF resumen).
- Historial de importaciones y auditoría de cambios.
- Ajustes visuales de plantillas de impresión por tipo de etiqueta.

## Licencia

Este proyecto está licenciado bajo la licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más detalles.
