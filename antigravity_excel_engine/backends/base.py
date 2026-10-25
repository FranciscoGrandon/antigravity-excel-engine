# Creado por Francisco Grandón Vergara
"""
Clase base abstracta para backends de Excel en Antigravity Excel Engine.
Define el contrato unificado para backends interactivos (Live COM) y de archivo (Headless OpenPyXL).
"""

import re
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Union

from antigravity_excel_engine.formatting.styles import FinancialStyleGuide


class BaseExcelBackend(ABC):
    """
    Contrato base abstracto para manipuladores y controladores de Excel.
    Soporta operaciones atómicas de celda, rangos masivos, tablas dinámicas,
    gráficos, pipelines de curación y despacho interactivo.
    """

    @abstractmethod
    def open_workbook(self, file_path: str) -> Any:
        """Abre o enlaza un libro de trabajo."""
        pass

    @abstractmethod
    def close(self) -> None:
        """Cierra el libro de trabajo o libera recursos asociados."""
        pass

    @abstractmethod
    def save(self, file_path: Optional[str] = None) -> str:
        """Guarda el libro de trabajo en disco."""
        pass

    @abstractmethod
    def get_sheet(self, sheet_name: Optional[str] = None) -> Any:
        """Obtiene un objeto de hoja de trabajo (Worksheet)."""
        pass

    @abstractmethod
    def get_cell_ranges(
        self,
        ranges: List[str],
        sheet: Optional[str] = None,
        include_formulas: bool = False,
        include_styles: bool = False
    ) -> Dict[str, Any]:
        """Obtiene valores, fórmulas y estilos de uno o varios rangos."""
        pass

    @abstractmethod
    def set_cell_range(
        self,
        start_cell: str,
        values: List[List[Any]],
        sheet: Optional[str] = None,
        formulas: Optional[List[List[str]]] = None,
        styles: Optional[Dict[str, Any]] = None,
        autofit: bool = False,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """Escribe una matriz rectangular de celdas a partir de una celda inicial."""
        pass

    @abstractmethod
    def set_cells(
        self,
        cells_map: Dict[str, Dict[str, Any]],
        sheet: Optional[str] = None,
        autofit: bool = False,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """Escribe celdas declarativas con valores, fórmulas y formatos individuales."""
        pass

    @abstractmethod
    def get_range_as_csv(
        self,
        range_str: str,
        sheet: Optional[str] = None,
        delimiter: str = ";"
    ) -> str:
        """Exporta un rango como cadena CSV."""
        pass

    @abstractmethod
    def copy_paste_range(
        self,
        src_range: str,
        dst_start_cell: str,
        src_sheet: Optional[str] = None,
        dst_sheet: Optional[str] = None,
        paste_type: str = "all"
    ) -> Dict[str, Any]:
        """Copia y pega un rango entre hojas o dentro de la misma hoja."""
        pass

    @abstractmethod
    def check_formula_errors(
        self,
        range_str: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Audita y reporta errores de fórmulas en el rango."""
        pass

    @abstractmethod
    def audit_data_quality(
        self,
        range_str: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Realiza una auditoría exhaustiva de calidad de datos en el rango."""
        pass

    def read_range(
        self,
        range_str: str,
        sheet: Optional[str] = None,
        include_formulas: bool = False,
        include_styles: bool = False
    ) -> Dict[str, Any]:
        """
        Lee los datos de un rango específico y retorna su diccionario estructurado.
        """
        res = self.get_cell_ranges(
            [range_str],
            sheet=sheet,
            include_formulas=include_formulas,
            include_styles=include_styles
        )
        return res.get(range_str, {})

    def write_range(
        self,
        start_cell: str,
        values: List[List[Any]],
        sheet: Optional[str] = None,
        formulas: Optional[List[List[str]]] = None,
        styles: Optional[Dict[str, Any]] = None,
        autofit: bool = False
    ) -> Dict[str, Any]:
        """
        Escribe una matriz rectangular de celdas a partir de una celda inicial.
        """
        return self.set_cell_range(
            start_cell=start_cell,
            values=values,
            sheet=sheet,
            formulas=formulas,
            styles=styles,
            autofit=autofit
        )

    def get_cell_value(self, cell: str, sheet: Optional[str] = None) -> Any:
        """
        Obtiene el valor atómico de una única celda.
        """
        data = self.read_range(cell, sheet=sheet)
        vals = data.get("values", [])
        if vals and len(vals) > 0 and len(vals[0]) > 0:
            return vals[0][0]
        return None

    def set_cell_value(
        self,
        cell: str,
        value: Any,
        sheet: Optional[str] = None,
        autofit: bool = False
    ) -> Dict[str, Any]:
        """
        Escribe un valor atómico en una celda individual.
        """
        return self.write_range(
            start_cell=cell,
            values=[[value]],
            sheet=sheet,
            autofit=autofit
        )

    def supports_formula_recalc(self) -> bool:
        """
        Indica si el backend soporta motor nativo de recálculo dinámico de fórmulas.
        """
        return False

    def is_busy(self) -> bool:
        """
        Indica si el backend se encuentra ocupado o bloqueado por el usuario o proceso externo.
        """
        return False

    @staticmethod
    def _calculate_placement_coordinates(
        start_col: int,
        start_row: int,
        num_cols: int,
        num_rows: int,
        placement: str = "auto",
        anchor: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Helper de cálculo de posición geométrica no invasiva anti-solapamiento.
        Si anchor está especificado (ej. 'I3'), utiliza esa celda.
        Si placement='auto':
          - si num_cols <= 8: ubica a la derecha (dejando 2 columnas libres de margen, misma fila inicial).
          - si num_cols > 8: ubica abajo (dejando 2 filas libres de margen, misma columna inicial).
        Si placement='right': a la derecha dejando 2 columnas libres.
        Si placement='bottom': abajo dejando 2 filas libres.
        """
        if anchor:
            m = re.match(r"^([A-Za-z]+)(\d+)$", anchor.strip())
            if m:
                col_letters = m.group(1).upper()
                r_num = int(m.group(2))
                c_num = 0
                for char in col_letters:
                    c_num = c_num * 26 + (ord(char) - ord('A') + 1)
                return {
                    "col": c_num,
                    "row": r_num,
                    "cell": f"{col_letters}{r_num}",
                    "placement": "custom"
                }

        plc = placement.lower() if placement else "auto"
        if plc == "auto":
            plc = "right" if num_cols <= 8 else "bottom"

        if plc == "right":
            target_col = start_col + num_cols + 2  # 2 columnas libres de margen
            target_row = start_row
        elif plc == "bottom":
            target_col = start_col
            target_row = start_row + num_rows + 2  # 2 filas libres de margen
        else:
            target_col = start_col + num_cols + 2
            target_row = start_row

        col_temp = target_col
        letters = []
        while col_temp > 0:
            col_temp, rem = divmod(col_temp - 1, 26)
            letters.append(chr(ord('A') + rem))
        col_str = "".join(reversed(letters))

        return {
            "col": target_col,
            "row": target_row,
            "cell": f"{col_str}{target_row}",
            "placement": plc
        }

    @abstractmethod
    def create_pivot_table(
        self,
        source: Union[str, Dict[str, Any]],
        rows: List[str],
        cols: Optional[List[str]] = None,
        value_field: Optional[str] = None,
        value_func: str = "sum",
        value_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT,
        dest_sheet: Optional[str] = None,
        dest_cell: str = "A3",
        table_name: Optional[str] = None,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Crea una tabla dinámica / resumen pivoteado multidimensional."""
        pass

    @abstractmethod
    def create_chart(
        self,
        source: Union[str, Dict[str, Any]],
        chart_type: str = "column_clustered",
        title: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        placement: str = "auto",
        anchor: Optional[str] = None,
        width: int = 480,
        height: int = 300,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Inserta un gráfico en la hoja destino con posicionamiento inteligente."""
        pass

    @abstractmethod
    def aggregate(
        self,
        group_by_col: str,
        metric_col: str,
        agg_func: str = "sum",
        sheet: Optional[str] = None,
        top_n: int = 10,
        ascending: bool = False
    ) -> Dict[str, Any]:
        """Agrupa y agrega una métrica según una columna categórica."""
        pass

    @abstractmethod
    def add_calculated_column(
        self,
        header: str,
        formula_template: str,
        number_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT,
        sheet: Optional[str] = None,
        autofit: bool = True,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """Inserta una columna calculada propagando fórmulas fila por fila."""
        pass

    @abstractmethod
    def apply_style_preset(
        self,
        range_str: str,
        preset: str,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Aplica un preset tipográfico y de color de FinancialStyleGuide."""
        pass

    @abstractmethod
    def create_summary_table(
        self,
        title: str,
        headers: List[str],
        rows_spec: List[Dict[str, Any]],
        dest_sheet: str,
        start_cell: str = "B2",
        parameters: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """Construye un bloque financiero resumen estructurado."""
        pass

    @abstractmethod
    def remove_duplicates(
        self,
        range_or_table: str,
        key_columns: Optional[List[str]] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Elimina registros duplicados según columnas clave."""
        pass

    @abstractmethod
    def detect_and_flag_outliers(
        self,
        column: str,
        method: str = "iqr",
        range_or_table: Optional[str] = None,
        sheet: Optional[str] = None,
        flag_style: str = "warning",
        add_comment: bool = True
    ) -> Dict[str, Any]:
        """Detecta valores atípicos y resalta celdas visualmente."""
        pass

    @abstractmethod
    def clean_table_dataset(
        self,
        target_range_or_table: str,
        rules: Dict[str, Any],
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Ejecuta un conjunto de reglas de limpieza y normalización sobre el dataset."""
        pass

    @abstractmethod
    def curate_pipeline(
        self,
        target_range_or_table: str,
        pipeline_spec: Dict[str, Any],
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Ejecuta un pipeline declarativo completo de curación y transformación de datos."""
        pass

    @abstractmethod
    def status(self) -> Dict[str, Any]:
        """Reporta el estado operativo y diagnóstico del backend."""
        pass

    @abstractmethod
    def wait_until_ready(
        self,
        timeout_sec: float = 10.0,
        retry_interval: float = 0.25
    ) -> Dict[str, Any]:
        """Espera activamente hasta que el backend esté listo para aceptar comandos."""
        pass

    @abstractmethod
    def dispatch_com(
        self,
        target: str,
        member: str,
        args: Optional[List[Any]] = None,
        kwargs: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Despacha dinámicamente llamadas sobre el modelo de objetos COM."""
        pass

    @abstractmethod
    def inspect_workbook(
        self,
        sheet: Optional[str] = None,
        include_charts: bool = True
    ) -> Dict[str, Any]:
        """Inspecciona metadatos globales del libro y estructura detallada por hoja."""
        pass

    @abstractmethod
    def inspect_charts(
        self,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Inspecciona todos los gráficos en la hoja especificada o en el libro completo."""
        pass

    @abstractmethod
    def extend_chart_series(
        self,
        to_col: Optional[str] = None,
        to_row: Optional[int] = None,
        sheet: Optional[str] = None,
        charts: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """Extiende las series y rangos de categorías de los gráficos indicados."""
        pass

    @abstractmethod
    def update_chart_series(
        self,
        chart_name: str,
        series_index: int,
        x_range: Optional[str] = None,
        y_range: Optional[str] = None,
        name_ref: Optional[str] = None,
        formula: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Modifica directamente los parámetros o la fórmula de una serie de gráfico."""
        pass

    @abstractmethod
    def create_table(
        self,
        range_str: str,
        table_name: Optional[str] = None,
        style: str = "TableStyleMedium9",
        has_headers: bool = True,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Convierte un rango libre en tabla estructurada (ListObject)."""
        pass

    @abstractmethod
    def list_tables(
        self,
        sheet: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Lista las tablas estructuradas existentes."""
        pass

    @abstractmethod
    def resize_table(
        self,
        table_name: str,
        new_range: str,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Redimensiona una tabla estructurada existente."""
        pass

    def inspect_table_schema(
        self,
        table_or_range: str,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """Inspecciona y describe exhaustivamente el esquema de una tabla o rango."""
        from antigravity_excel_engine.operations.grounding import inspect_table_schema
        return inspect_table_schema(self, table_or_range=table_or_range, sheet=sheet)

    def get_ai_grounding_context(
        self,
        table_or_range: str,
        sheet: Optional[str] = None,
        format: str = "json"
    ) -> Union[str, Dict[str, Any]]:
        """Serializa el esquema para inyección de grounding en LLMs."""
        from antigravity_excel_engine.operations.grounding import get_ai_grounding_context
        return get_ai_grounding_context(self, table_or_range=table_or_range, sheet=sheet, format=format)


