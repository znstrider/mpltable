"""Module containing Table Class to plot matplotlib tables."""

from __future__ import annotations

from numbers import Number
from typing import Any, Callable, Dict, List, Tuple

import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd

from .cell import Column, Row, SubplotCell, TextCell, create_cell
from .column_def import ColumnDefinition, ColumnType
from .font import contrasting_font_color
from .formatters import apply_formatter
from .helpers import _replace_lw_key


class Table:
    """Class to plot a beautiful matplotlib table.

    Args:
        df (pd.DataFrame):
            A pandas DataFrame with your table data
        ax (mpl.axes.Axes, optional):
            matplotlib axes. Defaults to None.
        index_col (str, optional):
            column to set as the DataFrame index. Defaults to None.
        columns (List[str], optional):
            columns to use. If None defaults to all columns.
        column_definitions (List[plottable.column_def.ColumnDefinition], optional):
            ColumnDefinitions for columns that should be styled. Defaults to None.
        textprops (Dict[str, Any], optional):
            textprops are passed to each TextCells matplotlib.pyplot.text. Defaults to {}.
        cell_kw (Dict[str, Any], optional):
            cell_kw are passed to to each cells matplotlib.patches.Rectangle patch.
            Defaults to {}.
        col_label_cell_kw (Dict[str, Any], optional):
            col_label_cell_kw are passed to to each ColumnLabels cells
            matplotlib.patches.Rectangle patch. Defaults to {}.
        col_label_divider (bool, optional):
            Whether to plot a divider line below the column labels. Defaults to True.
        col_label_divider_kw (Dict[str, Any], optional):
            col_label_divider_kw are passed to plt.plot. Defaults to {}.
        footer_divider (bool, optional):
            Whether to plot a divider line below the table. Defaults to False.
        footer_divider_kw (Dict[str, Any], optional):
            footer_divider_kw are passed to plt.plot. Defaults to {}.
        row_dividers (bool, optional):
            Whether to plot divider lines between rows. Defaults to True.
        row_divider_kw (Dict[str, Any], optional):
            row_divider_kw are passed to plt.plot. Defaults to {}.
        column_border_kw (Dict[str, Any], optional):
            column_border_kw are passed to plt.plot. Defaults to {}.
        even_row_color (str | Tuple, optional):
            facecolor of the even row cell's patches. Top Row has an even (0) index.
        odd_row_color (str | Tuple, optional):
            facecolor of the even row cell's patches. Top Row has an even (0) index.

    Examples
    --------

    >>> import matplotlib.pyplot as plt
    >>> import numpy as np
    >>> import pandas as pd
    >>>
    >>> from plottable import Table
    >>>
    >>> d = pd.DataFrame(np.random.random((10, 5)), columns=["A", "B", "C", "D", "E"]).round(2)
    >>> fig, ax = plt.subplots(figsize=(5, 8))
    >>> tab = Table(d)
    >>>
    >>> plt.show()

    """

    def __init__(
        self,
        df: pd.DataFrame,
        ax: mpl.axes.Axes = None,
        index_col: str = None,
        columns: List[str] = None,
        column_definitions: List[ColumnDefinition] = None,
        textprops: Dict[str, Any] = {},
        cell_kw: Dict[str, Any] = {},
        col_label_cell_kw: Dict[str, Any] = {},
        col_label_divider: bool = True,
        footer_divider: bool = False,
        row_dividers: bool = True,
        row_divider_kw: Dict[str, Any] = {},
        col_label_divider_kw: Dict[str, Any] = {},
        footer_divider_kw: Dict[str, Any] = {},
        column_border_kw: Dict[str, Any] = {},
        even_row_color: str | Tuple = None,
        odd_row_color: str | Tuple = None,
    ):
        if index_col is not None:
            if index_col in df.columns:
                df = df.set_index(index_col)
            else:
                raise KeyError(
                    f"The index_col `{index_col}` you provided does not exist."
                )
            self.index_col = index_col

        if columns is not None:
            self.df = df[columns]
        else:
            self.df = df

        if self.df.index.name is None:
            self.df.index.name = "index"

        self.ax = ax or plt.gca()
        self.figure = self.ax.figure

        self.n_rows, self.n_cols = self.df.shape

        self.column_names = [self.df.index.name] + list(self.df.columns)
        self._init_column_definitions(column_definitions)

        self.column_name_to_idx = {col: i for i, col in enumerate(self.column_names)}
        self.cell_kw = cell_kw
        self.col_label_cell_kw = col_label_cell_kw
        self.textprops = textprops
        if "ha" not in textprops:
            self.textprops.update({"ha": "right"})

        self.cells = {}
        self._init_columns()
        self._init_rows()
        self.ax.axis("off")

        self.set_alternating_row_colors(even_row_color, odd_row_color)
        self._apply_column_formatters()
        self._apply_column_cmaps()
        self._apply_column_text_cmaps()

        self._plot_col_group_labels()

        ymax = self.n_rows

        if col_label_divider:
            self._plot_col_label_divider(**col_label_divider_kw)
        if footer_divider:
            self._plot_footer_divider(**footer_divider_kw)
        if row_dividers:
            self._plot_row_dividers(**row_divider_kw)
        self._plot_column_borders(**column_border_kw)

        self.ax.set_xlim(-0.025, sum(self._get_column_widths()) + 0.025)
        if self.col_group_cells:
            miny = -2
        else:
            miny = -1

        self.ax.set_ylim(miny - 0.025, ymax + 0.05)
        self.ax.invert_yaxis()

        self._make_subplots()

    def _init_column_definitions(
        self, column_definitions: List[ColumnDefinition]
    ) -> None:
        """Initializes the Tables ColumnDefinitions.

        Args:
            column_definitions (List[ColumnDefinition]):
                List of ColumnDefinitions
        """
        if column_definitions is not None:
            self.column_definitions = {
                _def.name: _def._as_non_none_dict() for _def in column_definitions
            }
        else:
            self.column_definitions = {}
        for col in self.column_names:
            if col not in self.column_definitions:
                self.column_definitions[col] = ColumnDefinition(
                    name=col
                )._as_non_none_dict()

    def _get_column_titles(self) -> List[str]:
        """Returns a List of Column Titles.

        Returns:
            List[str]: List of Column Titles
        """
        return [
            self.column_definitions[col].get("title", col) for col in self.column_names
        ]

    def _get_col_groups(self) -> set[str]:
        """Gets the column_groups from the ColumnDefinitions.

        Returns:
            set[str]: a set of column group names
        """
        return set(
            _dict.get("group")
            for _dict in self.column_definitions.values()
            if _dict.get("group") is not None
        )

    def _get_non_group_colnames(self) -> set[str]:
        """Gets the column_names that have no column_group.

        Returns:
            set[str]: a set of column names
        """
        return set(
            _dict.get("name")
            for _dict in self.column_definitions.values()
            if _dict.get("group") is None
        )

    def _plot_col_group_labels(self) -> None:
        """Plots the column group labels."""
        col_groups = self._get_col_groups()

        self.col_group_cells = {}

        for group in col_groups:
            columns = [
                self.columns[colname]
                for colname, _dict in self.column_definitions.items()
                if _dict.get("group") == group
            ]
            x_min = min(col.get_xrange()[0] for col in columns)
            x_max = max(col.get_xrange()[1] for col in columns)
            dx = x_max - x_min

            # CHANGED Substract the group label (default) height (1) to the `y` coordinates to position properly
            y = 0 - self.col_label_row.height - 1

            textprops = self.textprops.copy()
            textprops.update({"ha": "center", "va": "bottom"})

            self.col_group_cells[group] = TextCell(
                xy=(x_min, y),
                content=group,
                row_idx=y,
                col_idx=columns[0].index,
                width=x_max - x_min,
                height=1,
                ax=self.ax,
                textprops=textprops,
            )
            self.col_group_cells[group].draw()
            self.ax.plot(
                [x_min + 0.05 * dx, x_max - 0.05 * dx],
                # CHANGED Add height (1) to have the border at the botton of the Rectangle patch of the group label
                [y + 1, y + 1],
                lw=0.2,
                color=plt.rcParams["text.color"],
            )

    def _plot_col_label_divider(self, **kwargs):
        """Plots a line below the column labels."""
        COL_LABEL_DIVIDER_KW = {"color": plt.rcParams["text.color"], "linewidth": 1}
        if "lw" in kwargs:
            kwargs["linewidth"] = kwargs.pop("lw")
        COL_LABEL_DIVIDER_KW.update(kwargs)
        self.COL_LABEL_DIVIDER_KW = COL_LABEL_DIVIDER_KW

        x0, x1 = self.rows[0].get_xrange()
        self.ax.plot(
            [x0, x1],
            [0, 0],
            **COL_LABEL_DIVIDER_KW,
        )

    def _plot_footer_divider(self, **kwargs):
        """Plots a line below the bottom TableRow."""
        FOOTER_DIVIDER_KW = {"color": plt.rcParams["text.color"], "linewidth": 1}
        if "lw" in kwargs:
            kwargs["linewidth"] = kwargs.pop("lw")
        FOOTER_DIVIDER_KW.update(kwargs)
        self.FOOTER_DIVIDER_KW = FOOTER_DIVIDER_KW

        x0, x1 = list(self.rows.values())[-1].get_xrange()
        y = len(self.df)
        self.ax.plot([x0, x1], [y, y], **FOOTER_DIVIDER_KW)

    def _plot_row_dividers(self, **kwargs):
        """Plots lines between all TableRows."""
        ROW_DIVIDER_KW = {
            "color": plt.rcParams["text.color"],
            "linewidth": 0.2,
        }
        kwargs = _replace_lw_key(kwargs)
        ROW_DIVIDER_KW.update(kwargs)

        for idx, row in list(self.rows.items())[1:]:
            x0, x1 = row.get_xrange()

            self.ax.plot([x0, x1], [idx, idx], **ROW_DIVIDER_KW)

    def _plot_column_borders(self, **kwargs):
        """Plots lines between all TableColumns where "border" is defined."""
        COLUMN_BORDER_KW = {"linewidth": 1, "color": plt.rcParams["text.color"]}

        kwargs = _replace_lw_key(kwargs)
        COLUMN_BORDER_KW.update(kwargs)

        for name, _def in self.column_definitions.items():
            if "border" in _def:
                col = self.columns[name]

                y0, y1 = col.get_yrange()

                if "l" in _def["border"].lower() or _def["border"].lower() == "both":
                    x = col.get_xrange()[0]
                    self.ax.plot([x, x], [y0, y1], **COLUMN_BORDER_KW)

                if "r" in _def["border"].lower() or _def["border"].lower() == "both":
                    x = col.get_xrange()[1]
                    self.ax.plot([x, x], [y0, y1], **COLUMN_BORDER_KW)

    def _init_columns(self):
        """Initializes the Tables columns."""
        self.columns = {}
        for idx, name in enumerate(self.column_names):
            self.columns[name] = Column(index=idx, cells=[], name=name)

    def _init_rows(self):
        """Initializes the Tables Rows."""
        self.rows = {}
        for idx, values in enumerate(self.df.to_records()):
            self.rows[idx] = self._get_row(idx, values)

        self.col_label_row = self._get_col_label_row(-1, self._get_column_titles())

    def get_column(self, name: str) -> Column:
        """Gets a Column by its column_name.

        Args:
            name (str): the column_name in the df.

        Returns:
            Column: A Column of the Table
        """
        return self.columns[name]

    def get_column_by_index(self, index: int) -> Column:
        """Gets a Column by its numeric index.

        Args:
            index (int): numeric index

        Returns:
            Column: A Column of the Table
        """
        return list(self.columns.values())[index]

    def get_even_rows(self) -> List[Row]:
        return list(self.rows.values())[::2]

    def get_odd_rows(self) -> List[Row]:
        return list(self.rows.values())[1::2]

    def set_alternating_row_colors(
        self, color: str | Tuple[float] = None, color2: str | Tuple[float] = None
    ) -> Table:
        """Sets the color of even row's rectangle patches to `color`.

        Args:
            color (str): color recognized by matplotlib for the even rows 0 ...
            color2 (str): color recognized by matplotlib for the odd rows 1 ...

        Returns:
            Table: plottable.table.Table
        """
        if color is not None:
            for row in self.get_even_rows():
                row.set_facecolor(color)

        if color2 is not None:
            for row in self.get_odd_rows():
                row.set_facecolor(color2)

        return self

    def _get_column_widths(self):
        """Gets the Column Widths."""
        return [
            self.column_definitions[col].get("width", 1) for col in self.column_names
        ]

    def _get_col_label_row(self, idx: int, content: List[str | Number]) -> Row:
        """Creates the Column Label Row.

        Args:
            idx (int): index of the Row
            content (List[str  |  Number]): content that is plotted as text.

        Returns:
            Row: Column Label Row
        """
        widths = self._get_column_widths()

        if "height" in self.col_label_cell_kw:
            height = self.col_label_cell_kw["height"]
        else:
            height = 1

        x = 0

        row = Row(cells=[], index=idx)

        for col_idx, (colname, width, _content) in enumerate(
            zip(self.column_names, widths, content)
        ):
            col_def = self.column_definitions[colname]
            textprops = self._get_column_textprops(col_def)

            # don't apply bbox around text in header
            if "bbox" in textprops:
                textprops.pop("bbox")

            cell = create_cell(
                column_type=ColumnType.STRING,
                xy=(
                    x,
                    idx + 1 - height,
                ),  # if height is different from 1 we need to adjust y
                content=_content,
                row_idx=idx,
                col_idx=col_idx,
                width=width,
                height=height,
                rect_kw=self.col_label_cell_kw,
                textprops=textprops,
                ax=self.ax,
            )

            row.append(cell)
            cell.draw()

            x += width

        return row

    def _get_subplot_cells(self) -> Dict[Tuple[int, int], SubplotCell]:
        return {
            key: cell
            for key, cell in self.cells.items()
            if isinstance(cell, SubplotCell)
        }

    def _make_subplots(self) -> None:
        self.subplots = {}
        for key, cell in self._get_subplot_cells().items():
            self.subplots[key] = cell.make_axes_inset()
            self.subplots[key].axis("off")
            cell.draw()
            cell.plot()

    def _get_column_textprops(self, col_def: ColumnDefinition) -> Dict[str, Any]:
        textprops = self.textprops.copy()
        column_textprops = col_def.get("textprops", {})
        textprops.update(column_textprops)
        textprops["multialignment"] = textprops["ha"]

        return textprops

    def _get_row(self, idx: int, content: List[str | Number]) -> Row:
        widths = self._get_column_widths()

        x = 0

        row = Row(cells=[], index=idx)

        for col_idx, (colname, width, _content) in enumerate(
            zip(self.column_names, widths, content)
        ):
            col_def = self.column_definitions[colname]

            if "plot_fn" in col_def:
                plot_fn = col_def.get("plot_fn")
                plot_kw = col_def.get("plot_kw", {})

                cell = create_cell(
                    column_type=ColumnType.SUBPLOT,
                    xy=(x, idx),
                    content=_content,
                    plot_fn=plot_fn,
                    plot_kw=plot_kw,
                    row_idx=idx,
                    col_idx=col_idx,
                    width=width,
                    rect_kw=self.cell_kw,
                    ax=self.ax,
                )

            else:
                textprops = self._get_column_textprops(col_def)

                # FIXME should pass `highlight_textprops` to the constructor
                cell = create_cell(
                    # column_type=ColumnType.STRING,
                    column_type=ColumnType.HIGHLIGHTTEXT,
                    xy=(x, idx),
                    content=_content,
                    row_idx=idx,
                    col_idx=col_idx,
                    width=width,
                    rect_kw=self.cell_kw,
                    textprops=textprops,
                    ax=self.ax,
                )

            row.append(cell)
            self.columns[colname].append(cell)
            self.cells[(idx, col_idx)] = cell
            cell.draw()

            x += width

        return row

    def _apply_column_formatters(self) -> None:
        for colname, _dict in self.column_definitions.items():
            formatter = _dict.get("formatter")
            if formatter is None:
                continue

            for cell in self.columns[colname].cells:
                if not hasattr(cell, "text"):
                    continue

                formatted = apply_formatter(formatter, cell.content)
                cell.text.set_text(formatted)

    def _apply_column_cmaps(self) -> None:
        for colname, _dict in self.column_definitions.items():
            cmap_fn = _dict.get("cmap")
            if cmap_fn is None:
                continue

            for cell in self.columns[colname].cells:
                if not isinstance(cell.content, Number):
                    continue

                if ("bbox" in _dict.get("textprops")) & hasattr(cell, "text"):
                    cell.text.set_bbox(
                        {
                            "color": cmap_fn(cell.content),
                            **_dict.get("textprops").get("bbox"),
                        }
                    )
                else:
                    cell.rectangle_patch.set_facecolor(cmap_fn(cell.content))

    def _apply_column_text_cmaps(self) -> None:
        for colname, _dict in self.column_definitions.items():
            cmap_fn = _dict.get("text_cmap")
            if cmap_fn is None:
                continue

            for cell in self.columns[colname].cells:
                if isinstance(cell.content, Number) & hasattr(cell, "text"):
                    cell.text.set_color(cmap_fn(cell.content))

    def autoset_fontcolors(
        self, fn: Callable = None, colnames: List[str] = None, **kwargs
    ) -> Table:
        """Sets the fontcolor of each table cell based on the facecolor of its rectangle patch.

        Args:
            fn (Callable, optional):
                Callable that takes the rectangle patches facecolor as
                rgba-value as argument.
                Defaults to plottable.font.contrasting_font_color if fn is None.
            colnames (List[str], optional):
                columns to apply the function to
            kwargs are passed to fn.

        Returns:
            plottable.table.Table
        """

        if fn is None:
            fn = contrasting_font_color
            if "thresh" not in kwargs:
                kwargs.update({"thresh": 150})

        if colnames is not None:
            cells = []
            for col in colnames:
                cells.extend(self.get_column(col).cells)

        else:
            cells = self.cells.values()

        for cell in cells:
            if hasattr(cell, "text"):
                text_bbox = cell.text.get_bbox_patch()

                if text_bbox:
                    bg_color = text_bbox.get_facecolor()
                else:
                    bg_color = cell.rectangle_patch.get_facecolor()

                textcolor = fn(bg_color, **kwargs)
                cell.text.set_color(textcolor)

        return self
