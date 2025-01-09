from __future__ import annotations

from numbers import Number
from typing import Any, Callable, Dict, List, Tuple

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

from .column_def import ColumnType


from highlight_text import HighlightText


def create_cell(column_type: ColumnType, *args, **kwargs) -> TableCell:
    """Factory Function to create a specific TableCell depending on `column_type`.

    Args:
        column_type (ColumnType): plottable.column_def.ColumnType

    Returns:
        TableCell: plottable.cell.TableCell
    """
    if column_type is ColumnType.SUBPLOT:
        return SubplotCell(*args, **kwargs)
    elif column_type is ColumnType.STRING:
        return TextCell(*args, **kwargs)
    elif column_type is ColumnType.HIGHLIGHTTEXT:
        return HighlightTextCell(*args, **kwargs)


class Cell:
    """A cell is a rectangle defined by the lower left corner xy and it's width and height."""

    def __init__(self, xy: Tuple[float, float], width: float = 1, height: float = 1):
        """
        Args:
            xy (Tuple[float, float]): lower left corner of a rectangle
            width (float, optional): width of the rectangle cell. Defaults to 1.
            height (float, optional): height of the rectangle cell. Defaults to 1.
        """
        self.xy = xy
        self.width = width
        self.height = height

    @property
    def x(self) -> float:
        return self.xy[0]

    @property
    def y(self) -> float:
        return self.xy[1]

    def __repr__(self) -> str:
        return f"Cell({self.xy}, {self.width}, {self.height})"


class TableCell(Cell):
    """A TableCell class for a plottable.table.Table."""

    def __init__(
        self,
        xy: Tuple[float, float],
        content: Any,
        row_idx: int,
        col_idx: int,
        width: float = 1,
        height: float = 1,
        ax: mpl.axes.Axes = None,
        rect_kw: Dict[str, Any] = {},
    ):
        """
        Args:
            xy (Tuple[float, float]):
                lower left corner of a rectangle
            content (Any):
                the content of the cell
            row_idx (int):
                row index
            col_idx (int):
                column index
            width (float, optional):
                width of the rectangle cell. Defaults to 1.
            height (float, optional):
                height of the rectangle cell. Defaults to 1.
            ax (mpl.axes.Axes, optional):
                matplotlib Axes object. Defaults to None.
            rect_kw (Dict[str, Any], optional):
                keywords passed to matplotlib.patches.Rectangle. Defaults to {}.
        """

        super().__init__(xy, width, height)
        self.index = (row_idx, col_idx)
        self.content = content
        self.row_idx = row_idx
        self.col_idx = col_idx
        self.ax = ax or plt.gca()
        self.rect_kw = {
            "linewidth": 0.0,
            "edgecolor": self.ax.get_facecolor(),
            "facecolor": self.ax.get_facecolor(),
            "width": width,
            "height": height,
        }

        self.rect_kw.update(rect_kw)
        self.rectangle_patch = Rectangle(xy, **self.rect_kw)

    def draw(self):
        self.ax.add_patch(self.rectangle_patch)

    def __repr__(self) -> str:
        return (
            f"TableCell(xy={self.xy}, row_idx={self.index[0]}, col_idx={self.index[1]})"  # noqa
        )


class SubplotCell(TableCell):
    """A SubplotTableCell class for a plottable.table.Table that creates a subplot on top of
    it's rectangle patch.
    """

    def __init__(
        self,
        xy: Tuple[float, float],
        content: Any,
        row_idx: int,
        col_idx: int,
        plot_fn: Callable,
        plot_kw: Dict[str, Any] = {},
        width: float = 1,
        height: float = 1,
        ax: mpl.axes.Axes = None,
        rect_kw: Dict[str, Any] = {},
    ):
        """
        Args:
            xy (Tuple[float, float]):
                lower left corner of a rectangle
            content (Any):
                the content of the cell
            row_idx (int):
                row index
            col_idx (int):
                column index
            plot_fn (Callable):
                function that draws onto the created subplot.
            plot_kw (Dict[str, Any], optional):
                keywords for the plot_fn. Defaults to {}.
            width (float, optional):
                width of the rectangle cell. Defaults to 1.
            height (float, optional):
                height of the rectangle cell. Defaults to 1.
            ax (mpl.axes.Axes, optional):
                matplotlib Axes object. Defaults to None.
            rect_kw (Dict[str, Any], optional):
                keywords passed to matplotlib.patches.Rectangle. Defaults to {}.
        """
        super().__init__(
            xy=xy,
            width=width,
            height=height,
            content=content,
            row_idx=row_idx,
            col_idx=col_idx,
            ax=ax,
            rect_kw=rect_kw,
        )

        self._plot_fn = plot_fn
        self._plot_kw = plot_kw
        self.fig = self.ax.figure
        self.draw()

    def plot(self):
        self._plot_fn(self.axes_inset, self.content, **self._plot_kw)

    def make_axes_inset(self):
        rect_fig_coords = self._get_rectangle_bounds()
        self.axes_inset = self.fig.add_axes(rect_fig_coords)
        return self.axes_inset

    def _get_rectangle_bounds(self, padding: float = 0.2) -> List[float]:
        transformer = self.fig.transFigure.inverted()
        display_coords = self.rectangle_patch.get_window_extent()
        (xmin, ymin), (xmax, ymax) = transformer.transform(display_coords)
        y_range = ymax - ymin
        return [
            xmin,
            ymin + padding * y_range,
            xmax - xmin,
            ymax - ymin - 2 * padding * y_range,
        ]

    def draw(self):
        self.ax.add_patch(self.rectangle_patch)

    def __repr__(self) -> str:
        return f"SubplotCell(xy={self.xy}, row_idx={self.index[0]}, col_idx={self.index[1]})"  # noqa


class TextCell(TableCell):
    """A TextCell class for a plottable.table.Table that creates a text inside it's rectangle patch."""

    def __init__(
        self,
        xy: Tuple[float, float],
        content: str | Number,
        row_idx: int,
        col_idx: int,
        width: float = 1,
        height: float = 1,
        ax: mpl.axes.Axes = None,
        rect_kw: Dict[str, Any] = {},
        textprops: Dict[str, Any] = {},
        padding: float = 0.1,
    ):
        """
        Args:
            xy (Tuple[float, float]):
                lower left corner of a rectangle
            content (Any):
                the content of the cell
            row_idx (int):
                row index
            col_idx (int):
                column index
            width (float, optional):
                width of the rectangle cell. Defaults to 1.
            height (float, optional):
                height of the rectangle cell. Defaults to 1.
            ax (mpl.axes.Axes, optional):
                matplotlib Axes object. Defaults to None.
            rect_kw (Dict[str, Any], optional):
                keywords passed to matplotlib.patches.Rectangle. Defaults to {}.
            textprops (Dict[str, Any], optional):
                textprops passed to matplotlib.text.Text. Defaults to {}.
            padding (float, optional):
                Padding around the text within the rectangle patch. Defaults to 0.1.

        """
        super().__init__(
            xy=xy,
            width=width,
            height=height,
            content=content,
            row_idx=row_idx,
            col_idx=col_idx,
            ax=ax,
            rect_kw=rect_kw,
        )

        self.textprops = {"ha": "right", "va": "center"}
        self.textprops.update(textprops)
        self.ha = self.textprops["ha"]
        self.va = self.textprops["va"]
        self.padding = padding

    def draw(self):
        self.ax.add_patch(self.rectangle_patch)
        self.set_text()

    def set_text(self):
        x, y = self._get_text_xy()

        self.text = self.ax.text(x, y, str(self.content), **self.textprops)

    def _get_text_xy(self):
        x, y = self.xy

        if self.ha == "left":
            x = x + self.padding * self.width
        elif self.ha == "right":
            x = x + (1 - self.padding) * self.width
        elif self.ha == "center":
            x = x + self.width / 2
        else:
            raise ValueError(
                f"ha can be either 'left', 'center' or 'right'. You provided {self.ha}."
            )

        if self.va == "center":
            y += self.height / 2
        elif self.va == "bottom":
            # CHANGED
            y = y + (1 - self.padding) * self.height
        elif self.va == "top":
            # CHANGED
            y = y + self.padding * self.height

        return x, y

    def __repr__(self) -> str:
        return f"TextCell(xy={self.xy}, content={self.content}, row_idx={self.index[0]}, col_idx={self.index[1]})"  # noqa


class HighlightTextCell(TextCell):
    """A HighlightTextCell class for a plottable.table.Table that creates a text inside its rectangle patch."""

    def __init__(
        self,
        xy: Tuple[float, float],
        content: str | Number,
        row_idx: int,
        col_idx: int,
        width: float = 1,
        height: float = 1,
        ax: mpl.axes.Axes = None,
        rect_kw: Dict[str, Any] = {},
        textprops: Dict[str, Any] = {},
        highlight_textprops: Dict[str, Any] | None = None,
        padding: float = 0.1,
    ):
        """
        Args:
            xy (Tuple[float, float]):
                lower left corner of a rectangle
            content (Any):
                the content of the cell
            row_idx (int):
                row index
            col_idx (int):
                column index
            width (float, optional):
                width of the rectangle cell. Defaults to 1.
            height (float, optional):
                height of the rectangle cell. Defaults to 1.
            ax (mpl.axes.Axes, optional):
                matplotlib Axes object. Defaults to None.
            rect_kw (Dict[str, Any], optional):
                keywords passed to matplotlib.patches.Rectangle. Defaults to {}.
            textprops (Dict[str, Any], optional):
                textprops passed to matplotlib.text.Text. Defaults to {}.
            padding (float, optional):
                Padding around the text within the rectangle patch. Defaults to 0.1.

        """
        super().__init__(
            xy=xy,
            width=width,
            height=height,
            content=content,
            row_idx=row_idx,
            col_idx=col_idx,
            ax=ax,
            rect_kw=rect_kw,
            textprops=textprops,
            padding=padding,
        )

        self.highlight_textprops = highlight_textprops

    def set_text(self):
        x, y = self._get_text_xy()

        content = str(self.content)

        self.text = HighlightText(
            x,
            y,
            content,
            highlight_textprops=self.highlight_textprops,
            ax=self.ax,
            textalign="right",
            **self.textprops,
        )

    def __repr__(self) -> str:
        return f"HighlightTextCell(xy={self.xy}, content={self.content}, row_idx={self.index[0]}, col_idx={self.index[1]})"  # noqa


class Sequence:  # Row and Column can inherit from this
    """A Sequence of Table Cells."""

    def __init__(self, cells: List[TableCell], index: int):
        """

        Args:
            cells (List[TableCell]): List of TableCells.
            index (int): an index denoting the sequences place in a Table.
        """
        self.cells = cells
        self.index = index

    def append(self, cell: TableCell):
        """Appends another TableCell to its `cells` propery.

        Args:
            cell (TableCell): A TableCell object
        """
        self.cells.append(cell)

    def set_alpha(self, *args) -> Sequence:
        """Sets the alpha for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_alpha(*args)
        return self

    def set_color(self, *args) -> Sequence:
        """Sets the color for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_color(*args)
        return self

    def set_facecolor(self, *args) -> Sequence:
        """Sets the facecolor for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_facecolor(*args)
        return self

    def set_edgecolor(self, *args) -> Sequence:
        """Sets the edgecolor for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_edgecolor(*args)
        return self

    def set_fill(self, *args) -> Sequence:
        """Sets the fill for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_fill(*args)
        return self

    def set_hatch(self, *args) -> Sequence:
        """Sets the hatch for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_hatch(*args)
        return self

    def set_linestyle(self, *args) -> Sequence:
        """Sets the linestyle for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_linestyle(*args)
        return self

    def set_linewidth(self, *args) -> Sequence:
        """Sets the linewidth for all cells of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            cell.rectangle_patch.set_linewidth(*args)
        return self

    def set_fontcolor(self, *args) -> Sequence:
        """Sets the fontcolor for all cells texts of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            if hasattr(cell, "text"):
                cell.text.set_color(*args)
        return self

    def set_fontfamily(self, *args) -> Sequence:
        """Sets the fontfamily for all cells texts of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            if hasattr(cell, "text"):
                cell.text.set_fontfamily(*args)
        return self

    def set_fontsize(self, *args) -> Sequence:
        """Sets the fontsize for all cells texts of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            if hasattr(cell, "text"):
                cell.text.set_fontsize(*args)
        return self

    def set_ha(self, *args) -> Sequence:
        """Sets the horizontal alignment for all cells texts of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            if hasattr(cell, "text"):
                cell.text.set_ha(*args)
        return self

    def set_ma(self, *args) -> Sequence:
        """Sets the multialignment for all cells tests of the Sequence and returns self.

        Return:
            self[Sequence]: A Sequence of Cells
        """
        for cell in self.cells:
            if hasattr(cell, "text"):
                cell.text.set_ma(*args)
        return self


class Row(Sequence):
    """A Row of TableCells."""

    def __init__(self, cells: List[TableCell], index: int):
        super().__init__(cells=cells, index=index)

    def get_xrange(self) -> Tuple[float, float]:
        """Gets the xrange of the Row.

        Returns:
            Tuple[float, float]: Tuple of min and max x.
        """
        return min([cell.xy[0] for cell in self.cells]), max(
            [cell.xy[0] + cell.width for cell in self.cells]
        )

    def get_yrange(self) -> Tuple[float, float]:
        """Gets the yrange of the Row.

        Returns:
            Tuple[float, float]: Tuple of min and max y.
        """
        cell = self.cells[0]
        return cell.xy[1], cell.xy[1] + cell.height

    @property
    def x(self) -> float:
        return self.cells[0].xy[0]

    @property
    def y(self) -> float:
        return self.cells[0].xy[1]

    @property
    def height(self) -> float:
        return self.cells[0].height

    def __repr__(self) -> str:
        return f"Row(cells={self.cells}, index={self.index})"


class Column(Sequence):
    """A Column of TableCells."""

    def __init__(self, cells: List[TableCell], index: int, name: str = None):
        super().__init__(cells=cells, index=index)
        self.name = name

    def get_xrange(self) -> Tuple[float, float]:
        """Gets the xrange of the Column.

        Returns:
            Tuple[float, float]: Tuple of min and max x.
        """
        cell = self.cells[0]
        return cell.xy[0], cell.xy[0] + cell.width

    def get_yrange(self) -> Tuple[float, float]:
        """Gets the yrange of the Column.

        Returns:
            Tuple[float, float]: Tuple of min and max y.
        """
        return min([cell.xy[1] for cell in self.cells]), max(
            [cell.xy[1] + cell.height for cell in self.cells]
        )

    @property
    def x(self) -> float:
        return self.cells[0].xy[0]

    @property
    def y(self) -> float:
        return self.cells[0].xy[1]

    @property
    def width(self) -> float:
        return self.cells[0].width

    def __repr__(self) -> str:
        return f"Column(cells={self.cells}, index={self.index})"
