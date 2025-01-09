from collections.abc import Callable, Sequence
from numbers import Number
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from plottable import ColumnDefinition, Table
from plottable.formatters import decimal_to_percent
from plottable.plots import bar

plt.rcParams["svg.fonttype"] = "none"


PATH = Path(r"G:\Drive partagés\CL_Chanel APAC\2024 PPA\4.Excels")
assert PATH.exists()

tabs = {
    "recruitment": "SAS 12.1 RECRUT REACH SSCAT",
    "conversion": "SAS 12.1b Focus converties",
    "generation": "SAS 12.2 RECRUT REACH SSCAT GYZ",
    "labels_mappings": "Params",
}

APAC_file = PATH / "2024 - Chanel PPA Asia - Global APAC - V0.5.xlsx"
CHN_file = PATH / "Par marché" / "2024 - Chanel PPA Asia - China - V0.2.xlsx"


def load_data_from_files(region, market, tabs) -> dict:
    region_data = _read_data(path=region, tab=tabs["recruitment"], skiprows=1)
    market_data = _read_data(path=market, tab=tabs["recruitment"], skiprows=1)

    label_mapping = (
        _read_data(
            path=market,
            tab=tabs["labels_mappings"],
            skiprows=1,
            usecols=[6, 7],
            header=None,
            names=["micro_categorie", "chart"],
        )
        .dropna()
        .drop_duplicates(subset="micro_categorie", keep="first")
        .set_index("micro_categorie")
        .squeeze()
    )

    return {
        "region": region_data,
        "market": market_data,
        "label_mapping": label_mapping,
    }


def _read_data(path, tab, **kwargs) -> pd.DataFrame:
    return (
        pd.read_excel(path, sheet_name=tab, **kwargs).replace(",", pd.NA)
        # .pipe(col_types_on_substr, col_types=SUBSTR_TYPES)
    )


def pick_colour(val: float):
    if val < -0.05:
        return "darkred"
    elif val > 0.05:
        return "darkgreen"
    else:
        return "dimgray"


def apply_colour(val, colour_fn: Callable[[float], str], strfmt: str = "+.0%"):
    if isinstance(val, Number):
        return f'<{val:{strfmt}}::{{"color": "{colour_fn(val)}"}}>\n'

    return val


def combine_cols_to_multiline(table, cols: Sequence, mapfn: Callable):
    return table[cols].map(mapfn).sum(axis=1)
    # return table[cols].map(lambda v: apply_colour(v, pick_colour)).sum(axis=1)


def kdollar(x):
    return f"${x/1000:,.1f}k"


def percentage_points(x):
    return f"{x:+.0f}ₚₚ"


def pct(x):
    return f"{x:+.0%}"


def main():
    data = load_data_from_files(region=APAC_file, market=CHN_file, tabs=tabs)

    first_purch_to_cols = [
        "ca_prem_ach_only_ty",
        "ca_prem_ach_xsell_ty",
        "ca_prem_ach_ty",
    ]
    first_purch_to_cols_evol = [
        "evol_ca_prem_ach_only",
        "evol_ca_prem_ach_xsell",
        "evol_ca_prem_ach",
    ]

    data["market"] = data["market"].assign(
        first_purch_to=combine_cols_to_multiline(
            data["market"],
            cols=first_purch_to_cols,
            mapfn=lambda v: "<${:,.1f}k::{{}}>\n".format(v / 1000),
        ).str.strip(),
        evol_first_purch_to=combine_cols_to_multiline(
            data["market"],
            cols=first_purch_to_cols_evol,
            mapfn=lambda v: apply_colour(v, pick_colour),
        ).str.strip(),
    )

    CHN_HB_2023 = (
        data["market"]
        .query("macro_categorie == 'A_HANDBAG' & nom_micro_categorie == 'sous_cat_hb'")
        .loc[
            :,
            [
                "micro_categorie",
                "first_purch_to",
                "evol_first_purch_to",
                "share_of_entry_ty",
                "evol_share_of_entry",
                "recr_index_ty",
                "evol_recr_index",
                "repeat_rate_ty",
                "evol_repeat_rate",
                "repeat_to_ty",
                "evol_repeat_to",
            ],
        ]
        .set_index("micro_categorie")
    )

    col_defs = [
        ColumnDefinition(name="micro_categorie", title=" ", width=1.25),
        ColumnDefinition(
            name="first_purch_to",
            title="",
            group="Entry T.O.",
            width=0.75,
            textprops={},
        ),
        ColumnDefinition(
            name="evol_first_purch_to",
            title="vs LY.",
            group="Entry T.O.",
            width=0.75,
            textprops={},
        ),
        ColumnDefinition(
            "share_of_entry_ty",
            group="Share of entry",
            title="",
            width=1.25,
            plot_fn=bar,
            plot_kw={
                "annotate": True,
                "height": 0.95,
                "lw": 0.0,
                "xlim": (-0.1, 0.5),  # TODO dynamic max
                "formatter": decimal_to_percent,
                "color": "#d2b496",
            },
        ),
        ColumnDefinition(
            name="evol_share_of_entry",
            title="vs LY.",
            group="Share of entry",
            width=0.75,
            formatter=percentage_points,
            textprops={"family": "sans-serif"},
        ),
        ColumnDefinition(
            name="recr_index_ty",
            title="",
            group="Recruitment Index",
            width=0.75,
            # formatter=pct,
        ),
        ColumnDefinition(
            name="evol_recr_index",
            title="vs LY.",
            group="Recruitment Index",
            width=0.75,
            formatter=pct,
        ),
        ColumnDefinition(
            "repeat_rate_ty",
            title=" ",
            group="Repeat Rate",
            width=1.25,
            plot_fn=bar,
            plot_kw={
                "annotate": True,
                "height": 0.95,
                "lw": 0.0,
                "xlim": (-0.1, 0.7),  # TODO dynamic max
                "formatter": decimal_to_percent,
                "color": "#d2b496",
            },
        ),
        ColumnDefinition(
            name="evol_repeat_rate",
            title="vs LY.",
            group="Repeat Rate",
            formatter=percentage_points,
            textprops={"family": "sans-serif"},
            width=0.75,
        ),
        ColumnDefinition(
            name="repeat_to_ty", title=" ", group="Repeat TO", formatter=kdollar
        ),
        ColumnDefinition(
            name="evol_repeat_to",
            title="vs LY.",
            group="Repeat TO",
            formatter=pct,
            width=0.75,
        ),
    ]

    fig, ax = plt.subplots(figsize=(15, 7))
    plt.rcParams["font.family"] = "Century Gothic"
    fig.tight_layout()
    # ax.grid(linestyle="--")

    tab = Table(
        CHN_HB_2023,
        column_definitions=col_defs,
        col_label_divider=False,
        col_label_cell_kw={"height": 0.5},
        row_dividers=False,
    )
    tab.col_label_row.set_fontfamily("Century Gothic")
    tab.col_label_row.set_fontsize(8)
    # ax.axis("on")
    # ax.set_ylim(top=-1)
    # print(ax.get_ylim())

    plt.show()


if __name__ == "__main__":
    main()
