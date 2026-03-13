from __future__ import annotations

import io
import math
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


DEFAULT_PLANTS = ["1000", "1100", "2000", "2100"]
DEMAND_PATTERNS = [
    ("stable", 0.35),
    ("volatile", 0.20),
    ("seasonal", 0.15),
    ("intermittent", 0.20),
    ("zero", 0.10),
]
ABC_DISTRIBUTION = [("A", 0.20), ("B", 0.30), ("C", 0.50)]
SERVICE_LEVELS = {"A": 0.98, "B": 0.95, "C": 0.90}
Z_VALUES = {"A": 2.05, "B": 1.65, "C": 1.28}
MATERIAL_TYPES = ["ROH", "HALB", "FERT"]
UOMS = ["EA", "KG", "L"]
MOVEMENT_SIGNS = np.array(["S", "H"])
FIELD_DESCRIPTIONS = {
    "T001": {
        "BUKRS": "Buchungskreis",
        "BUTXT": "Name des Buchungskreises",
        "LAND1": "Land",
        "WAERS": "Hauswaehrung",
    },
    "T001W": {
        "WERKS": "Werk",
        "NAME1": "Werkbezeichnung",
        "BWKEY": "Bewertungskreis",
        "BUKRS": "Zugeordneter Buchungskreis",
    },
    "T006": {
        "MSEHI": "Mengeneinheit",
        "MSEHL": "Bezeichnung der Einheit",
        "DIMID": "Dimensionsschluessel",
    },
    "T134": {
        "MTART": "Materialart",
        "MTBEZ": "Bezeichnung der Materialart",
    },
    "MARA": {
        "MATNR": "Materialnummer",
        "MTART": "Materialart",
        "MATKL": "Warengruppe",
        "MEINS": "Basismengeneinheit",
        "XCHPF": "Chargenpflicht",
    },
    "MARC": {
        "MATNR": "Materialnummer",
        "WERKS": "Werk",
        "DISPO": "Disponent",
        "PLIFZ": "Wiederbeschaffungszeit in Tagen",
        "EISBE": "Sicherheitsbestand",
        "BESKZ": "Beschaffungsart",
        "DISMM": "Dispositionsmerkmal",
        "MINBE": "Meldebestand",
        "ABC_CLASS": "ABC-Kennzeichen",
        "SERVICE_LEVEL": "Servicegrad",
        "DEMAND_PATTERN": "Nachfragemuster",
        "REORDER_POINT": "Bestellpunkt",
    },
    "MBEW": {
        "MATNR": "Materialnummer",
        "BWKEY": "Bewertungskreis",
        "BKLAS": "Bewertungsklasse",
        "VPRSV": "Preissteuerung",
        "STPRS": "Standardpreis",
        "PEINH": "Preiseinheit",
        "WAERS": "Waehrung",
        "LBKUM": "Bewerteter Bestand",
    },
    "MATDOC": {
        "MATNR": "Materialnummer",
        "WERKS": "Werk",
        "BUDAT": "Buchungsdatum",
        "MENGE": "Bewegungsmenge",
        "SHKZG": "Soll/Haben-Kennzeichen",
    },
}


@dataclass(frozen=True)
class MaterialRecord:
    matnr: str
    werks: str
    plant_name: str
    mtart: str
    meins: str
    demand_pattern: str
    abc_class: str
    service_level: float
    lead_time_days: int
    standard_price: float
    currency: str
    safety_stock: float | None
    reorder_point: int
    current_stock: int
    base_monthly_demand: float


def build_rng() -> tuple[np.random.Generator, int]:
    seed = time.time_ns()
    return np.random.default_rng(seed), seed


def weighted_choice(rng: np.random.Generator, options: list[tuple[str, float]]) -> str:
    labels = [label for label, _ in options]
    weights = [weight for _, weight in options]
    return str(rng.choice(labels, p=weights))


def generate_plant_codes(num_plants: int) -> list[str]:
    if num_plants <= len(DEFAULT_PLANTS):
        return DEFAULT_PLANTS[:num_plants]

    plants = DEFAULT_PLANTS.copy()
    next_code = 2200
    while len(plants) < num_plants:
        plants.append(str(next_code))
        next_code += 100
    return plants


def generate_price(rng: np.random.Generator) -> float:
    tier = weighted_choice(
        rng,
        [("cheap", 0.30), ("medium", 0.50), ("expensive", 0.20)],
    )
    if tier == "cheap":
        return round(float(rng.uniform(0.5, 10.0)), 2)
    if tier == "medium":
        return round(float(rng.uniform(10.0, 200.0)), 2)
    return round(float(rng.uniform(200.0, 5000.0)), 2)


def generate_material_id(plant_idx: int, material_idx: int) -> str:
    return f"MAT{plant_idx + 1:02d}{material_idx + 1:04d}"


def plant_name(werks: str) -> str:
    return f"Plant {werks}"


def company_code(werks: str) -> str:
    return f"{werks[:2]}00"


def generate_demand_dates(
    rng: np.random.Generator,
    pattern: str,
    start_date: date,
    end_date: date,
) -> list[date]:
    days = (end_date - start_date).days
    if pattern == "zero":
        count = int(rng.integers(0, 3))
    elif pattern == "stable":
        count = int(rng.integers(max(12, days // 10), max(20, days // 5)))
    elif pattern == "volatile":
        count = int(rng.integers(max(18, days // 15), max(30, days // 4)))
    elif pattern == "seasonal":
        count = int(rng.integers(max(12, days // 20), max(24, days // 8)))
    else:
        count = int(rng.integers(max(6, days // 50), max(12, days // 18)))

    offsets = rng.integers(0, max(days, 1) + 1, size=count)
    dates = [start_date + timedelta(days=int(offset)) for offset in offsets]
    dates.sort()
    return dates


def quantity_for_pattern(
    rng: np.random.Generator,
    pattern: str,
    current_date: date,
    base_monthly_demand: float,
) -> int:
    if pattern == "zero":
        return 0

    daily_base = max(base_monthly_demand / 4.0, 1.0)

    if pattern == "stable":
        quantity = rng.normal(daily_base, daily_base * 0.15)
    elif pattern == "volatile":
        quantity = rng.normal(daily_base, daily_base * 0.70)
    elif pattern == "seasonal":
        seasonal_factor = 1.0 + 0.60 * math.sin((current_date.timetuple().tm_yday / 365.0) * 2 * math.pi)
        quantity = rng.normal(daily_base * seasonal_factor, daily_base * 0.25)
    else:
        burst_multiplier = 1.0 if rng.random() > 0.35 else rng.uniform(2.0, 4.5)
        quantity = rng.normal(daily_base * burst_multiplier, daily_base * 0.85)

    return max(1, int(round(quantity)))


def build_matdoc_rows(
    rng: np.random.Generator,
    material: MaterialRecord,
    start_date: date,
    end_date: date,
) -> list[dict[str, object]]:
    dates = generate_demand_dates(rng, material.demand_pattern, start_date, end_date)
    rows: list[dict[str, object]] = []
    for movement_date in dates:
        quantity = quantity_for_pattern(
            rng,
            material.demand_pattern,
            movement_date,
            material.base_monthly_demand,
        )
        if quantity <= 0:
            continue
        shkzg = str(rng.choice(MOVEMENT_SIGNS, p=[0.92, 0.08]))
        rows.append(
            {
                "MATNR": material.matnr,
                "WERKS": material.werks,
                "BUDAT": pd.Timestamp(movement_date),
                "MENGE": quantity,
                "SHKZG": shkzg,
            }
        )
    return rows


def calculate_safety_stock(
    demand_quantities: list[int],
    lead_time_days: int,
    abc_class: str,
) -> int:
    if not demand_quantities:
        return 0
    if len(demand_quantities) > 1:
        demand_std = float(np.std(demand_quantities, ddof=1))
    else:
        demand_std = float(demand_quantities[0]) * 0.1
    safety_stock = Z_VALUES[abc_class] * demand_std * math.sqrt(lead_time_days)
    return max(0, int(round(safety_stock)))


def build_materials(
    rng: np.random.Generator,
    plants: list[str],
    materials_per_plant: int,
    start_date: date,
    end_date: date,
) -> tuple[list[MaterialRecord], pd.DataFrame]:
    materials: list[MaterialRecord] = []
    matdoc_rows: list[dict[str, object]] = []

    for plant_idx, werks in enumerate(plants):
        for material_idx in range(materials_per_plant):
            matnr = generate_material_id(plant_idx, material_idx)
            abc_class = weighted_choice(rng, ABC_DISTRIBUTION)
            demand_pattern = weighted_choice(rng, DEMAND_PATTERNS)
            lead_time_days = int(rng.integers(5, 61))
            standard_price = generate_price(rng)
            base_monthly_demand = float(rng.uniform(20, 400))
            mtart = str(rng.choice(MATERIAL_TYPES, p=[0.45, 0.25, 0.30]))
            meins = str(rng.choice(UOMS, p=[0.60, 0.25, 0.15]))

            base_record = MaterialRecord(
                matnr=matnr,
                werks=werks,
                plant_name=plant_name(werks),
                mtart=mtart,
                meins=meins,
                demand_pattern=demand_pattern,
                abc_class=abc_class,
                service_level=SERVICE_LEVELS[abc_class],
                lead_time_days=lead_time_days,
                standard_price=standard_price,
                currency="EUR",
                safety_stock=0.0,
                reorder_point=0,
                current_stock=0,
                base_monthly_demand=base_monthly_demand,
            )

            material_rows = build_matdoc_rows(rng, base_record, start_date, end_date)
            matdoc_rows.extend(material_rows)

            issued_quantities = [int(row["MENGE"]) for row in material_rows if row["SHKZG"] == "S"]
            safety_stock = calculate_safety_stock(issued_quantities, lead_time_days, abc_class)
            avg_daily_demand = (
                sum(issued_quantities) / max((end_date - start_date).days, 1)
                if issued_quantities
                else 0
            )
            reorder_point = int(round(avg_daily_demand * lead_time_days + safety_stock))
            current_stock = max(
                0,
                int(round(reorder_point * rng.uniform(0.8, 1.6) + rng.uniform(0, max(safety_stock, 1)))),
            )

            materials.append(
                MaterialRecord(
                    matnr=matnr,
                    werks=werks,
                    plant_name=plant_name(werks),
                    mtart=mtart,
                    meins=meins,
                    demand_pattern=demand_pattern,
                    abc_class=abc_class,
                    service_level=SERVICE_LEVELS[abc_class],
                    lead_time_days=lead_time_days,
                    standard_price=standard_price,
                    currency="EUR",
                    safety_stock=float(safety_stock),
                    reorder_point=reorder_point,
                    current_stock=current_stock,
                    base_monthly_demand=base_monthly_demand,
                )
            )

    apply_special_safety_stock_rules(rng, materials)
    matdoc_df = pd.DataFrame(matdoc_rows)
    if matdoc_df.empty:
        matdoc_df = pd.DataFrame(columns=["MATNR", "WERKS", "BUDAT", "MENGE", "SHKZG"])
    else:
        matdoc_df = matdoc_df.sort_values(["WERKS", "MATNR", "BUDAT"]).reset_index(drop=True)
    return materials, matdoc_df


def apply_special_safety_stock_rules(
    rng: np.random.Generator,
    materials: list[MaterialRecord],
) -> None:
    total = len(materials)
    if total == 0:
        return

    null_count = max(1, int(round(total * 0.10)))
    high_count = max(1, int(round(total * 0.15)))
    indices = np.arange(total)
    rng.shuffle(indices)
    null_indices = set(indices[:null_count].tolist())
    high_indices = set(indices[null_count:null_count + high_count].tolist())

    for idx, material in enumerate(materials):
        if idx in null_indices:
            materials[idx] = MaterialRecord(**{**material.__dict__, "safety_stock": None})
        elif idx in high_indices and material.safety_stock is not None:
            multiplier = float(rng.uniform(2.0, 4.0))
            materials[idx] = MaterialRecord(
                **{**material.__dict__, "safety_stock": float(int(round(material.safety_stock * multiplier)))}
            )


def build_tables(
    materials: list[MaterialRecord],
    plants: list[str],
    matdoc: pd.DataFrame,
    rng: np.random.Generator,
) -> dict[str, pd.DataFrame]:
    abc_visibility_mask = np.ones(len(materials), dtype=bool)
    hidden_count = len(materials) // 2
    if hidden_count:
        hidden_indices = rng.choice(len(materials), size=hidden_count, replace=False)
        abc_visibility_mask[hidden_indices] = False
    unique_companies = sorted({company_code(plant) for plant in plants})
    t001 = pd.DataFrame(
        [
            {
                "BUKRS": bukrs,
                "BUTXT": f"Company {bukrs}",
                "LAND1": "DE",
                "WAERS": "EUR",
            }
            for bukrs in unique_companies
        ]
    )

    t001w = pd.DataFrame(
        [
            {
                "WERKS": werks,
                "NAME1": plant_name(werks),
                "BWKEY": werks,
                "BUKRS": company_code(werks),
            }
            for werks in plants
        ]
    )

    t006 = pd.DataFrame(
        [
            {"MSEHI": "EA", "MSEHL": "Each", "DIMID": "QUAN"},
            {"MSEHI": "KG", "MSEHL": "Kilogram", "DIMID": "MASS"},
            {"MSEHI": "L", "MSEHL": "Liter", "DIMID": "VOLUME"},
        ]
    )

    t134 = pd.DataFrame(
        [
            {"MTART": "ROH", "MTBEZ": "Raw Material"},
            {"MTART": "HALB", "MTBEZ": "Semi-Finished"},
            {"MTART": "FERT", "MTBEZ": "Finished Product"},
        ]
    )

    mara = pd.DataFrame(
        [
            {
                "MATNR": material.matnr,
                "MTART": material.mtart,
                "MATKL": f"MG{(idx % 12) + 1:02d}",
                "MEINS": material.meins,
                "XCHPF": "X" if idx % 5 == 0 else "",
            }
            for idx, material in enumerate(materials)
        ]
    ).drop_duplicates(subset=["MATNR"]).reset_index(drop=True)

    marc = pd.DataFrame(
        [
            {
                "MATNR": material.matnr,
                "WERKS": material.werks,
                "DISPO": f"{(idx % 9) + 1:03d}",
                "PLIFZ": material.lead_time_days,
                "EISBE": material.safety_stock,
                "BESKZ": "F",
                "DISMM": "PD",
                "MINBE": max(0, int(round((material.safety_stock or 0) * 0.5))),
                "ABC_CLASS": material.abc_class if abc_visibility_mask[idx] else "",
                "SERVICE_LEVEL": material.service_level,
                "DEMAND_PATTERN": material.demand_pattern,
                "REORDER_POINT": material.reorder_point,
            }
            for idx, material in enumerate(materials)
        ]
    )

    mbew = pd.DataFrame(
        [
            {
                "MATNR": material.matnr,
                "BWKEY": material.werks,
                "BKLAS": {"A": "3000", "B": "3100", "C": "3200"}[material.abc_class],
                "VPRSV": "S",
                "STPRS": material.standard_price,
                "PEINH": 1,
                "WAERS": material.currency,
                "LBKUM": material.current_stock,
            }
            for material in materials
        ]
    )

    return {
        "T001": t001,
        "T001W": t001w,
        "T006": t006,
        "T134": t134,
        "MARA": mara,
        "MARC": marc,
        "MBEW": mbew,
        "MATDOC": matdoc,
    }


def generate_sap_dataset(
    num_plants: int = 4,
    materials_per_plant: int = 100,
    years_of_history: int = 3,
) -> dict[str, object]:
    rng, seed = build_rng()
    end_date = datetime.utcnow().date()
    start_date = end_date - timedelta(days=365 * years_of_history)
    plants = generate_plant_codes(num_plants)

    materials, matdoc = build_materials(
        rng=rng,
        plants=plants,
        materials_per_plant=materials_per_plant,
        start_date=start_date,
        end_date=end_date,
    )
    tables = build_tables(materials, plants, matdoc, rng)

    stats = {
        "total_materials": len(tables["MARC"]),
        "total_plants": len(plants),
        "matdoc_rows": len(tables["MATDOC"]),
    }

    metadata = {
        "seed": seed,
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
    }

    return {"tables": tables, "stats": stats, "metadata": metadata}


def create_excel_file(tables: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name in ["T001", "T001W", "T006", "T134", "MARA", "MARC", "MBEW", "MATDOC"]:
            df = tables[sheet_name]
            df.to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=2)

            ws = writer.book[sheet_name]
            descriptions = FIELD_DESCRIPTIONS.get(sheet_name, {})
            for col_idx, column_name in enumerate(df.columns, start=1):
                ws.cell(row=1, column=col_idx, value=descriptions.get(column_name, column_name))
                header_cell = ws.cell(row=2, column=col_idx, value=column_name)
                header_cell.font = Font(bold=True)

            last_col = get_column_letter(len(df.columns))
            last_row = len(df) + 2
            ws.auto_filter.ref = f"A2:{last_col}{max(last_row, 2)}"

            for col_idx in range(1, len(df.columns) + 1):
                max_len = 0
                for row_idx in range(1, min(ws.max_row, 200) + 1):
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    if cell_value is None:
                        continue
                    max_len = max(max_len, len(str(cell_value)))
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)
    output.seek(0)
    return output.getvalue()
