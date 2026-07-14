#!/usr/bin/env python3
"""Non-interactive sample run of the module calculator."""

from module_calculator import (
    calculate,
    load_bom_from_excel,
    load_inventory_from_excel,
    print_table,
)

PROJECT_NAME = "Test Run"

# Sample build quantities for each module type (0 = skip)
SAMPLE_MODULE_COUNTS = {
    "Tip Holder 1000": 1,
    "Tip Holder 100": 1,
    "MO:HEAT": 4,
    "Tube Holder 8T2": 1,
    "Tube Holder 888": 1,
    "MO:TILT PASSIVE": 1,
    "MO:TILT": 0,
    "MO:COOL 2ml": 1,
    "MO:COOL 15m": 0,
    "MO:CROSCOPE": 0,
    "MO:SHAKE": 0,
}


def main() -> None:
    module_type_names, part_category, module_types_bom = load_bom_from_excel()
    inventory_stock = load_inventory_from_excel()

    module_counts = {
        module_type: SAMPLE_MODULE_COUNTS.get(module_type, 0)
        for module_type in module_type_names
    }

    total_modules, total_parts, _ = calculate(module_counts, module_types_bom)
    print_table(
        PROJECT_NAME,
        module_type_names,
        part_category,
        module_counts,
        total_modules,
        total_parts,
        inventory_stock,
    )


if __name__ == "__main__":
    main()
