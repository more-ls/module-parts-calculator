#!/usr/bin/env python3
"""Non-interactive sample run of the module calculator."""

from module_calculator import (
    INVENTORY_XLSX,
    PART_DEFINITIONS_XLSX,
    calculate,
    load_bom_from_excel,
    load_inventory_from_excel,
    print_table,
)

PROJECT_NAME = "Test Run"

# Sample build quantities for each module type (0 = skip)
SAMPLE_MODULE_COUNTS = {
    "tip holder 1000": 1,
    "tip holder 100": 1,
    "MO:HEAT": 4,
    "tube holder 8T2": 1,
    "tube holder 888": 1,
    "MO:TILT PASSIVE": 1,
    "MO:TILT": 0,
    "MO:COOL 2ml": 1,
    "MO:COOL 15m": 0,
    "MO:CROSCOPE": 0,
    "MO:SHAKE": 0,
}


def main() -> None:
    print(f"Loading {PART_DEFINITIONS_XLSX.name}...")
    module_type_names, part_category, module_types_bom = load_bom_from_excel()

    inventory_stock = load_inventory_from_excel()
    print(f"Loading {INVENTORY_XLSX.name}...")

    module_counts = {
        module_type: SAMPLE_MODULE_COUNTS.get(module_type, 0)
        for module_type in module_type_names
    }

    print(f"\nProject: {PROJECT_NAME}")
    print("Module counts:")
    for module_type, count in module_counts.items():
        if count > 0:
            print(f"  {module_type}: {count}")

    total_modules, total_parts, _ = calculate(module_counts, module_types_bom)
    pdf_path = print_table(
        PROJECT_NAME,
        module_type_names,
        part_category,
        module_counts,
        total_modules,
        total_parts,
        inventory_stock,
    )

    print(f"\nTest run complete.")
    print(f"  Modules: {total_modules}")
    print(f"  Unique parts: {len(total_parts)}")
    print(f"  Total parts needed: {sum(total_parts.values())}")
    print(f"  PDF: {pdf_path}")


if __name__ == "__main__":
    main()
