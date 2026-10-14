# Iran's Crop-Based Exports and Virtual Water (1986–2023)

Replication materials for the study *“Real export value per cubic metre of virtual water in Iran’s crop-based exports: Bennet decomposition and fixed-effects panel analysis, 1986–2023.”*

## Files

- `code_vwe_v1.4.R` — Data processing, Bennet decomposition, fixed-effects estimation, robustness checks, tables, and figures.
- `data_vwe_v1.4.xlsx` — Source data compiled from FAOSTAT, CWASI v1.4, and World Bank WDI.

## Reproduction

Requires R 4.2 or later.

Run in RStudio using **Source** and select the Excel file when prompted, or run:

`Rscript code_vwe_v1.4.R data_vwe_v1.4.xlsx`

Missing R packages are installed automatically when required.

## Data Sources

- **FAOSTAT:** [Crops and Livestock Products Trade](https://www.fao.org/faostat/en/#data/TCL)
- **CWASI v1.4:** [Zenodo](https://doi.org/10.5281/zenodo.18964006)
- **World Bank:** [World Development Indicators](https://databank.worldbank.org/source/world-development-indicators)

## License

- **Code:** MIT License.
- **Data compilation and documentation:** Creative Commons Attribution 4.0 International (CC BY 4.0).
