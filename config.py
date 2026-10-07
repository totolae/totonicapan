"""
Department-specific settings. To adapt the tool to another department
(San Marcos, Quetzaltenango...), this is the only file that should change.
"""

DEPARTAMENTO = "Totonicapán"

# Municipality shown in the dropdown -> internal id
MUNICIPIOS_OPCIONES = {
    "Totonicapán": 1,
    "San Cristóbal Totonicapán": 2,
    "San Francisco El Alto": 3,
    "San Andrés Xecul": 4,
    "Momostenango": 5,
    "Santa María Chiquimula": 6,
    "Santa Lucía La Reforma": 7,
    "San Bartolo Aguas Calientes": 8,
}

# Internal id -> text used to find that municipality's row in MAGA's Excel
EXCEL_MAPPINGS = {
    1: "totonicapán", 2: "san cristobal", 3: "san francisco", 4: "san andres",
    5: "momostenango", 6: "santa maria", 7: "santa lucia", 8: "san bartolo",
}

# A factura is flagged in 'Extra Detalles' when abarrotes exceed this share
UMBRAL_ALERTA_ABARROTES = 0.30
