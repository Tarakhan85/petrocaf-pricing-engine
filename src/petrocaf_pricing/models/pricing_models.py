from pydantic import BaseModel, Field

class PricingOptions(BaseModel):
    scenario_name: str = "base"
    default_rounding: int = 2
    allow_zero_rates: bool = False

class BOQRow(BaseModel):
    item_code: str
    description: str
    discipline: str
    unit: str
    quantity: float = Field(gt=0)
    base_material_rate: float = Field(ge=0)
    base_labor_rate: float = Field(ge=0)
    base_equipment_rate: float = Field(ge=0)
    environment_factor: float = Field(gt=0, default=1.0)
    complexity_factor: float = Field(gt=0, default=1.0)
    location_factor: float = Field(gt=0, default=1.0)
