from pydantic import BaseModel, Field
from typing import List, Dict, Optional, Literal


class Shape(BaseModel):
    text: str
    l: int
    t: int
    w: Optional[int]
    h: Optional[int]
    type: Optional[str] = None
    rotation: Optional[float] = None
    angle_deg: Optional[float] = None
    begin_arrow_style: Optional[int] = None
    end_arrow_style: Optional[int] = None
    direction: Optional[Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"]] = None


class CellRow(BaseModel):
    r: int
    c: Dict[str, int | float | str]


class ChartSeries(BaseModel):
    name: str
    name_range: Optional[str] = None
    x_range: Optional[str] = None
    y_range: Optional[str] = None


class Chart(BaseModel):
    name: str
    chart_type: str
    title: Optional[str]
    y_axis_title: str
    y_axis_range: List[float] = Field(default_factory=list)
    series: List[ChartSeries]
    l: int
    t: int
    error: Optional[str] = None


class SheetData(BaseModel):
    rows: List[CellRow] = Field(default_factory=list)
    shapes: List[Shape] = Field(default_factory=list)
    charts: List[Chart] = Field(default_factory=list)
    tables: List[str] = Field(default_factory=list)


class WorkbookData(BaseModel):
    book_name: str
    sheets: Dict[str, SheetData]
