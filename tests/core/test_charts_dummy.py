from dataclasses import dataclass

from exstruct.core.charts import get_charts
from exstruct.models.maps import XL_CHART_TYPE_MAP


@dataclass(frozen=True)
class _DummyAxisTitle:
    Text: str


@dataclass(frozen=True)
class _DummyAxis:
    HasTitle: bool
    AxisTitle: _DummyAxisTitle
    MinimumScale: float
    MaximumScale: float


@dataclass(frozen=True)
class _DummyChartTitle:
    Text: str


@dataclass(frozen=True)
class _DummySeries:
    Name: str
    Formula: str


@dataclass(frozen=True)
class _DummyChartCom:
    ChartType: int
    _series: list[_DummySeries]
    _axis: _DummyAxis
    HasTitle: bool
    ChartTitle: _DummyChartTitle

    def SeriesCollection(self) -> list[_DummySeries]:
        return self._series

    def Axes(self, *_args: int) -> _DummyAxis:
        return self._axis


@dataclass(frozen=True)
class _DummyChartObject:
    Chart: _DummyChartCom


class _DummySheetApi:
    def __init__(self, chart_objects: dict[str, _DummyChartObject]) -> None:
        self._chart_objects = chart_objects

    def ChartObjects(self, name: str) -> _DummyChartObject:
        return self._chart_objects[name]


@dataclass(frozen=True)
class _DummyChartShape:
    name: str
    width: float
    height: float
    left: float
    top: float


@dataclass(frozen=True)
class _DummySheet:
    charts: list[_DummyChartShape]
    api: object


def test_get_charts_builds_chart_and_series() -> None:
    series = [
        _DummySeries(
            Name="Series1",
            Formula='=SERIES("Sales",Sheet1!$A$1:$A$2,Sheet1!$B$1:$B$2,1)',
        )
    ]
    axis = _DummyAxis(
        HasTitle=True,
        AxisTitle=_DummyAxisTitle(Text="Y Axis"),
        MinimumScale=0.0,
        MaximumScale=100.0,
    )
    chart_com = _DummyChartCom(
        ChartType=4,
        _series=series,
        _axis=axis,
        HasTitle=True,
        ChartTitle=_DummyChartTitle(Text="Chart Title"),
    )
    chart_obj = _DummyChartObject(Chart=chart_com)
    chart_shape = _DummyChartShape(
        name="Chart1", width=200.0, height=100.0, left=10.0, top=20.0
    )
    sheet = _DummySheet(
        charts=[chart_shape],
        api=_DummySheetApi({"Chart1": chart_obj}),
    )

    charts = get_charts(sheet)

    assert len(charts) == 1
    chart = charts[0]
    assert chart.name == "Chart1"
    assert chart.chart_type == XL_CHART_TYPE_MAP[4]
    assert chart.title == "Chart Title"
    assert chart.y_axis_title == "Y Axis"
    assert chart.y_axis_range == [0.0, 100.0]
    assert chart.w == 200
    assert chart.h == 100
    assert chart.series
    assert chart.series[0].name == "Series1"
    assert chart.series[0].name_range is None


def test_get_charts_sets_error_on_failure() -> None:
    chart_shape = _DummyChartShape(
        name="Chart1", width=200.0, height=100.0, left=10.0, top=20.0
    )

    class _FailingSheetApi:
        def ChartObjects(self, _name: str) -> _DummyChartObject:
            raise RuntimeError("boom")

    sheet = _DummySheet(charts=[chart_shape], api=_FailingSheetApi())

    charts = get_charts(sheet)

    assert len(charts) == 1
    assert charts[0].error is not None
