# msoShapeType
MSO_SHAPE_TYPE_MAP = {
    30: "3DModel",  # 3D モデル
    1: "AutoShape",  # オートシェイプ
    2: "Callout",  # 吹き出し
    20: "Canvas",  # キャンバス
    3: "Chart",  # グラフ
    4: "Comment",  # コメント
    27: "ContentApp",  # コンテンツ Office アドイン
    21: "Diagram",  # ダイアグラム
    7: "EmbeddedOLEObject",  # 埋め込み OLE オブジェクト
    8: "FormControl",  # フォーム コントロール
    5: "Freeform",  # フリーフォーム
    28: "Graphic",  # グラフィック
    6: "Group",  # グループ
    24: "IgxGraphic",  # SmartArt グラフィック
    22: "Ink",  # インク
    23: "InkComment",  # インク コメント
    9: "Line",  # Line
    31: "Linked3DModel",  # リンクされた 3D モデル
    29: "LinkedGraphic",  # リンクされたグラフィック
    10: "LinkedOLEObject",  # リンク OLE オブジェクト
    11: "LinkedPicture",  # リンク画像
    16: "Media",  # メディア
    12: "OLEControlObject",  # OLE コントロール オブジェクト
    13: "Picture",  # 画像
    14: "Placeholder",  # プレースホルダー
    18: "ScriptAnchor",  # スクリプト アンカー
    -2: "ShapeTypeMixed",  # 図形の種類の組み合わせ
    25: "Slicer",  # Slicer
    19: "Table",  # テーブル
    17: "TextBox",  # テキスト ボックス
    15: "TextEffect",  # テキスト効果
    26: "WebVideo",  # Web ビデオ
}
# AutoShapeType（Type==1のときのみ）
MSO_AUTO_SHAPE_TYPE_MAP = {
    149: "10pointStar",  # 10-point star
    150: "12pointStar",  # 12-point star
    94: "16pointStar",  # 16-point star
    95: "24pointStar",  # 24-point star
    96: "32pointStar",  # 32-point star
    91: "4pointStar",  # 4-point star
    92: "5pointStar",  # 5-point star
    147: "6pointStar",  # 6-point star
    148: "7pointStar",  # 7-point star
    93: "8pointStar",  # 8-point star
    129: "ActionButtonBackorPrevious",  # Back or Previous button
    131: "ActionButtonBeginning",  # Beginning button
    125: "ActionButtonCustom",  # Custom action button
    134: "ActionButtonDocument",  # Document button
    132: "ActionButtonEnd",  # End button
    130: "ActionButtonForwardorNext",  # Forward or Next button
    127: "ActionButtonHelp",  # Help button
    126: "ActionButtonHome",  # Home button
    128: "ActionButtonInformation",  # Information button
    136: "ActionButtonMovie",  # Movie button
    133: "ActionButtonReturn",  # Return button
    135: "ActionButtonSound",  # Sound button
    25: "Arc",  # Arc
    137: "Balloon",  # Balloon
    41: "BentArrow",  # Bent Arrow
    44: "BentUpArrow",  # Bent Up Arrow
    15: "Bevel",  # Bevel
    20: "BlockArc",  # Block Arc
    13: "Can",  # Can
    182: "ChartPlus",  # Chart Plus
    181: "ChartStar",  # Chart Star
    180: "ChartX",  # Chart X
    52: "Chevron",  # Chevron
    161: "Chord",  # Chord
    60: "CircularArrow",  # Circular Arrow
    179: "Cloud",  # Cloud
    108: "CloudCallout",  # Cloud Callout
    162: "Corner",  # Corner
    169: "CornerTabs",  # Corner Tabs
    11: "Cross",  # Cross
    14: "Cube",  # Cube
    48: "CurvedDownArrow",  # Curved Down Arrow
    100: "CurvedDownRibbon",  # Curved Down Ribbon
    46: "CurvedLeftArrow",  # Curved Left Arrow
    45: "CurvedRightArrow",  # Curved Right Arrow
    47: "CurvedUpArrow",  # Curved Up Arrow
    99: "CurvedUpRibbon",  # Curved Up Ribbon
    144: "Decagon",  # Decagon
    141: "DiagonalStripe",  # Diagonal Stripe
    4: "Diamond",  # Diamond
    146: "Dodecagon",  # Dodecagon
    18: "Donut",  # Donut
    27: "DoubleBrace",  # Double Brace
    26: "DoubleBracket",  # Double Bracket
    104: "DoubleWave",  # Double Wave
    36: "DownArrow",  # Down Arrow
    56: "DownArrowCallout",  # Down Arrow Callout
    98: "DownRibbon",  # Down Ribbon
    89: "Explosion1",  # Explosion 1
    90: "Explosion2",  # Explosion 2
    62: "FlowchartAlternateProcess",  # Alternate Process
    75: "FlowchartCard",  # Card
    79: "FlowchartCollate",  # Collate
    73: "FlowchartConnector",  # Connector
    64: "FlowchartData",  # Data
    63: "FlowchartDecision",  # Decision
    84: "FlowchartDelay",  # Delay
    87: "FlowchartDirectAccessStorage",  # Direct Access Storage
    88: "FlowchartDisplay",  # Display
    67: "FlowchartDocument",  # Document
    81: "FlowchartExtract",  # Extract
    66: "FlowchartInternalStorage",  # Internal Storage
    86: "FlowchartMagneticDisk",  # Magnetic Disk
    71: "FlowchartManualInput",  # Manual Input
    72: "FlowchartManualOperation",  # Manual Operation
    82: "FlowchartMerge",  # Merge
    68: "FlowchartMultidocument",  # Multidocument
    139: "FlowchartOfflineStorage",  # Offline Storage
    74: "FlowchartOffpageConnector",  # Offpage Connector
    78: "FlowchartOr",  # Or
    65: "FlowchartPredefinedProcess",  # Predefined Process
    70: "FlowchartPreparation",  # Preparation
    61: "FlowchartProcess",  # Process
    76: "FlowchartPunchedTape",  # Punched Tape
    85: "FlowchartSequentialAccessStorage",  # Sequential Access Storage
    80: "FlowchartSort",  # Sort
    83: "FlowchartStoredData",  # Stored Data
    77: "FlowchartSummingJunction",  # Summing Junction
    69: "FlowchartTerminator",  # Terminator
    16: "FoldedCorner",  # Folded Corner
    158: "Frame",  # Frame
    174: "Funnel",  # Funnel
    172: "Gear6",  # Gear with six teeth
    173: "Gear9",  # Gear with nine teeth
    159: "HalfFrame",  # Half Frame
    21: "Heart",  # Heart
    145: "Heptagon",  # Heptagon
    10: "Hexagon",  # Hexagon
    102: "HorizontalScroll",  # Horizontal Scroll
    7: "IsoscelesTriangle",  # Isosceles Triangle
    34: "LeftArrow",  # Left Arrow
    54: "LeftArrowCallout",  # Left Arrow Callout
    31: "LeftBrace",  # Left Brace
    29: "LeftBracket",  # Left Bracket
    176: "LeftCircularArrow",  # Left Circular Arrow
    37: "LeftRightArrow",  # LeftRight Arrow
    57: "LeftRightArrowCallout",  # LeftRight Arrow Callout
    177: "LeftRightCircularArrow",  # LeftRight Circular Arrow
    140: "LeftRightRibbon",  # LeftRight Ribbon
    40: "LeftRightUpArrow",  # LeftRightUp Arrow
    43: "LeftUpArrow",  # LeftUp Arrow
    22: "LightningBolt",  # Lightning Bolt
    109: "LineCallout1",  # Line Callout 1
    113: "LineCallout1AccentBar",  # Line Callout 1 Accent Bar
    121: "LineCallout1BorderandAccentBar",  # Line Callout 1 Border and Accent Bar
    117: "LineCallout1NoBorder",  # Line Callout 1 No Border
    110: "LineCallout2",  # Line Callout 2
    114: "LineCallout2AccentBar",  # Line Callout 2 Accent Bar
    122: "LineCallout2BorderandAccentBar",  # Line Callout 2 Border and Accent Bar
    118: "LineCallout2NoBorder",  # Line Callout 2 No Border
    111: "LineCallout3",  # Line Callout 3
    115: "LineCallout3AccentBar",  # Line Callout 3 Accent Bar
    123: "LineCallout3BorderandAccentBar",  # Line Callout 3 Border and Accent Bar
    119: "LineCallout3NoBorder",  # Line Callout 3 No Border
    112: "LineCallout4",  # Line Callout 4
    116: "LineCallout4AccentBar",  # Line Callout 4 Accent Bar
    124: "LineCallout4BorderandAccentBar",  # Line Callout 4 Border and Accent Bar
    120: "LineCallout4NoBorder",  # Line Callout 4 No Border
    183: "LineInverse",  # Line inverse
    166: "MathDivide",  # Math Divide
    167: "MathEqual",  # Math Equal
    164: "MathMinus",  # Math Minus
    165: "MathMultiply",  # Math Multiply
    168: "MathNotEqual",  # Math Not Equal
    163: "MathPlus",  # Math Plus
    -2: "Mixed",  # Mixed (combination)
    24: "Moon",  # Moon
    143: "NonIsoscelesTrapezoid",  # Non Isosceles Trapezoid
    19: "NoSymbol",  # No symbol
    50: "NotchedRightArrow",  # Notched Right Arrow
    138: "NotPrimitive",  # Not supported
    6: "Octagon",  # Octagon
    9: "Oval",  # Oval
    107: "OvalCallout",  # Oval Callout
    2: "Parallelogram",  # Parallelogram
    51: "Pentagon",  # Pentagon
    142: "Pie",  # Pie
    175: "PieWedge",  # Pie Wedge
    28: "Plaque",  # Plaque
    171: "PlaqueTabs",  # Plaque Tabs
    39: "QuadArrow",  # Quad Arrow
    59: "QuadArrowCallout",  # Quad Arrow Callout
    1: "Rectangle",  # Rectangle
    105: "RectangularCallout",  # Rectangular Callout
    12: "RegularPentagon",  # Regular Pentagon
    33: "RightArrow",  # Right Arrow
    53: "RightArrowCallout",  # Right Arrow Callout
    32: "RightBrace",  # Right Brace
    30: "RightBracket",  # Right Bracket
    8: "RightTriangle",  # Right Triangle
    151: "Round1Rectangle",  # Round 1 Rectangle
    157: "Round2DiagRectangle",  # Round 2 Diagonal Rectangle
    152: "Round2SameRectangle",  # Round 2 Same Rectangle
    5: "RoundedRectangle",  # Rounded Rectangle
    106: "RoundedRectangularCallout",  # Rounded Rectangular Callout
    17: "SmileyFace",  # Smiley Face
    155: "Snip1Rectangle",  # Snip 1 Rectangle
    157: "Snip2DiagRectangle",  # Snip 2 Diagonal Rectangle
    156: "Snip2SameRectangle",  # Snip 2 Same Rectangle
    154: "SnipRoundRectangle",  # Snip Round Rectangle
    170: "SquareTabs",  # Square Tabs
    49: "StripedRightArrow",  # Striped Right Arrow
    23: "Sun",  # Sun
    178: "SwooshArrow",  # Swoosh Arrow
    160: "Tear",  # Tear
    3: "Trapezoid",  # Trapezoid
    35: "UpArrow",  # Up Arrow
    55: "UpArrowCallout",  # Up Arrow Callout
    38: "UpDownArrow",  # Up Down Arrow
    58: "UpDownArrowCallout",  # Up Down Arrow Callout
    97: "UpRibbon",  # Up Ribbon
    42: "UTurnArrow",  # U-Turn Arrow
    101: "VerticalScroll",  # Vertical Scroll
    103: "Wave",  # Wave
}

XL_CHART_TYPE_MAP = {
    -4098: "3DArea",
    78: "3DAreaStacked",
    79: "3DAreaStacked100",
    60: "3DBarClustered",
    61: "3DBarStacked",
    62: "3DBarStacked100",
    -4100: "3DColumn",
    54: "3DColumnClustered",
    55: "3DColumnStacked",
    56: "3DColumnStacked100",
    -4101: "3DLine",
    -4102: "3DPie",
    70: "3DPieExploded",
    1: "Area",
    135: "AreaEx",
    76: "AreaStacked",
    77: "AreaStacked100",
    137: "AreaStacked100Ex",
    136: "AreaStackedEx",
    57: "BarClustered",
    132: "BarClusteredEx",
    71: "BarOfPie",
    58: "BarStacked",
    59: "BarStacked100",
    134: "BarStacked100Ex",
    133: "BarStackedEx",
    121: "Boxwhisker",
    15: "Bubble",
    87: "Bubble3DEffect",
    139: "BubbleEx",
    51: "ColumnClustered",
    124: "ColumnClusteredEx",
    52: "ColumnStacked",
    53: "ColumnStacked100",
    126: "ColumnStacked100Ex",
    125: "ColumnStackedEx",
    -4152: "Combo",
    115: "ComboAreaStackedColumnClustered",
    113: "ComboColumnClusteredLine",
    114: "ComboColumnClusteredLineSecondaryAxis",
    102: "ConeBarClustered",
    103: "ConeBarStacked",
    104: "ConeBarStacked100",
    105: "ConeCol",
    99: "ConeColClustered",
    100: "ConeColStacked",
    101: "ConeColStacked100",
    95: "CylinderBarClustered",
    96: "CylinderBarStacked",
    97: "CylinderBarStacked100",
    98: "CylinderCol",
    92: "CylinderColClustered",
    93: "CylinderColStacked",
    94: "CylinderColStacked100",
    -4120: "Doughnut",
    131: "DoughnutEx",
    80: "DoughnutExploded",
    123: "Funnel",
    118: "Histogram",
    4: "Line",
    127: "LineEx",
    65: "LineMarkers",
    66: "LineMarkersStacked",
    67: "LineMarkersStacked100",
    63: "LineStacked",
    64: "LineStacked100",
    129: "LineStacked100Ex",
    128: "LineStackedEx",
    116: "OtherCombinations",
    122: "Pareto",
    5: "Pie",
    130: "PieEx",
    69: "PieExploded",
    68: "PieOfPie",
    109: "PyramidBarClustered",
    110: "PyramidBarStacked",
    111: "PyramidBarStacked100",
    112: "PyramidCol",
    106: "PyramidColClustered",
    107: "PyramidColStacked",
    108: "PyramidColStacked100",
    -4151: "Radar",
    82: "RadarFilled",
    81: "RadarMarkers",
    140: "RegionMap",
    88: "StockHLC",
    89: "StockOHLC",
    90: "StockVHLC",
    91: "StockVOHLC",
    -2: "SuggestedChart",
    120: "Sunburst",
    83: "Surface",
    85: "SurfaceTopView",
    86: "SurfaceTopViewWireframe",
    84: "SurfaceWireframe",
    117: "Treemap",
    119: "Waterfall",
    -4169: "XYScatter",
    138: "XYScatterEx",
    74: "XYScatterLines",
    75: "XYScatterLinesNoMarkers",
    72: "XYScatterSmooth",
    73: "XYScatterSmoothNoMarkers",
}

XL_CHART_TYPE_PIE = {5, 18, 69, 68, 130}  # Pie, PieExploded, PieOfPie, PieEx
XL_CHART_TYPE_DOUGHNUT = {-4120, 131, 80}  # Doughnut, DoughnutEx, DoughnutExploded
XL_CHART_TYPE_XY_SCATTER = {-4169, 72, 74, 75, 73, 138}  # XYScatter variants
XL_CHART_TYPE_BOXWHISKER = {121}
XL_CHART_TYPE_BUBBLE = {15, 87, 139}
XL_CHART_TYPE_LINE = {4, 127, 65, 66, 67, 63, 64, 128, 129}
