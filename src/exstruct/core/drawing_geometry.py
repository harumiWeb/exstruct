"""Pure connector angle helpers shared by OOXML and COM shapes."""

import math
from typing import Literal, cast


def compute_line_angle_deg(w: float, h: float) -> float:
    """
    Compute the clockwise angle (in degrees) in Excel coordinates where 0° points East.

    Parameters:
        w (float): Horizontal delta (width, positive to the right).
        h (float): Vertical delta (height, positive downward).

    Returns:
        float: Angle in degrees measured clockwise from East (e.g., 0° = East, 90° = South).
    """
    return math.degrees(math.atan2(h, w)) % 360.0


def angle_to_compass(
    angle: float,
) -> Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"]:
    """
    Map an angle in degrees to one of eight compass directions.

    The angle is interpreted with 0 degrees at East and increasing values rotating counterclockwise (45 -> NE, 90 -> N).

    Parameters:
        angle (float): Angle in degrees.

    Returns:
        str: One of `"E"`, `"SE"`, `"S"`, `"SW"`, `"W"`, `"NW"`, `"N"`, or `"NE"` corresponding to the nearest 8-point compass direction.
    """
    dirs = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"]
    idx = int(((angle + 22.5) % 360) // 45)
    return cast(Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"], dirs[idx])
