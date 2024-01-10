from pyxlwave.timing import Timing
import schemdraw
from schemdraw import logic

t = Timing()

t.read_xls("example.xlsx", "Example1")
signal_list = [
    "signal1 (H)",
    "signal2 (H)",
    "signal7 (H)",
    "signal9 (D)",
    "signal10 (D)"
]
# Only plot certain signals (match lower case for now)
dia_dict = t.get_diagram(signal_list)
dia_dict["config"] = {'hscale': 0.5}
diagram = logic.TimingDiagram(dia_dict)

d = schemdraw.Drawing()
d.add(diagram)
d.draw()
