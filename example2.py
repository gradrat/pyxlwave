from pyxlwave.timing import Timing
import schemdraw
from schemdraw import logic

t = Timing()

t.read_xls("example.xlsx", "Example2")
# Get all signals possible & include them
dia_dict = t.get_diagram()
dia_dict["config"] = {'hscale': 0.5}
diagram = logic.TimingDiagram(dia_dict)

d = schemdraw.Drawing()
d.add(diagram)
d.draw()

