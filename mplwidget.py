from PyQt5.QtWidgets import*
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
from matplotlib.figure import Figure

from matplotlib.backends.backend_qt5agg import FigureCanvas

from matplotlib.figure import Figure

    
class MplWidget(QWidget):
    
    def __init__(self, parent = None):

        QWidget.__init__(self, parent)
        
        self.canvas = FigureCanvas(Figure(figsize= (250, 200)))
        vertical_layout = QVBoxLayout()
        vertical_layout.addWidget(self.canvas)
        self.canvas.figure.patch.set_facecolor('none')
        self.canvas.axes = self.canvas.figure.add_subplot(111, projection='3d')
        self.canvas.figure.tight_layout()
        self.setLayout(vertical_layout)