import sys
from PyQt5.QtWidgets import QApplication, QMainWindow

#Create a fucntion for the window
def window():
    #Collect all arguements to run code
    app = QApplication(sys.argv)
    #Create a visual window
    win = QMainWindow()
    #Place window closer to center and resize when program is run
    win.setGeometry(500,500,500,300)
    #Change application name
    win.setWindowTitle("Data Collector")
    #If hovering over nothing add "Data" tooltip
    win.setToolTip("Data")
    #Makes window visible
    win.show()
    #Exits program if x is pressed
    sys.exit(app.exec_())

#Run "window" function


window()
