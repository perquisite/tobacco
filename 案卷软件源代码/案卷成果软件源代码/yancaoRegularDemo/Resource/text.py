from PyQt5 import QtWidgets
import sys
from PyQt5.QtWidgets import QFileDialog, QListView, QAbstractItemView, QTreeView


class Window(QtWidgets.QWidget):

    def __init__(self):

        super(Window, self).__init__()

        self.button = QtWidgets.QPushButton('Test', self)

        self.button.clicked.connect(self.handleButton)

        layout = QtWidgets.QVBoxLayout(self)

        layout.addWidget(self.button)

    def handleButton(self):
        fileDlg = QFileDialog()
        fileDlg.setFileMode(QFileDialog.DirectoryOnly)
        fileDlg.setOption(QFileDialog.DontUseNativeDialog, True)
        fileDlg.setDirectory("d:/")
        listView = fileDlg.findChild(QListView, "listView")
        if listView:
            listView.setSelectionMode(QAbstractItemView.ExtendedSelection)
        treeView = fileDlg.findChild(QTreeView, "treeView")
        if treeView:
            treeView.setSelectionMode(QAbstractItemView.ExtendedSelection)
        if fileDlg.exec_():
            folders = fileDlg.selectedFiles()
            print(folders)
            # if folders.size()>0:
            #     nativePath = QDir.toNativeSeparators(folders[0])
            #     strDir = nativePath.left(nativePath.lastIndexOf(QDir.separator()))
            #     print(nativePath)
            #     print(strDir)
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
