import shutil
import sys
import os
import win32api

from PyQt5.QtCore import Qt, QMimeData, QUrl
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFileSystemModel,
    QLineEdit,
    QLabel,
    QMenu,
    QVBoxLayout,
    QComboBox,
    QTreeView,
    QFileDialog,
    QMessageBox,
    QStatusBar,
)


class Explorer(QDialog):
    def __init__(self) -> None:
        """Initialize"""
        super().__init__()
        self.setWindowTitle("Simple File Explorer")
        self.resize(1000, 600)
        self.initUI()

    def initUI(self):
        """Create UI"""
        self._createNavigator()
        self._createModel()
        self._createTree()
        self._createLocationLabels()
        self._createStatusBar()
        self._createLayout()

    def _createNavigator(self):
        """
        # Create the top-most navigation
        # Afterwards identify user disk drives
        """
        self.navigator = QComboBox()
        self.navigator.addItems(
            win32api.GetLogicalDriveStrings().split("\000")[:-1]
        )
        self.navigator.activated.connect(self.switchDir)
        self.rootItems = win32api.GetLogicalDriveStrings().split("\000")[:-1]

    def _createModel(self):
        """Initialize the model"""
        self.model = QFileSystemModel()
        self.model.setRootPath(self.rootItems[0])

    def _createTree(self):
        """Create the tree widget and link to needed signals"""
        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setColumnWidth(0, 250)
        self.tree.setRootIndex(self.model.index(self.rootItems[0]))
        self.tree.clicked.connect(self.showLocation)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.contextMenuSignals)

    def _createLocationLabels(self):
        """Create bottom-most widget for file path"""
        self.labelFilePath = QLabel("File Path:")
        self.locationFilePath = QLineEdit()

    def _createStatusBar(self):
        """Create a simple status bar at the bottom"""
        self.statusBar = QStatusBar()
        self.statusBar.showMessage("Explorer Ready", 3000)

    def _createLayout(self):
        """Add widgets to layout"""
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.navigator)
        self.layout.addWidget(self.tree)
        self.layout.addWidget(self.labelFilePath)
        self.layout.addWidget(self.locationFilePath)
        self.layout.addWidget(self.statusBar)
        self.setLayout(self.layout)

    def _createContextMenu(self):
        """Create Context Menu"""
        self.contextMenu = QMenu()
        self.openAction = self.contextMenu.addAction("&Open File")
        self.contextMenu.addSeparator()
        self.copyAction = self.contextMenu.addAction("&Copy")
        self.pasteAction = self.contextMenu.addAction("&Paste")
        self.deleteAction = self.contextMenu.addAction("&Delete")
        self.contextMenu.addSeparator()
        self.exitAction = self.contextMenu.addAction("&Exit Explorer")

    def contextMenuSignals(self, event):
        """
        Get the working index and create the context menu
        Afterwards create the signal and check with
        signal is selected from the context menu
        """
        index = self.tree.indexAt(event)
        self._createContextMenu()
        signal = self.contextMenu.exec(self.tree.mapToGlobal(event))
        if signal is not None:
            if signal is self.openAction:
                self.model.setReadOnly(False)
                os.startfile(os.path.realpath(self.model.filePath(index)))
                self.model.setReadOnly(True)
                self.statusBar.showMessage("File or folder opened", 3000)
            if signal is self.copyAction:
                data = QMimeData()
                url = QUrl.fromLocalFile(
                    os.path.realpath(self.model.filePath(index))
                )
                data.setUrls([url])
                app.clipboard().setMimeData(data)
                self.statusBar.showMessage("File or folder copied", 3000)
            if signal is self.pasteAction:
                self.model.setReadOnly(False)
                source = app.clipboard().text().split("file")[1][4:]
                currentDir = QFileDialog()
                try:
                    shutil.copy2(source, currentDir.getExistingDirectory(self))
                except (FileNotFoundError, shutil.SameFileError):
                    self.errorMsg()
                self.model.setReadOnly(True)
                self.statusBar.showMessage("File or folder placed", 3000)
            if signal is self.deleteAction:
                self.model.setReadOnly(False)
                os.remove(os.path.relpath(self.model.filePath(index)))
                self.model.setReadOnly(True)
                self.statusBar.showMessage("File deleted", 3000)
            if signal is self.exitAction:
                self.exit()

    def showLocation(self, index):
        """Capture currently selected file location"""
        self.itemIndex = self.model.index(index.row(), 0, index.parent())
        self.locationFilePath.setText(self.model.filePath(self.itemIndex))

    def switchDir(self):
        """Switch disks based on user selection"""
        index = self.navigator.currentIndex()
        self.tree.setRootIndex(self.model.index(self.rootItems[index]))

    def errorMsg(self):
        """
        # Throws a generic message to the user
        # if a file is copied onto itself or operation is cancelled
        """
        msg = QMessageBox()
        msg.setText("An unexpected error occurred")
        msg.setWindowTitle("Warning")
        msg.setInformativeText("Please try again")
        msg.exec()

    def exit(self):
        """Asks the user whether to exit the explorer"""
        buttonReply = QMessageBox.question(
            None,
            "Exit",
            "Quit Program?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if buttonReply == QMessageBox.Yes:
            app.exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    explorer = Explorer()
    explorer.show()
    sys.exit(app.exec())
