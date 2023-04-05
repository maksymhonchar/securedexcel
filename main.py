import sys

from PyQt5.QtWidgets import QApplication

from controller import SpreadsheetController
from model import SpreadsheetModel
from view import SpreadsheetView


def main():
    app = QApplication(sys.argv)

    model = SpreadsheetModel()
    view = SpreadsheetView()
    controller = SpreadsheetController(
        model=model,
        view=view
    )
    view.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
