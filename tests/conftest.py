import os

# 画面のない環境(CIなど)でもQtのウィジェットを生成できるようにする
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt5.QtWidgets import QApplication


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication(["test"])
