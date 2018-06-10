from Qt import QtCore, QtGui, QtWidgets


class GuiTemplate(QtGui.QMainWindow):
    refreshed = QtCore.Signal()

    header_icon = None
    title = ''
    subtitle = ''

    def __init__(self, parent=None):
        super(GuiTemplate, self).__init__(parent=parent)

        main_layout = QtWidgets.QVBoxLayout()

        _header = self.setup_header()
        main_layout.addWidget(_header, 1)

        _body = self.setup_body()
        main_layout.addWidget(_body, 100)

        _footer = self.setup_footer()
        main_layout.addWidget(_footer, 1)

        main_layout.setContentsMargins(0, 0, 0, 0)

        self.setLayout(main_layout)

        self.minimum_size_filter = MaintainMinimumWidgetSizeFilter(self)
        self.installEventFilter(self.minimum_size_filter)

        _header.wikiURL = self.wikiURL
        _header.subURL = self.subURL
        self.setWindowTitle("{0} {1}".format(self.title, self.subtitle))

class MaintainMinimumWidgetSizeFilter(QtCore.QObject):
    """
    QObject that receives all events

    Args:
        QtCore.QObject : Base class of all Qt objects which allows for
                         object communication via signals and slots
    """

    def __init__(self, widget, parent=None):
        super(MaintainMinimumWidgetSizeFilter, self).__init__(parent=parent)
        self.widget = widget

    def event_filter(self, obj, event):
        """Checks for QEvents that are Layout Requests then modifies an object's size

        Args:
            obj (QObject): Generic object
            event (QEvent): Layout Request

        Returns:
            bool: Returns False if Layout Request QEvent is not detected otherwise modify the object
        """

        if isinstance(event, QtCore.QEvent) and event.type() == QtCore.QEvent.Type.LayoutRequest:
            size_hint = obj.sizeHint()
            size_hint.setHeight(size_hint.height() - 200)
            self.widget.resize(obj.geometry().width(), size_hint.height())

        return False
