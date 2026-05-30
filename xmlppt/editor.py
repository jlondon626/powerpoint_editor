from .package import PackageMixin
from .slides import SlideSectionMixin
from .text import TextMixin
from .charts import ChartMixin
from .debug import DebugMixin
from .tables import TableMixin
from .slide_proxy import SlideProxy
from .constants import *


class PowerPointEditor(
    PackageMixin,
    SlideSectionMixin,
    TextMixin,
    ChartMixin,
    DebugMixin,
    TableMixin,
):
    """A lightweight editor for .pptx files.

    The implementation is split into focused mixins, while this facade keeps
    the public API stable for callers importing ``PowerPointEditor``.
    """

    def __init__(self, input_pptx: str):
        """Initialize the editor and load PPTX parts into memory."""

        self.input_pptx = input_pptx
        self.files = self._load_pptx_files(input_pptx)
        self._has_changes = False


__all__ = ["PowerPointEditor", "SlideProxy"]
