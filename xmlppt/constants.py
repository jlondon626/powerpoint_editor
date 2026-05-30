import re


# Namespaces
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

NS = {
    "p": P_NS,
    "c": C_NS,
    "r": R_NS,
    "pr": PR_NS,
    "cx": CX_NS,
    "a": A_NS,
    "ct": CT_NS,
    "p14": P14_NS,
}


REL_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_TYPE_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
REL_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
REL_TYPE_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_TYPE_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_TYPE_CHART_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
REL_TYPE_CHART_COLOR_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"

SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
