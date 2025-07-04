def changeToDarkMode(parentWidget):
    changeChildrenColor(parentWidget, "#202124", "white", "gray20", "gray20")

def changeToLightMode(parentWidget):
    changeChildrenColor(parentWidget, "#F1F1F1", "#000000", "#F1F1F1", "#FFFFFF")

def changeChildrenColor(widget, bgColor, fgColor, activeBgColor, selectColor):
    widgetAttributes = widget.keys()

    if "bg" in widgetAttributes:
        widget.config(bg=bgColor)
    if "fg" in widgetAttributes:
        widget.config(fg=fgColor)
    if "activebackground" in widgetAttributes:
        widget.config(activebackground=activeBgColor)
    if "activeforeground" in widgetAttributes:
        widget.config(activeforeground=fgColor)
    if "selectcolor" in widgetAttributes:
        widget.config(selectcolor=selectColor)
    if "insertbackground" in widgetAttributes:
        widget.config(insertbackground=fgColor)

    for child in widget.winfo_children():
        changeChildrenColor(child, bgColor, fgColor, activeBgColor,selectColor)
