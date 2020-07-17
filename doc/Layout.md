# Way to create layout with Word macro

---

### Usage

### Page Layout
	Set [Size] in [PAGE LAYOUT] tab.
	Set [Margins] in [PAGE LAYOUT] tab.

### Text Box
- Select [Draw Text Box] in [Text Box] of [INSERT] tab and place it.
- Change the text box name in [Selection Pane] of [DRAWING TOOLS] [FORMAT] tab.
- Change the Outline Color and Weight in [Shape Outline] of [DRAWING TOOLS] [FORMAT] tab.
 
### Pictures
- Select [Pictures] in [INSERT] tab and place it.
- Change the picture name in [Selection Pane] of [PICTURE TOOLS] [FORMAT] tab.

### Import Macro
- Select [Macros] [View Macros] in [VIEW] tab and select [Create].
- Select [File] [Import File] in VBA window.
- Select LayoutMacros.bas and push [Open].

### Customize Keyboard Shortcuts
- Select “Categories:” “Macros” of [Options] [Customize Ribbon] “Keyboard shortcuts:” [Customize] in [FILE] tab.
- Select “GetLeyout” in “Macros:” click a text box “Press new shortcut keys” and input Alt+L.  Push [Assign].
- Select “ClearLeyout” in “Macros:” click a text box “Press new shortcut keys” and input Alt+C.  Push [Assign].

### Run Macro
- Input Alt+L in a page with layout to run GetLayout.
- Layout CSS will be displayed in a new text box (layout), so select the text, copy and paste to other editor.  Then save as a .css file.
- Input Alt+C to delete the text box (layout).

---
![Note JavaScript](../img/Javascript.jpg)
---

### Table of Contents
- gacco: ga028 Get Started with Physical Programming
- Small Basic Interpreter
- BASIC Parser
- Development Environment Changes
- Layout Editor
