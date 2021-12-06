#!BPY
"""
Name: 'TikZ (.tex)...'
Blender: 245
Group: 'Export'
Tooltip: 'Export selected curves as TikZ paths for use with (La)TeX'
"""

__author__ = 'Kjell Magne Fauske'
__version__ = "1.0"
__url__ = ("Documentation, http://www.fauskes.net/code/blend2tikz/documentation/",
           "Author's homepage, http://www.fauskes.net/",
           "TikZ examples, http://www.fauskes.net/pgftikzexamples/")

__bpydoc__ = """\
This script exports selected curves and empties to TikZ format for use with TeX.
PGF and TikZ is a powerful macro package for creating high quality illustrations
and graphics for use with (La|Con)TeX.

Important: TikZ is primarily for creating 2D illustrations. This script will
therefore only export the X and Y coordinates. However, the Z coordinate is used
to determine draw order. 

Usage:

Select the objects you want to export and invoke the script from the
"File->Export" menu[1]. Alternatively you can load and run the script from
inside Blender.

A dialog box will pop up with various options:<br>

    - Draw: Insert a draw operation in the generated path.<br>
    - Fill: Insert a fill operation in the generated path.<br>
    - Transform: Apply translation and scale transformations.<br>
    - Materials: Export materials assigned to curves.<br>
    - Empties: Export empties as named coordinates.<br>
    - Only properties: Use on the style property of materials if set.<br>
    - Standalone: Create a standalone document.<br>
    - Only code: Generate only code for drawing paths.<br>
    - Clipboard: Copy generated code to the clipboard. <br>

Properties:

If an object is assigned a ID property or game property named 'style' of type
string, its value will be added to the path as an option. You can use the
Help->Property ID browser to set this value, or use the Logic panel to
add a game property. 

Materials:

The exporter has basic support for materials. By default the material's RGB
value is used as fill or draw color. You can also set the alpha value for
transparency effects.

An alternative is to specify style options
directly by putting the values in a 'style' property assigned to the material.
You can use the Help->Property ID browser to set this value.

Issues:<br>

- Only bezier and polyline curves are supported.<br>
- A full Python install is required for clipboard support on Windows. Other platforms
need the standard subprocess module (requires Python 2.4 or later). Additionally:<br>
    * Windows users need to install the PyWin32 module.<br>
    * Unix-like users need the xclip command line tool or the PyGTK_ module installed.<br>
    * OS X users need the pbcopy command line tool installed.<br>

[1] Requires you to put the script in Blender's scripts folder. Blender will
then automatically detect the script.
"""

import bpy
from bpy.props import StringProperty, IntProperty, BoolProperty
from bpy_extras.io_utils import ExportHelper
import subprocess
import dependencie_importer
import math
from collections import namedtuple
from textwrap import wrap
import functools

Dependency = namedtuple("Dependency", ["module", "package", "name"])
dependencies = (Dependency(module="clipboard", package=None, name=None),)
dependencies_installed = False

# Curve types
TYPE_POLY = 0
TYPE_BEZIER = 1
TYPE_NURBS = 4

R2D = 180.0 / math.pi

# Start of configuration section -------

# Templates
standalone_template = r"""
\documentclass{article}
\usepackage{tikz}
%(preamble)s
%(materials)s
\begin{document}
\begin{tikzpicture}
%(pathcode)s
\end{tikzpicture}
\end{document}
"""

fig_template = r"""
%(materials)s
\begin{tikzpicture}
%(pathcode)s
\end{tikzpicture}
"""

# config options:


bl_info = {
    "name": "Export Curve as TikZ",
    "author": "Addon: Kjell Magne Fauske, Revision: Benjamin Pionczewski",
    "version": (2, 0),
    "blender": (3, 0, 0),
    "location": "File > Export > TikZ (.tex)",
    "description": "Export Curves as TikZ (.tex)",
    "warning": "",
    "doc_url": "http://www.fauskes.net/code/blend2tikz/documentation/",
    "category": "Import-Export",
}

# Start of GUI section ------------------------------------------------

# End of GUI section ----------------------

# End of configuration section ---------
X = 0
Y = 1

used_materials = {}


# Utility functions

def nsplit(seq, n=2):
    """Split a sequence into pieces of length n

    If the lengt of the sequence isn't a multiple of n, the rest is discareded.
    Note that nsplit will strings into individual characters.

    Examples:
    >>> nsplit('aabbcc')
    [('a', 'a'), ('b', 'b'), ('c', 'c')]
    >>> nsplit('aabbcc',n=3)
    [('a', 'a', 'b'), ('b', 'c', 'c')]

    # Note that cc is discarded
    >>> nsplit('aabbcc',n=4)
    [('a', 'a', 'b', 'b')]
    """
    return [xy for xy in zip(*[iter(seq)] * n)]


def mreplace(s, chararray, newchararray):
    for a, b in zip(chararray, newchararray):
        s = s.replace(a, b)
    return s


def tikzify(s):
    if s.strip():
        return mreplace(s, r'\,:.', '-+_*')
    else:
        return ""


def cmp(a, b):
    return (a > b) - (a < b)


def copy_to_clipboard(text):
    """Copy text to the clipboard

    Returns True if successful. False otherwise.

    Works on Windows, *nix and Mac. Tries the following:
    1. Use the win32clipboard module from the win32 package.
    2. Calls the xclip command line tool (*nix)
    3. Calls the pbcopy command line tool (Mac)
    4. Try pygtk
    """
    # try windows first
    try:
        import clipboard
        clipboard.copy(text)
        return True
    except:
        return False


def get_property(obj, name):
    """Get named object property

    Looks first in custom properties, then game properties. Returns a list.
    """
    prop_value = []
    try:
        prop_value.append(obj.properties[name])
    except:
        pass
    try:
        # look for game properties
        prop = obj.getProperty(name)
        if prop.type == "STRING" and prop.data.strip():
            prop_value.append(prop.data)
    except:
        pass
    return prop_value


def get_material(material):
    """Convert material to TikZ options"""
    if not material:
        return ""
    opts = ""
    mat_name = tikzify(material.name)
    used_materials[mat_name] = material
    return mat_name


def write_materials(used_materials, ONLY_PROPERTIES):
    """Return code for the used materials
    :param ONLY_PROPERTIES:
    """
    c = "% Materials section \n"
    for material in used_materials.values():
        mat_name = tikzify(material.name)
        matopts = ''
        proponly = ONLY_PROPERTIES
        try:
            proponly = material.properties['onlyproperties']
            if proponly and type(proponly) == str:
                proponly = proponly.lower() not in ('0', 'false')
        except:
            pass
        try:
            matopts = material.properties['style']
        except:
            pass

        rgb = material.rgbCol
        spec = material.specCol
        alpha = material.alpha
        flags = material.getMode()
        options = []

        if not (proponly and matopts):
            c += "\\definecolor{%s_col}{rgb}{%s,%s,%s}\n" \
                 % tuple([mat_name] + rgb)
            options.append('%s_col' % mat_name)
            if alpha < 1.0:
                options.append('opacity=%s' % alpha)
        if matopts:
            options += [matopts]
        c += "\\tikzstyle{%s}= [%s]\n" % (mat_name, ",".join(options))
    return c


def write_object(obj, empties, USE_PLOTPATH, WRAP_LINES, DRAW_CURVE, FILL_CLOSED_CURVE, TRANSFORM_CURVE,
                 EXPORT_MATERIALS, EMPTIES):
    """Write Curves
    :param DRAW_CURVE:
    :param FILL_CLOSED_CURVE:
    :param TRANSFORM_CURVE:
    :param EXPORT_MATERIALS:
    :param EMPTIES:
    :param USE_PLOTPATH:
    :param WRAP_LINES:
    """
    s = ""
    name = obj.name

    x, y, z = obj.location
    rot = obj.rotation_euler
    scale_x, scale_y, scale_z = obj.scale

    rot_z = rot.z
    if obj.type not in ["CURVE", "Empty"]:
        return s

    ps = ""
    if obj.type == 'CURVE':
        curvedata = obj.data
        s += "%% %s\n" % name
        for spline in curvedata.splines:
            if spline.type == "BEZIER":
                knots = []
                handles = []
                # Build lists of knots and handles
                for point in spline.bezier_points:
                    h1 = point.handle_left
                    knot = point.co
                    h2 = point.handle_right
                    handles.extend([h1, h2])
                    knots.append("(+%.4f,+%.4f)" % (knot.x, knot.y)) # todo: fix string conversion

                if spline.use_cyclic_u:
                    # The curve is closed.
                    # Move the first handle to the end of the handles list.
                    handles = handles[1:] + [handles[0]]
                    # Repeat the first knot at the end of the knot list
                    knots.append(knots[0])
                else:
                    # We don't need the first and last handles since the curve is
                    # not closed.
                    handles = handles[1:-1]
                hh = []
                for h1, h2 in nsplit(handles, 2):
                    hh.append("controls (+%.4f,+%.4f) and (+%.4f,+%.4f)" % (h1.x, h1.y, h2.x, h2.y))

                ps += "%s\n" % knots[0]
                for h, k in zip(hh, knots[1:]):
                    ps += "  .. %s .. %s\n" % (h, k)
                if spline.use_cyclic_u:
                    ps += "  -- cycle\n"
            elif spline.type == "POLY":
                coords = [f"(+{point.co.x}.4f,+{point.co.y}.4f)" for point in spline.SplinePoint]

                if USE_PLOTPATH:
                    plotopts = get_property(obj, 'plotstyle')
                    if plotopts:
                        poptstr = "[%s]" % ",".join(plotopts)
                    else:
                        poptstr = ''
                    ps += " plot%s coordinates {%s}" % (poptstr, " ".join(coords))
                    if spline.use_cyclic_u:
                        ps += " -- cycle"
                    if WRAP_LINES:
                        ps = "\n".join(wrap(ps, 80, subsequent_indent="  ", break_long_words=False))

                else:
                    if spline.use_cyclic_u:
                        coords.extend([coords[0], 'cycle\n  '])
                    # Join the coordinates. Could have used "--".join(coords), but
                    # have to add some logic for pretty printing.
                    if WRAP_LINES:
                        ps += "%s\n  " % coords[0]
                        i = 0
                        for c in coords[1:]:
                            i += 1
                            if i % 3:
                                ps += "-- %s" % c
                            else:
                                ps += "  -- %s\n  " % c
                    else:
                        ps += "%s" % " -- ".join(coords)
            else:
                continue

        if not ps:
            return s
        options = []
        if DRAW_CURVE:
            options += ['draw']
        if FILL_CLOSED_CURVE:
            if ps.find('cycle') > 0:
                options += ['fill']
        if TRANSFORM_CURVE:
            if x != 0: options.append('xshift=%.4fcm' % x)
            if y != 0: options.append('yshift=%.4fcm' % y)
            if rot_z != 0: options.append('rotate=%.4f' % rot_z)
            if scale_x != 1: options += ['xscale=%.4f' % scale_x]
            if scale_y != 1: options += ['yscale=%.4f' % scale_y]
        if EXPORT_MATERIALS:
            try:
                materials = obj.data.materials
            except:
                materials = []
            if materials:
                # pick first material
                for mat in materials:
                    if mat:
                        matopts = get_material(mat)
                        options.append(matopts)
                        break
        extraopts = get_property(obj, 'style')
        if extraopts:
            options.extend(extraopts)

        optstr = ",".join(options)
        print("Options: ", options)
        emptstr = ""
        if EMPTIES:
            if obj in empties:
                for empty in empties[obj]:
                    # Get correct coordinate relative to the parent
                    if TRANSFORM_CURVE:
                        ex, ey, ez = (empty.mat * (obj.mat.copy()).invert()).translationPart
                    else:
                        ex, ey, ez = (empty.matrix - obj.matrix).translation
                    emptstr += "  (+%.4f,+%.4f) coordinate (%s)\n" \
                               % (ex, ey, empty.name)

        if not WRAP_LINES:
            ps = ' '.join(ps.replace('\n', ' ').split())
        if len(optstr) > 50 or emptstr:
            s += "\\path[%s]\n%s  %s;\n" % (optstr, emptstr, ps.rstrip())
        else:
            s += "\\path[%s] %s;\n" % (optstr, ps.rstrip())
    elif obj.type == 'Empty' and EMPTIES and not obj.parent:
        x, y, z = obj.matrix_world.translation
        s += "\\coordinate (%s) at (%.4f,%.4f);\n" % (tikzify(obj.name), x, y)

    return s




# Start of script -----------------------------------------------------
def write_tex(context, filepath, USE_PLOTPATH, WRAP_LINES, DRAW_CURVE, FILL_CLOSED_CURVE,
              TRANSFORM_CURVE, EXPORT_MATERIALS, EMPTIES, ONLY_PROPERTIES, STANDALONE, CODE_ONLY, CLIPBOARD_OUTPUT):
    # Ensure that at leas one object is selected

    if len(context.selected_objects) == 0:
        # no objects selected. Print error message and quit
        return 'ERROR: Please select at least one curve'
    else:
        def z_comp(a, b):
            x, y, z1 = a.matrix_world.translation
            x, y, z2 = b.matrix_world.translation
            return cmp(z1, z2)

        # get all selected objects
        objects = context.selected_objects
        # get current scene
        scn = context.scene
        # iterate over each object
        code = ""
        # Find all empties with parents
        empties_wp = [obj for obj in objects if obj.type == 'Empty' and obj.parent]
        empties_dict = {}
        for empty in empties_wp:
            if empty.parent in empties_dict:
                empties_dict[empty.parent] += [empty]
            else:
                empties_dict[empty.parent] = [empty]

        for obj in sorted(objects, key=functools.cmp_to_key(z_comp)):
            code += write_object(obj, empties_dict, USE_PLOTPATH, WRAP_LINES, DRAW_CURVE, FILL_CLOSED_CURVE,
                                 TRANSFORM_CURVE, EXPORT_MATERIALS, EMPTIES)
        s = ""
        if EXPORT_MATERIALS:
            matcode = write_materials(used_materials, ONLY_PROPERTIES)
        else:
            matcode = ""

        try:
            preamblecode = scn.properties['preamble']
        except:
            preamblecode = ''
        templatevars = dict(pathcode=code, preamble=preamblecode, materials=matcode)
        if STANDALONE:
            extra = ""
            try:
                preambleopt = scn.properties['preamble']
                templatevars['preamble'] = str(preambleopt)
            except:
                pass
            template = standalone_template

        elif CODE_ONLY:
            template = "%(pathcode)s"
        else:
            template = fig_template

        s = template % templatevars
        if not CLIPBOARD_OUTPUT:
            print(s)
            try:
                with open(filepath, 'w') as f:
                    # write header to file
                    f.write('%% Generated by tikz_export.py v %s \n' % str(bl_info["version"]))
                    f.write(s)
                    print(f"Code written to {filepath}")
                return "File sucesfully writen"
            except Exception as error:
                print("Wirte File Fail:", error)
                return "Failed to write file"
        else:
            success = copy_to_clipboard(s)
            if not success:
                print("Failed to copy code to the clipboard")
                return "Failed  to write to clipboard"

    return "tikz_export ended ..."


class TechFileExport(bpy.types.Operator, ExportHelper):
    bl_idname = "export_tex.curves"
    bl_label = "Write Curves as TikZ"

    filename_ext = ".tex"

    STANDALONE: BoolProperty(name="Standalone",
                             description="Output standalone document",
                             default=True)
    DRAW_CURVE: BoolProperty(name="Draw",
                             description="Draw curves",
                             default=True)
    FILL_CLOSED_CURVE: BoolProperty(name="Fill",
                                    description="Fill closed curves",
                                    default=False)
    TRANSFORM_CURVE: BoolProperty(name="Transform",
                                  description="Apply Transformations",
                                  default=False)
    USE_PLOTPATH: BoolProperty(name="Use Plot Paths",
                               description="Use the plot path operations for polyline",
                               default=False)
    EXPORT_MATERIALS: BoolProperty(name="Materials",
                                   description="Apply materials to curves",
                                   default=False)
    EMPTIES: BoolProperty(name="Empties",
                          description="Export empties",
                          default=False)
    ONLY_PROPERTIES: BoolProperty(name="Only properties",
                                  description="Use only properties for materials with the style property set",
                                  default=False)
    CODE_ONLY: BoolProperty(name="Only Code",
                            description="Output pathcode only",
                            default=False)
    CLIPBOARD_OUTPUT: BoolProperty(name="Clipboard",
                                   description="Copy code to clipboard",
                                   default=False)
    WRAP_LINES: BoolProperty(name="Wrap lines",
                             description="Wrap long lines",
                             default=True)

    def execute(self, context):
        status = write_tex(context, self.filepath, self.USE_PLOTPATH, self.WRAP_LINES, self.DRAW_CURVE,
                           self.FILL_CLOSED_CURVE,
                           self.TRANSFORM_CURVE, self.EXPORT_MATERIALS, self.EMPTIES, self.ONLY_PROPERTIES,
                           self.STANDALONE, self.CODE_ONLY,
                           self.CLIPBOARD_OUTPUT)
        print(status)
        return {"FINISHED"}

    def invoke(self, context, event):
        wm = context.window_manager
        wm.fileselect_add(self)
        return {'RUNNING_MODAL'}




class PT_warning_panel(bpy.types.Panel):
    """
    Dependency Installer Panel
    """
    bl_label = "Missing dependencies"
    bl_category = "Install dependencies"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"

    @classmethod
    def poll(self, context):
        return not dependencies_installed

    def draw(self, context):
        layout = self.layout

        lines = [f"Please install the missing dependencies for the \"{bl_info.get('name')}\" add-on.",
                 f"1. Open the preferences (Edit > Preferences > Add-ons).",
                 f"2. Search for the \"{bl_info.get('name')}\" add-on.",
                 f"3. Open the details section of the add-on.",
                 f"4. Click on the \"{OT_install_dependencies.bl_label}\" button.",
                 f"   This will download and install the missing Python packages, if Blender has the required",
                 f"   permissions.",
                 f"If you're attempting to run the add-on from the text editor, you won't see the options described",
                 f"above. Please install the add-on properly through the preferences.",
                 f"1. Open the add-on preferences (Edit > Preferences > Add-ons).",
                 f"2. Press the \"Install\" button.",
                 f"3. Search for the add-on file.",
                 f"4. Confirm the selection by pressing the \"Install Add-on\" button in the file browser."]

        for line in lines:
            layout.label(text=line)


class OT_install_dependencies(bpy.types.Operator):
    """
    Dependency Installer
    """
    bl_idname = "example.install_dependencies"
    bl_label = "Install dependencies"
    bl_description = ("Downloads and installs the required python packages for this add-on. "
                      "Internet connection is required. Blender may have to be started with "
                      "elevated permissions in order to install the package")
    bl_options = {"REGISTER", "INTERNAL"}

    @classmethod
    def poll(self, context):
        # Deactivate when dependencies have been installed
        return not dependencies_installed

    def execute(self, context):
        try:
            dependencie_importer.install_pip()
            for dependency in dependencies:
                dependencie_importer.install_and_import_module(module_name=dependency.module,
                                                               package_name=dependency.package,
                                                               global_name=dependency.name)
        except (subprocess.CalledProcessError, ImportError) as err:
            self.report({"ERROR"}, str(err))
            return {"CANCELLED"}

        global dependencies_installed
        dependencies_installed = True

        # Register the panels, operators, etc. since dependencies are installed
        bpy.utils.register_class(TechFileExport)
        bpy.types.TOPBAR_MT_file_export.append(menu_export)

        return {"FINISHED"}


class preferences(bpy.types.AddonPreferences):
    bl_idname = __name__

    def draw(self, context):
        layout = self.layout
        layout.operator(OT_install_dependencies.bl_idname, icon="CONSOLE")


preference_classes = (PT_warning_panel,
                      OT_install_dependencies,
                      preferences)


def menu_export(self, context):
    import os
    default_path = os.path.splitext(bpy.data.filepath)[0] + ".tex"
    self.layout.operator(TechFileExport.bl_idname, text="Curves to Tex (.tex)").filepath = default_path


def register():
    global dependencies_installed
    dependencies_installed = False

    for cls in preference_classes:
        bpy.utils.register_class(cls)

    try:
        for dependency in dependencies:
            dependencie_importer.import_module(module_name=dependency.module, global_name=dependency.name)
        dependencies_installed = True
    except ModuleNotFoundError:
        # Don't register other panels, operators etc.
        return

    bpy.utils.register_class(TechFileExport)
    bpy.types.TOPBAR_MT_file_export.append(menu_export)


def unregister():
    for cls in preference_classes:
        bpy.utils.unregister_class(cls)

    if dependencies_installed:
        bpy.utils.unregister_class(TechFileExport)
        bpy.types.TOPBAR_MT_file_export.remove(menu_export)


if __name__ == "__main__":
    register()
