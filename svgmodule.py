# -*- coding: cp1251 -*-
# svgmodule.py
# last revision: Mar 2, 2022
"""
Simply module for converting from KOMPAS-fragment (*.frw)
to Scalable Vector Graphics file (*.svg) ver. 1.1.
"""

__version__ = '0.4' # With supporting of CSS-styles for curves

import pythoncom
from win32com.client import Dispatch, gencache
from math import *
#from LDefin2D import *

## Get KOMPAS interfaces
KAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
iApplication = Dispatch('Kompas.Application.7')
KAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)
iKompasObj = Dispatch('KOMPAS.Application.5', None, KAPI5.KompasObject.CLSID)
iDocument2D = iKompasObj.ActiveDocument2D()
CONST2D = gencache.EnsureModule('{75C9F5D0-B5B8-4526-8681-9903C567D2ED}', 0, 1, 0).constants
pMath2D = iKompasObj.GetMathematic2D()
# aliases of math-function
Rotate, GetCurveLen = pMath2D.ksRotate, pMath2D.ksGetCurvePerimeter
Dist = pMath2D.ksDistancePntPnt

## XML-constants
X_IMAGE = 'image'
X_PATH  = 'path'
X_DEFS  = 'defs'
X_TEXT  = 'text'
X_CIRCL = 'circle'
X_GROUP = 'g'
X_MARK  = 'marker'

## Inner constants
ID_TEXT         = 'text#%d'
ID_DPOINT       = 'dpoint#%d'
ID_DARROW       = 'darrow#%d'
ID_DNOTCH       = 'dnotch#%d'
ID_MACRO        = 'Macro#%d'
ID_ARCUS        = 'arcus#%d'
ID_BEZIER       = 'bezier#%d'
ID_CIRCLE       = 'circle#%d'
ID_CONTOUR      = 'contour#%d'
ID_ELLIPSE      = 'ellipse#%d'
ID_ELLIPSEARC   = 'ellipsearc#%d'
ID_HATCH        = 'hatch#%d'
ID_COLOURING    = 'colouring#%d'
ID_LINESEG      = 'line#%d'
ID_NURBS        = 'nurbs#%d'
ID_POINT        = 'point#%d'
ID_POLYLINE     = 'polyline#%d'
ID_POLYGON      = 'polygon#%d'
ID_CMARK        = 'cmark#%d'
ID_LINEDIM      = 'linedimension#%d'
ID_BRLINEDIM    = 'breaklinedim#%d'
ID_DIAMETRALDIM = 'diametral#%d'
ID_RADIALDIM    = 'radialdim#%d'
ID_ANGLEDIM     = 'angledim#%d'
ID_LEADER       = 'leader#%d'
ID_BRANCH       = 'lbranch#%d'
ID_IMAGE        = 'image#%d'
ID_DEFS         = 'defs#%d'
ID_MARROWST     = 'ArrowStart'
ID_MARROWEND    = 'ArrowEnd'
ID_MARROWSTOUT  = 'ArrowStartOut'
ID_MARROWENDOUT = 'ArrowEndOut'
ID_MPOINT       = 'Point'
ID_MNOTCH       = 'RNotch'
ID_MLNOTCH      = 'LNotch'
ID_GENOBJ       = 'object#%d'

DL_START = 0
DL_END = 1
TXT_ALLOCATION = {0: 'start', 1: 'middle', 2: 'end'}

DEFAULT_PARAMS = {
    'WIDTH':            210,
    'HEIGHT':           297,
    'TOLERANCE':        0.1,
    'S_LINE':           0.48,
    'ALL_THIN':         False,
    'BLACK_WHITE':      False,
    'KEGEL':            5.0,
    'D_POINT':          1.2,
    'ARROWSTYLE':       'stroke:none',
    'OVERHANG':         2.0,
    'LineSegments':     True,
    'Circles':          True,
    'Arcs':             True,
    'DrawTexts':        True,
    'Points':           True,
    'Hatches':          True,
    'Curves':           True,
    'Dimensions':       True,
    'DrawContours':     True,
    'Leaders':          True,
    'PolyLines2D':      True,
    'Ellipses':         True,
    'EllipsesArcs':     True,
    'Polygons':         True,
    'CentreMarks':      True,
    'Rasters':          True,
    'AxisLines':        True,
    'MacroObjects':     True,
    'FIT_CANVAS':       True,
    'CANVAS_GAP':       1.0,
    'MATH_MODE':        True
    }

class CElement:
    """ Simple element of XML-tree """
    def __init__(self, tag, id='', text=''):
        self.tag = tag
        self.level = 0
        self.subelems = []
        self.attribs = []
        self.id = id
        self.text = text
    def append(self, element):
        """ Add subelement """
        if element:
            element.parent = self
            element.incLevel()
            self.subelems.append(element)
    def setAttribs(self, attributes):
        """ Assign attributes """
        if self.attribs:
            for at in attributes:
                self.attribs.append(at)
        else:
            self.attribs = list(attributes)
    def addAttrib(self, attribute):
        """ Add attribute """
        self.attribs.append(attribute)
    def incLevel(self):
        self.level += 1
        for element in self.subelems:
            element.incLevel()
    def printf(self):
        """ Return XML-tag as string """
        indent = '\t'*self.level
        stag = '%s<%s' %(indent, self.tag) # open tag
        if self.id:
            stag += ' id="%s"' %self.id
        for key, val in self.attribs:   # make attributes
            stag += '\n%s\t%s="%s"' %(indent, key, val)
        if self.subelems or self.text:
            stag += '>\n'
            for element in self.subelems:   # make subelements
                stag += element.printf()
            if self.text:
                stag += self.text
            stag += '%s</%s>\n' %(indent, self.tag)
            return stag
        stag += ' />\n' # close tag (simple)
        return stag

class CNode:
    """ Node of N-ary tree """
    def __init__(self, data):   #, parent=None):
        self.data = data    # type<data> == 'IContour'
        self.childs = []
        self.level = 0
        self._makeLowPoly(data)
    def addNode(self, node):
        if self.checkInside(node):  # if new node instant to current --> add new node to current
            node.incLevel()
            self._addChild(node)
            return True # yield as flag of success
        return False
    def _addChild(self, node):
        if self.childs:
            rmlist = []
            for child in self.childs:
                if node.addNode(child):
                    rmlist.append(child)    # search in childs of current root childs of new node
            self._remChilds(rmlist)
            for child in self.childs:
                if child.addNode(node):
                    return  # node hawe only 1 parent
        self.childs.append(node)
    def _remChilds(self, remlist):
        for node in remlist:
            self.childs.remove(node)
    def _makeLowPoly(self, IContour):
        self.lowpoly = []
        for iSegment in IContour.TmpObjects:
            if iSegment.Type == CONST2D.ksObjectLineSegment:
                self.lowpoly.append({'x': iSegment.X1, 'y': iSegment.Y1})
            else:
                point_arr = pMath2D.ksPointsOnCurve(iSegment.Reference, 2)
                item = iKompasObj.GetParamStruct(CONST2D.ko_MathPointParam)
                point_arr.ksGetArrayItem(0, item)
                self.lowpoly.append({'x': item.x, 'y': item.y})
                point_arr.ksGetArrayItem(1, item)
                self.lowpoly.append({'x': item.x, 'y': item.y})
    def incLevel(self):
        self.level += 1
        for node in self.childs:
            node.incLevel()
    def getMaxLevel(self):
        lvl = [self.level]
        for node in self.childs:
            lvl.append(node.getMaxLevel())
        return max(lvl)
    def getNodesByLevel(self, level):
        if self.level == level:
            return [self]
        elif self.level > level:
            return []
        rv = []
        for node in self.childs:
            rv += node.getNodesByLevel(level)
        return rv
    def checkInside(self, obj):
        for v in obj.lowpoly:
            if iDocument2D.ksIsPointInsideContour(self.data.Reference, v['x'], v['y'], 0.01) == 3:
                return True # enough only one match
        return False

def MakeForest(src):
    """ Создаёт "лес" из исходного списка узлов """
    # TODO: remove recursion, replace to stack
    skipped, temp = [], []
    root = src.pop()
    while src:
        node = src.pop()        # pop from end is faster
        if root.addNode(node):  # this case finds children
            continue
        elif node.addNode(root):    # this case finds a parent
            root = node # current node --> current tree
            while skipped:
                node = skipped.pop()
                if root.addNode(node):
                    continue
                else:
                    temp.append(node)
            if temp:
                skipped = temp[:]   # COPY 'temp' to 'skipped'
                temp = []
        else:   # this case finds nodes from other trees
            skipped.append(node)
    if skipped:
        # TODO: use the function extend() from built-in class list
        return [root] + MakeForest(skipped) # FUCKING RECURSION!!!
    else:
        return [root]

## ==== HELPER FUNCTIONS ====
def ColorLong2RGB(color, bw=False):
    """ Convert color from LONG to RGB (as hex-string) """
    # source from StackOverflow.com
    R = color & 0xFF
    G = (color >> 8) & 0xFF
    B = (color >> 16) & 0xFF
    if bw:  # convert to grayscale, HDTV format
        R = G = B = int(0.2126*R + 0.7152*G + 0.0722*B)
    return '%02x%02x%02x' %(R, G, B)

def getMIME(path):
    """ Primitive algorithm for reading MIME """
    src = path.strip().lower()
    if src:
        if src.endswith('.bmp'):
            return 'image/x-ms-bmp'
        elif src.endswith('.gif'):
            return 'image/gif'
        elif src.endswith(('jpg', 'jpeg', 'jfif', 'jpe')):
            return 'image/jpeg'
        elif src.endswith(('png', 'tga')):
            return 'image/png'
        elif src.endswith(('tiff', 'tif')):
            return 'image/tiff'
    return ''

def GetDimText(IDimension, nomin=True, pref=True, suff=True):
    """ Get text from any dimension object """
    txt = KAPI7.IDimensionText(IDimension)
    rv = pref*txt.Prefix.Str + nomin*txt.NominalText.Str + suff*txt.Suffix.Str
    return rv.strip()

def ScaleCoordinates(x, y, cx, cy, m=1.0, dm=None):
    """ To scale Cartesian coordinates [x, y] """
    if dm: m = 1 + dm/sqrt((cx - x)**2 + (cy - y)**2)
    x_new = cx + m*(x - cx)
    y_new = cy + m*(y - cy)
    return x_new, y_new

# BASE CLASS OF CONVERTER: # NB! DON'T USE THIS CLASS!
class _CFakeConverter():
    params = DEFAULT_PARAMS
    _idnumb = 100
    defs = None
    markersDB = {}  # database of markers
    cstylesDB = {}  # database of curve's styles
    patternDB = {}  # database of hatch patterns
    sCSS = '' #'path,line,rect,circle,ellipse {}\n'   # directly, description of style
    _X_start = 0.0
    _Y_start = 0.0

    def getID(self):
        self._idnumb += 1
        return self._idnumb

    def GetXY(self, x, y, transform):
        """ Converts coordinates from Cartesian to fucking screen CS and apply transformation """
        if transform:
            _, x, y = Rotate(x, y, 0.0, 0.0, transform['ang'])
            x += transform['x']
            y += transform['y']
        return x - self._X_start, self.params['HEIGHT'] - y + self._Y_start

    def GetAngle(self, angle, transform):
        """ Converts angle from physic to screen CS and apply transformation """
        if transform: return -angle - transform['ang']
        else: return -angle

    def GetTextPlacement(self, PCompositeObj, transform):
        """ Helper function for getting angle and coords of dimensions text per screen CS """
        pTmpGroup = iDocument2D.ksDecomposeObj(PCompositeObj, 1, 0.5, 0)
        iIterator = iKompasObj.GetIterator()
        iIterator.ksCreateIterator(TEXT_OBJ, pTmpGroup)
        iText = iKompasObj.TransferReference(iIterator.ksMoveIterator('F'), 0)
        Angle = self.GetAngle(iText.Angle, transform)
        X, Y = self.GetXY(iText.X, iText.Y, transform)
        iIterator.ksDeleteIterator()
        iDocument2D.ksClearGroup(pTmpGroup, True)
        return Angle, X, Y

    def MakeArcus(self, IArc, transform, tracing=True, full=True):
        """ Return XML-path of arc """
        X1, Y1 = self.GetXY(IArc.X1, IArc.Y1, transform)
        X2, Y2 = self.GetXY(IArc.X2, IArc.Y2, transform)
        R = IArc.Radius
        b = int(GetCurveLen(IArc.Reference, 0x1) > pi*R)
        m = int(IArc.Direction) # True/False --> 1/0
        if tracing == False:
            X1, Y1, X2, Y2 = X2, Y2, X1, Y1
            b, m = int(not b), int(not m)
        sPath = 'A %.4f,%.4f 0 %d %d %.4f,%.4f' %(R, R, b, m, X2, Y2)
        if full:    # add first point
            return 'M %.4f,%.4f %s' %(X1, Y1, sPath)
        return sPath

    def MakeEllipseArc(self, IEllipseArc, transform, full=True):
        """ Return XML-path of ellipse arc """
        Xc, Yc = self.GetXY(IEllipseArc.Xc, IEllipseArc.Yc, transform)
        Ang = self.GetAngle(IEllipseArc.Angle, transform)
        Ra, Rb = IEllipseArc.SemiAxisA, IEllipseArc.SemiAxisB
        T1, T2 = IEllipseArc.T1, IEllipseArc.T2
        X20 = Ra*cos(-T2) + Xc
        Y20 = Rb*sin(-T2) + Yc
        big = int((T2 - T1) > pi)
        mir = int(not IEllipseArc.Direction)    # Direction is reversed
        _, X2, Y2 = Rotate(X20, Y20, Xc, Yc, Ang)
        sPath = 'A %.4f,%.4f %.4f %d %d %.4f,%.4f ' %(Ra, Rb, Ang, big, mir, X2, Y2)
        if full:    # add first point
            X10 = Ra*cos(-T1) + Xc
            Y10 = Rb*sin(-T1) + Yc
            _, X1, Y1 = Rotate(X10, Y10, Xc, Yc, Ang)
            return 'M %.4f,%.4f %s' %(X1, Y1, sPath)
        return sPath

    def ApproxNurbs(self, INurbs, transform, full=True):
        """ Convert NURBS to polyline (XML-path not support NURBS) """
        # Пока что я не могу реализовать преобразование NURBS в сплайн Безье,
        # поэтому приходится ограничиться кусочно-линейной аппроксимацией
        # WARNING: this method is VERY-VERY slow
        pTmpGroup = iDocument2D.ksDecomposeObj(INurbs.Reference, 1, self.params['TOLERANCE'], 0)
        iTmpGroup = iKompasObj.TransferReference(pTmpGroup, 0)
        if INurbs.Closed: full = False  # sic!
        sPath = ''
        try:
            for seg in iTmpGroup.Objects():
                if full:
                    sPath += 'M %.4f,%.4f' %self.GetXY(seg.X1, seg.Y1, transform)
                    full = False
                sPath += ' L %.4f,%.4f' %self.GetXY(seg.X2, seg.Y2, transform)
        except TypeError:
            seg = iTmpGroup.Objects()   # if segment is singular
            if full:
                sPath += 'M %.4f,%.4f' %self.GetXY(seg.X1, seg.Y1, transform)
            sPath += ' L %.4f,%.4f' %self.GetXY(seg.X2, seg.Y2, transform)
        iTmpGroup.Delete()
        return sPath + ' Z'*INurbs.Closed

    def MakeBezier(self, IBezier, transform, full=True):
        """ Create XML-path for Bezier """
        raw = IBezier.Points(True)
        mx, my = raw[2:4]   # base coords of vertex #0
        # recombine
        if IBezier.Closed:
            full = True # на всякий случай
            raw = raw[4:] + raw[:4]
        else:
            raw = raw[4:-2]   # exclude last point
        sPath = ''
        for i in range(0, len(raw), 6):    # transform and make path
            X1, Y1 = self.GetXY(raw[i], raw[i+1], transform)
            X2, Y2 = self.GetXY(raw[i+2], raw[i+3], transform)
            X, Y = self.GetXY(raw[i+4], raw[i+5], transform)
            sPath += ' C %.4f,%.4f %.4f,%.4f %.4f,%.4f' %(X1, Y1, X2, Y2, X, Y)
        if full:
            X, Y = self.GetXY(mx, my, transform)
            sPath = 'M %.4f,%.4f %s%s' %(X, Y, sPath, ' Z '*IBezier.Closed)
        return sPath

    def ReadContour(self, IContour, dir, transform):
        """ Convert drawing contour to XML-path """
        sPath = ''
        Segments = IContour.TmpObjects  # get tuple of contour segments
        if len(Segments) == 1:  # circular/elliptic/closed NURBS contour contain only 1 segment
            iSegment = Segments[0]
            try:    # TODO: rewrite this fucking block
                Xc, Yc = self.GetXY(iSegment.Xc, iSegment.Yc, transform)
                if iSegment.Type == CONST2D.ksObjectCircle: # from enum KompasAPIObjectTypeEnum
                    RA, RB = iSegment.Radius, iSegment.Radius
                    Angle = 0.0
                    X1, Y1 = Xc + RA, Yc
                    X2, Y2 = Xc - RA, Yc
                elif iSegment.Type == CONST2D.ksObjectEllipse:
                    RA, RB = iSegment.SemiAxisA, iSegment.SemiAxisB
                    Angle = self.GetAngle(iSegment.Angle, transform) # [deg]
                    # transform
                    _, X1, Y1 = Rotate(Xc + RA, Yc, Xc, Yc, Angle)
                    _, X2, Y2 = Rotate(X1, Y1, Xc, Yc, 180.0)
                mirr = int(not dir)
                # represent circle/ellipse as two arcs/elliptic arcs
                sPath += 'M %.4f,%.4f A %.4f,%.4f %.4f 0 %d %.4f,%.4f' %(X1, Y1, RA, RB, Angle, mirr, X2, Y2)
                sPath += ' A %.4f,%.4f %.4f 0 %d %.4f,%.4f Z ' %(RA, RB, Angle, mirr, X1, Y1)
            except:
                sPath += self.ApproxNurbs(iSegment, transform)
        else:   # if contour is composite...
            first_seg = True    # flag of 1st segment
            for iSegment in Segments:   # ...enumerate all segments
                ## processing segments depedent from Type
                if iSegment.Type == CONST2D.ksObjectLineSegment:
                    if first_seg:
                        X1, Y1 = self.GetXY(iSegment.X1, iSegment.Y1, transform)
                        sPath += 'M %.4f,%.4f' %(X1, Y1)
                        first_seg = False   # reset flag of 1st segment
                    X2, Y2 = self.GetXY(iSegment.X2, iSegment.Y2, transform)
                    sPath += ' L %.4f,%.4f ' %(X2, Y2)
                elif iSegment.Type == CONST2D.ksObjectArc:
                    sPath += self.MakeArcus(iSegment, transform, full=first_seg)
                    first_seg = False
                elif iSegment.Type == CONST2D.ksObjectBezier:
                    _, _, MX, MY, X1, Y1, X2, Y2, X, Y, _, _ = iSegment.Points(True)
                    if first_seg:
                        sPath += 'M %.4f,%.4f ' %self.GetXY(MX, MY, transform)
                        first_seg = False
                    X1, Y1 = self.GetXY(X1, Y1, transform)
                    X2, Y2 = self.GetXY(X2, Y2, transform)
                    X, Y = self.GetXY(X, Y, transform)
                    sPath += 'C %.4f,%.4f %.4f,%.4f %.4f,%.4f ' %(X1, Y1, X2, Y2, X, Y)
                elif iSegment.Type == CONST2D.ksObjectEllipseArc:
                    sPath += self.MakeEllipseArc(iSegment, transform, full=first_seg)
                    first_seg = False
                else:   # for all other segments types
                    sPath += self.ApproxNurbs(iSegment, transform, full=first_seg)
                    first_seg = False
            sPath += ' Z '  # close path if all segments is processed
        return sPath.strip()

    def _setstyle(self):
        """ dynamic initialization of styles """
        S = self.params['S_LINE']   # Thickness of main lines
        if self.params['ALL_THIN']:
            S /= 3.0
            SB = ST = S
        else:
            SB = S*1.8  # thick lines
            ST = S/3.0  # thin lines
        style = { # Name                # Kind of Dash Array,             Linewidth,      Linecolor
            0x1:  {'Name': 'Main',      'DashArray': None,               'Width': S,     'Color': '#0000ff'}, # ksCSNormal
            0x2:  {'Name': 'Thin',      'DashArray': None,               'Width': ST,    'Color': '#000000'}, # ksCSThin
            0x3:  {'Name': 'Axial',     'DashArray': '7.5,1.5,1.5,1.5',  'Width': ST,    'Color': '#ff8000'}, # ksCSAxial
            0x4:  {'Name': 'Dashed',    'DashArray': '4.0,2.0',          'Width': ST,    'Color': '#000000'}, # ksCSDashed
            0x7:  {'Name': 'Thick',     'DashArray': None,               'Width': SB,    'Color': '#009696'},    # ksCSThick
            0x8:  {'Name': 'Dash2Dots', 'DashArray': '4.0,1.0,2.0,1.0,2.0,1.0', 'Width': ST, 'Color': '#000000'}, # ksCSDash2Dots
            0x9:  {'Name': 'DashedMain','DashArray': '4.0,2.0',          'Width': S,     'Color': '#0000ff'}, # ksCSDashedNormal
            0xa:  {'Name': 'MainDashDot','DashArray': '8.0,2.0,1.0,2.0', 'Width': S,     'Color': '#0000ff'}, # ksCSNormalDashDot
            0xc:  {'Name': 'ISO02',     'DashArray': '12.0,3.0',         'Width': ST,    'Color': '#000000'}, # ksCSISO02Dashed
            0xd:  {'Name': 'ISO03',     'DashArray': '12.0,18.0',        'Width': ST,    'Color': '#000000'}, # ksCSISO03DashedLSpace
            0xe:  {'Name': 'ISO04',     'DashArray': '24.0,3.0,0.5,3.0', 'Width': ST,    'Color': '#000000'}, # ksCSISO04DashDotLDash
            0xf:  {'Name': 'ISO05',     'DashArray': '24.0,3.0,0.5,3.0,0.5,3.0', 'Width': ST, 'Color': '#000000'}, # ksCSISO05DashDotLDash2Dots
            0x10: {'Name': 'ISO06',     'DashArray': '24.0,3.0,0.5,3.0,0.5,3.0,0.5,3.0', 'Width': ST, 'Color': '#000000'}, # ksCSISO06DashDotLDash3Dots
            0x11: {'Name': 'ISO07',     'DashArray': '0.5,3.0',          'Width': ST,    'Color': '#000000'}, # ksCSISO07Dotted
            0x12: {'Name': 'ISO08',     'DashArray': '24.0,3.0,6.0,3.0', 'Width': ST,    'Color': '#000000'}, # ksCSISO08DashDotLShDashes
            0x13: {'Name': 'ISO09',     'DashArray': '24.0,3.0,6.0,3.0,6.0,3.0', 'Width': ST, 'Color': '#000000'}, # ksCSISO09DashDot1L2ShDashes
            0x14: {'Name': 'ISO10',     'DashArray': '12.0,3.0,0.5,3.0', 'Width': ST,    'Color': '#ff8080'}, # ksCSISO10DashDot
            0x15: {'Name': 'ISO11',     'DashArray': '12.0,3.0,12.0,3.0,0.5,3.0','Width': ST, 'Color': '#0000ff'}, # ksCSISO11DashDot2Dashes
            0x16: {'Name': 'ISO12',     'DashArray': '12.0,3.0,0.5,3.0,0.5,3.0', 'Width': ST, 'Color': '#ff80c0'}, # ksCSISO12DashDot2Dots
            0x17: {'Name': 'ISO13',     'DashArray': '12.0,3.0,0.5,3.0,0.5,3.0,0.5,3.0', 'Width': ST, 'Color': '#ff0000'}, # ksCSISO13DashDot3Dots
            0x18: {'Name': 'ISO14',     'DashArray': '12.0,3.0,12.0,3.0,0.5,3.0,0.5,3.0', 'Width': ST, 'Color': '#ff00ff'}, # ksCSISO14DashDot2Dashes2Dots
            0x19: {'Name': 'ISO15',     'DashArray': '12.0,3.0,12.0,3.0,0.5,3.0,0.5,3.0,0.5,3.0', 'Width': ST, 'Color': '#00ff00'} # ksCSISO15DashDot2Dashes3Dots
            }
        """ while unused styles
            #0xb: {'Name': '', 'DashArray': None,           'Width': ST,   'Color': '#0000ff'}, # ksCSThinForHatch               
            #0x0: {'Name': '', 'DashArray': None,           'Width': ST,   'Color': '#0000ff'}, # ksCSUnvisible                  
            #0x5: {'Name': '', 'DashArray': None,           'Width': ST,   'Color': '#0000ff'}, # ksCSBrokenLine                
            #0x6: {'Name': '', 'DashArray': None,           'Width': ST,   'Color': '#0000ff'}  # ksCSConstruction              
            }"""
        self.ksStyle = style
        return

    def MakeStyle(self, Style, bw=False):
        """ Returns name of curve's style """
        try:
            return self.cstylesDB[Style]  # try get style
        except:
            try:
                current = self.ksStyle[Style] # get current style
            except: # TODO: maybe remove this test?
                current = self.ksStyle[0x2]   # 'Thin' is default style
                try: return self.cstylesDB[0x2]
                except: pass
            url   = current['Name']
            color = current['Color']
            width = current['Width']
            array = current['DashArray']
            self.cstylesDB[Style] = url # add style to db

            if not bw: bw = self.params['BLACK_WHITE']
            if bw: color = '#000000'    # reset color to black if b&w-mode 'ON'

            currCSS = '\t\t.%s {\n' %url
            currCSS += '\t\t\tstroke: %s;\n' %color
            currCSS += '\t\t\tstroke-width: %.4f;\n' %width
            currCSS += '\t\t\tstroke-linecap: %s;\n' %'round'
            currCSS += '\t\t\tstroke-linejoin: %s;\n' %'miter'  # It's a default value.
            currCSS += '\t\t\tstroke-miterlimit: %d;\n' %4  # It's a default value, too. 
            if array: currCSS += '\t\t\tstroke-dasharray: %s;\n' %array
            currCSS += '\t\t\tfill: none\n\t\t}\n'
            self.sCSS += currCSS
        return url

    def _getArrowPos(self, position, distance):
        """ Fucking dummy for auto-placement for arrows """
        if position == CONST2D.ksDimArrowAuto:  # 0x2
            if distance > 10.0: # TODO: remove magic constant
                return CONST2D.ksDimArrowInside # 0x0
            else:
                return CONST2D.ksDimArrowOutside    # 0x1
        else:
            return position

## =========== SVG conventer ==================
class CConverter(_CFakeConverter):
    def __init__(self, document):
        self.iksDocument2D = KAPI7.IKompasDocument2D(document)
        self.defs = CElement(X_DEFS, id=ID_DEFS %self.getID())  # init <defs>

    def CreateText(self, text, x, y, angle=0.0, alloc=0):
        """ Create XML-tag '<text>'
            Args:
                [in]  str: text     - text
                [in]  float: x, y   - coordinates
                [in]  float: angle  - angle
                [out] CElement: tag - XML-tag
        """
        if not text: return None    # It don't makes a tag, if text is empty
        style = 'font-size:%spx' %self.params['KEGEL']
        if alloc:
            style += ';text-anchor:%s' %TXT_ALLOCATION[alloc]
        if self.params['MATH_MODE']:
            text = '$%s$' %text
        tag = CElement(X_TEXT, id=ID_TEXT %self.getID(), text=text)
        tag.addAttrib(('x', '%.4f' %x))
        tag.addAttrib(('y', '%.4f' %y))
        tag.addAttrib(('style', style))
        if angle:
            tag.addAttrib(('transform', 'rotate(%.4f %.4f %.4f)' %(angle, x, y)))
        return tag

    def addMarker(self, type=0, dir=0, pos=0):
        """ Create SVG-marker
            Args:
                [in] int: type  - Kompas-type of marker
                [in] int: dir   - arrows placement (in/out/auto)
                [in] int: pos   - position an base object (start/end)
                [out] str: 
        """
        if not type: return ''  # without marker
        key = type + (dir << 8) + (pos << 16)
        try:
            url = self.markersDB[key]   # Attempt get marker from DB, if exist
        except KeyError:    # generate marker & add to DB
            if type == CONST2D.ksPoint: # make point marker
                url = ID_MPOINT
                subtag = CElement(X_CIRCL, id=ID_GENOBJ %self.getID())
                subtag.setAttribs( (
                    ('cx', '0'), ('cy', '0'),
                    ('r', '%.4f'%(self.params['D_POINT']/2))#, ('style', self.params['ARROWSTYLE'])
                    ) )

            elif type == CONST2D.ksArrow:   # make arrow
                len = 4.0   # [mm] length of arrow
                pole = 0.7  # [mm] offset pole
                wid = 0.7   # [mm] width/2 of arrow
                coord = (len, -wid, len-pole, 0.0, len, wid)
                sPathA = 'M 0,0 %.2f,%.2f Q %.2f,%.2f %.2f,%.2f Z' %coord
                sPathT = 'M 0,0 %.2f,%.2f' %(len + self.params['OVERHANG'], 0.0)
                
                subtag = CElement(X_GROUP, id=ID_GENOBJ %self.getID())
                tag_arrow = CElement(X_PATH, id=ID_DARROW %self.getID())
                #tag_arrow.setAttribs( (('d', sPathA), ('style', self.params['ARROWSTYLE'])) )
                tag_arrow.addAttrib(('d', sPathA))
                subtag.append(tag_arrow)
                
                tag_tail = CElement(X_PATH, id=ID_GENOBJ %self.getID())
                tag_tail.setAttribs( (('d', sPathT),('class', self.defstyle)) )
                subtag.append(tag_tail)

                if dir == 0:    # inner arrow pos
                    if pos == DL_START:    # start in
                        url = ID_MARROWST
                    else:
                        url = ID_MARROWEND
                        subtag.addAttrib( ('transform', 'scale(-1)') )
                else:   # out arrow pos
                    if pos == DL_START:    # start out
                        url = ID_MARROWSTOUT
                        subtag.addAttrib( ('transform', 'scale(-1)') )
                    else:
                        url = ID_MARROWENDOUT

            elif type in (CONST2D.ksNotch, CONST2D.ksLeftNotch):   # make notch
                subtag = CElement(X_GROUP, id=ID_GENOBJ %self.getID())
                sPath = 'M %.2f,%.2f L %.2f,%.2f'
                tag_notch = CElement(X_PATH, id=ID_GENOBJ %self.getID())
                tag_notch.setAttribs( (
                    ('d', sPath %(-1.06, 1.06, 1.06, -1.06)),
                    ('class', self.defstyle)) )
                subtag.append(tag_notch)

                tag_tail = CElement(X_PATH, id=ID_GENOBJ %self.getID())
                tag_tail.setAttribs( (
                    ('d', sPath %(-self.params['OVERHANG'], 0.0, self.params['OVERHANG'], 0.0)),
                    ('class', self.defstyle)) )
                subtag.append(tag_tail)

                if type == CONST2D.ksNotch:
                    url = ID_MNOTCH
                else:
                    url = ID_MLNOTCH
                    subtag.addAttrib( ('transform', 'scale(-1)') )
                
            else:   # without marker
                return ''
            m_tag = CElement(X_MARK, id=url)
            m_tag.setAttribs( (
                ('orient', 'auto'), ('style', 'overflow:visible'),
                ('refX', '0'), ('refY', '0'),
                ('markerUnits', 'userSpaceOnUse')) )
            m_tag.append(subtag)
            self.defs.append(m_tag)
            self.markersDB[key] = url # add markers url to DB current svg
        if pos == DL_START:
            return 'marker-start:url(#%s)' %url
        else: #== DL_END
            return 'marker-end:url(#%s)' %url

    def fitCanvas(self, IView):
        """ It fits canvas by gabarite of active view """
        if not self.params['FIT_CANVAS']:
            return
        ViewParams = iKompasObj.GetParamStruct(CONST2D.ko_RectParam)
        if iDocument2D.ksGetObjGabaritRect(IView.Reference, ViewParams):
            pt0 = ViewParams.GetpBot()  # bottom left point
            pt1 = ViewParams.GetpTop()  # top right point
            x0 = ceil(pt0.x) - 1.0 - self.params['CANVAS_GAP']
            y0 = ceil(pt0.y) - 1.0 - self.params['CANVAS_GAP']
            x1 = ceil(pt1.x) + self.params['CANVAS_GAP']
            y1 = ceil(pt1.y) + self.params['CANVAS_GAP']
            self.params['WIDTH'] = int(x1 - x0)
            self.params['HEIGHT'] = int(y1 - y0)
            self._X_start = x0
            self._Y_start = y0
        return

    def makePattern(self, params, ref):
        """ It makes patterns for hatches """
        color = param.Color
        style = param.Style
        angle = param.HathcAngle
        step = param.Step
        a, s = int(angle*100), int(step*100)
        url = 'HATCH%02dA%05dS%04dC%s' %(style, a, s, color)
        if self.patternDB.get(url, False): return url   # if pattern already exists, then return its url
        height = width = step*2.0
        scolor = ColorLong2RGB(color, bw=self.params['BLACK_WHITE']) # color as RGB
        tag_g = CElement(X_GROUP)   # make element wich contains description of pattern
        # pattern of hatch for a nonmetal or art stone
        if style in (CONST2D.ksHatchNonMetal, CONST2D.ksHatchArtificialStone):  # 0x1 or 0xA
            x1, x2 = y1, y2 = step/2.0, step*1.5
            line1h = CElement(X_PATH)
            line1h.addAttrib(('d', 'M 0,%.3f H %s' %(y1, width)))
            line1h.addAttrib(('class', self.defstyle))
            line2h = CElement(X_PATH)
            line2h.addAttrib(('d', 'M 0,%.3f H %s' %(y2, width)))
            line1h.addAttrib(('class', self.defstyle))
            line1v = CElement(X_PATH)
            line1v.addAttrib(('d', 'M %.3f,0 V %s' %(x1, height)))
            line1v.addAttrib(('class', self.defstyle))
            line2v = CElement(X_PATH)
            line2v.addAttrib(('d', 'M %.3f,0 V %s' %(x2, height)))
            line1v.addAttrib(('class', self.defstyle))
            if color:
                line1h.addAttrib(('style', 'stroke:#%s' %scolor))
                line1h.addAttrib(('style', 'stroke:#%s' %scolor))
                line1v.addAttrib(('style', 'stroke:#%s' %scolor))
                line1v.addAttrib(('style', 'stroke:#%s' %scolor))
            tag_g.append(line1h)
            tag_g.append(line2h)
            tag_g.append(line1v)
            tag_g.append(line2v)
            if style == CONST2D.ksHatchArtificialStone: angle = 45.0    # reset angle for stone
        # pattern of hatch for transversal section of timber
        elif style == CONST2D.ksHatchTimber:    # 0x2
            ViewParams = iKompasObj.GetParamStruct(CONST2D.ko_RectParam)
            if iDocument2D.ksGetObjGabaritRect(ref, ViewParams):
                pt0, pt1 = ViewParams.GetpBot(), ViewParams.GetpTop()
                x0, y0 = floor(pt0.x), floor(pt0.y)
                x1, y1 = ceil(pt1.x), ceil(pt1.y)
                width, height = abs(x1 - x0), abs(y1 - y0)
                if angle == 0.0:    # case 1:
                    d = sqrt(width**2 + height**2)
                    cx = cy = 0
                elif angle < 0.0:   # case 2
                    _, cx, cy = Rotate(0.0, 0.0, 0.0, height/2, 2*angle)
                    d1 = Dist(cx, cy, width, height)
                    d2 = Dist(cx, cy, width, 0.0)
                    d = max(d1, d2)
                elif angle > 0.0:   # case 3
                    _, cx, cy = Rotate(0.0, 0,0, width/2, 0.0, 2*angle)
                    d1 = Dist(cx, cy, width, height)
                    d2 = Dist(cx, cy, 0.0, height)
                    d = max(d1, d2)
            else:   # TODO: add handling of error
                print('Error of getting of gabarites!')
            for r in range(step, int(d), step): # Maybe 'int(d) + 1'?
                tag_arc = CElement(X_CIRCL)
                tag_arc.addAttrib(('cx', cx))
                tag_arc.addAttrib(('cy', cy))
                tag_arc.addAttrib(('r', r))
                tag_arc.addAttrib(('class', self.defstyle))
                if color:   # if color is not black (0), then add its to the tag
                    tag_arc.addAttrib(('style', 'stroke:#%s' %scolor))
                tag_g.append(tag_arc)
        # pattern for natural stone
        elif style == CONST2D.ksHatchNaturalStone:  # 0x3
            pass
        # pattern for ceramics
        elif style == CONST2D.ksHatchCeramics:  # 0x4
            y1, y2 = 0.0, step/3.0
            line1 = CElement(X_PATH)
            line1.addAttrib(('d', 'M 0,%.2f H %s' %(y1, width)))
            line1.addAttrib(('class', self.defstyle))
            line2.addAttrib(('d', 'M 0,5.2f H %s' %(y2, width)))
            line2.addAttrib(('class', self.defstyle))
            if color:
                line1.addAttrib(('style', 'stroke:#%s' %scolor))
                line2.addAttrib(('style', 'stroke:#%s' %scolor))
            tag_g.append(line1)
            tag_g.append(line2)
        # pattern for concrete
        elif style == CONST2D.ksHatchConcrete:  # 0x5
            pass
        elif style == CONST2D.ksHatchGlass:     # 0x6
            pass
        elif style == CONST2D.ksHatchLiquid:    # 0x7
            pass
        elif style == CONST2D.ksHatchNaturallyGround:   # 0x8
            pass
        elif style == CONST2D.ksHatchSpreadGround:      # 0x9
            pass
        elif style == CONST2D.ksHatchReinforcedConcrete:# 0xB
            pass
        elif style == CONST2D.ksHatchTenseReinforcedConcrete:   # 0xC
            pass
        elif style == CONST2D.ksHatchLongitudalTimber:  # 0xD
            # Draw 1 horizontal segment
            height = width = step
            del tag_g   # here, for one object, a group isn't requiring
            tag_g = CElement(X_PATH)
            tag_g.addAttrib(('d', 'M 0,%.3f H %s' %(step/2.0, width)))
            tag_g.addAttrib(('class', self.defstyle))
            if color: tag_g.addAttrib(('style', 'stroke:%s' %scolor))
        #elif style == CONST2D.ksHatchSand:  # 0xE
        #    pass   # It's while not implemented
        else: #CONST2D.ksHatchMetal: 0x0 - This is default style of hatch
            y1, y2 = step/2.0, step*1.5
            line1 = CElement(X_PATH)
            line1.addAttrib(('d', 'M 0,%.3f H %s' %(y1, width)))
            line1.addAttrib(('class', self.defstyle))
            line2 = CElement(X_PATH)
            line2.addAttrib(('d', 'M 0,%.3f H %s' %(y2, width)))
            line2.addAttrib(('class', self.defstyle))
            if color:
                line1.addAttrib(('style', 'stroke:%s' %scolor))
                line2.addAttrib(('style', 'stroke:%s' %scolor))
            tag_g.append(line1)
            tag_g.append(line2)
        # === === ===
        tag_pattern = CElement('pattern', id=url)
        tag_pattern.addAttrib(('patternUnits', 'userSpaceOnUse'))
        if angle != 0.0:
            tag_pattern.addAttrib(('patternTransform', 'rotate(-%.3f)' %angle))
        tag_pattern.addAttrib(('width', width))
        tag_pattern.addAttrib(('height', height))
        tag_pattern.append(tag_g)
        self.defs.append(tag_pattern)
        self.patternDB[url] = True
        return url

    def Convert(self, params=None):
        """
            Convert Kompas-fragment to SVG-file
            Args:
                [in] dict: params   - parameters of converting
                [out] str: svg      - text of svg-file
        """
        if params:
            self.params.update(params)
        self._setstyle()    # init styles
        self.defstyle = self.MakeStyle(0x2, bw=True) # small optimization
        iView = self.iksDocument2D.ViewsAndLayersManager.Views.ActiveView
        self.fitCanvas(iView)

        g = CElement(X_GROUP, id='layer1')  # top-level tag '<g>'
        g.addAttrib(('inkscape:label', 'Layer 1'))   # ...
        g.addAttrib(('inkscape:groupmode', 'layer')) # ... sugar for Inkscape
        for element in self.ObjectsProcessing(iView):
            g.append(element)
        root = CElement('svg')   # root tag '<svg>'
        root.setAttribs((
            # --- DON'T EDIT THIS BLOCK! ---
            ('version', '1.1'), ('baseProfile', 'full'),
            ('xmlns', 'http://www.w3.org/2000/svg'),
            ('xmlns:svg', 'http://www.w3.org/2000/svg'),
            ('xmlns:xlink', 'http://www.w3.org/1999/xlink'),
            ('xmlns:inkscape', 'http://www.inkscape.org/namespaces/inkscape'),
            # ^^^ DON'T EDIT THIS BLOCK! ^^^
            ('width', '%dmm'%self.params['WIDTH']),
            ('height', '%dmm'%self.params['HEIGHT']),
            ('viewBox', '0 0 %d %d'%(self.params['WIDTH'], self.params['HEIGHT']))
            ))
        e_style = CElement('style', text=self.sCSS) # tag <defs>/<style>
        e_style.addAttrib(('type', 'text/css'))
        self.defs.append(e_style)
        root.append(self.defs)
        root.append(g)
        sSVG = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<!-- Created with KOMPAS-3D & ks2svg macros -->\n\n'
        return sSVG + root.printf()

    def ObjectsProcessing(self, IView, transform=None):
        """ Enumerating and processing all objects from fragment
            Args:
                [in] IView: IView - current View
                [in] dict: transform - params of coordinates transformations
                [out] list: Result - list of converted objects
        """
        Result = []
        iDrawingContainer = KAPI7.IDrawingContainer(IView)
        iSymbols2DContainer = KAPI7.ISymbols2DContainer(IView)
        # Последний в списке вызовов класс будет на высшем уровне SVG-документа
        # TODO: добавить возможность выбора порядка преобразования
        # Colouring
        if self.params['Colourings']:
            iColourings = iDrawingContainer.Colourings
            for i in range(iColourings.Count):
                iObj = iColourings.Colouring(i)
                Result.append(self.ConvertColouring(iObj, transform))
        # Hatch
        if self.params['Hatches']:
            iHatches = iDrawingContainer.Hatches
            for i in range(iHatches.Count):
                iObj = iHatches.Hatch(i)
                Result.append(self.ConvertHatch(iObj, transform))
        # Dimensions
        if self.params['Dimensions']:
            # simple line dimension
            iLineDimensions = iSymbols2DContainer.LineDimensions
            for i in range(iLineDimensions.Count):
                iObj = iLineDimensions.LineDimension(i)
                Result.append(self.ConvertLineDimension(iObj, transform))
            # simple angle dimension
            iAngleDimensions = iSymbols2DContainer.AngleDimensions
            for i in range(iAngleDimensions.Count):
                iObj = iAngleDimensions.AngleDimension(i)
                Result.append(self.ConvertAngleDimension(iObj, transform))
            # diametral dimension
            iDiametralDimensions = iSymbols2DContainer.DiametralDimensions
            for i in range(iDiametralDimensions.Count):
                iObj = iDiametralDimensions.DiametralDimension(i)
                Result.append(self.ConvertDiametralDimension(iObj, transform))
            # simple radial dimension
            iRadialDimensions = iSymbols2DContainer.RadialDimensions
            for i in range(iRadialDimensions.Count):
                iObj = iRadialDimensions.RadialDimension(i)
                Result.append(self.ConvertRadialDimension(iObj, transform))
            # break line dimension
            iBreakLineDimensions = iSymbols2DContainer.BreakLineDimensions
            for i in range(iBreakLineDimensions.Count):
                iObj = iBreakLineDimensions.BreakLineDimension(i)
                Result.append(self.ConvertBreakLineDimension(iObj, transform))
            # break angle dimension (it's while not supported)
            #iBreakAngleDimensions =
        # Raster image
        if self.params['Rasters']:
            iRasters = iDrawingContainer.Rasters
            for i in range(iRasters.Count):
                iObj = iRasters.Raster(i)
                Result.append(self.ConvertRaster(iObj, transform))
        # Line Segment
        if self.params['LineSegments']:
            iLineSegments = iDrawingContainer.LineSegments
            for i in range(iLineSegments.Count):
                iObj = iLineSegments.LineSegment(i)
                Result.append(self.ConvertLineSeg(iObj, transform))
        # Circle
        if self.params['Circles']:
            iCircles = iDrawingContainer.Circles
            for i in range(iCircles.Count):
                iObj = iCircles.Circle(i)
                Result.append(self.ConvertCircle(iObj, transform))
        # Arc
        if self.params['Arcs']:
            iArcs = iDrawingContainer.Arcs
            for i in range(iArcs.Count):
                iObj = iArcs.Arc(i)
                Result.append(self.ConvertArc(iObj, transform))
        # Bezier, NURBS and Nurbs by points
        if self.params['Curves']:
            iBeziers = iDrawingContainer.Beziers
            for i in range(iBeziers.Count):
                iObj = iBeziers.Bezier(i)
                Result.append(self.ConvertBezier(iObj, transform))
            iNurbses = iDrawingContainer.Nurbses
            for i in range(iNurbses.Count):
                iObj = iNurbses.Nurbs(i)
                Result.append(self.ConvertNurbs(iObj, transform))
        # Drawing Contour
        if self.params['DrawContours']:
            iDrawingContours = iDrawingContainer.DrawingContours
            for i in range(iDrawingContours.Count):
                iObj = iDrawingContours.DrawingContour(i)
                Result.append(self.ConvertDrawingContour(iObj, transform))
        # 2D-Polyline
        if self.params['PolyLines2D']:
            iPolylines2D = iDrawingContainer.PolyLines2D
            for i in range(iPolylines2D.Count):
                iObj = iPolylines2D.PolyLine2D(i)
                Result.append(self.ConvertPolyline2D(iObj, transform))
        # Ellipse
        if self.params['Ellipses']:
            iEllipses = iDrawingContainer.Ellipses
            for i in range(iEllipses.Count):
                iObj = iEllipses.Ellipse(i)
                Result.append(self.ConvertEllipse(iObj, transform))
        # Ellipses arc
        if self.params['EllipsesArcs']:
            iEllipseArcs = iDrawingContainer.EllipseArcs
            for i in range(iEllipseArcs.Count):
                iObj = iEllipseArcs.EllipsesArc(i)
                Result.append(self.ConvertEllipseArc(iObj, transform))
        # Polygon
        if self.params['Polygons']:
            iRectangles = iDrawingContainer.Rectangles
            for i in range(iRectangles.Count):
                iObj = iRectangles.Rectangle(i)
                Result.append(self.ConvertPolygon(iObj, transform))
            iRegularPolygons = iDrawingContainer.RegularPolygons
            for i in range(iRegularPolygons.Count):
                iObj = iRegularPolygons.RegularPolygon(i)
                Result.append(self.ConvertPolygon(iObj, transform))
        # Centre Mark
        if self.params['CentreMarks']:
            iCentreMarks = iSymbols2DContainer.CentreMarkers
            for i in range(iCentreMarks.Count):
                iObj = iCentreMarks.CentreMarker(i)
                Result.append(self.ConvertCentreMarker(iObj, transform))
        # Axis Line
        if self.params['AxisLines']:
            iAxisLines = iSymbols2DContainer.AxisLines
            for i in range(iAxisLines.Count):
                iObj = iAxisLines.AxisLine(i)
                Result.append(self.ConvertAxisLine(iObj, transform))
        # Base Leader
        if self.params['Leaders']:
            iLeaders = iSymbols2DContainer.Leaders
            for i in range(iLeaders.Count):
                iObj = iLeaders.Leader(i)
                Result.append(self.ConvertLeader(iObj, transform))
        # MACRO OBJECT
        if self.params['MacroObjects']:
            iMacroObjects = iDrawingContainer.MacroObjects
            for i in range(iMacroObjects.Count):
                iObj = iMacroObjects.MacroObject(i)
                Result.append(self.ConvertMacro(iObj, transform))
        # Point
        if self.params['Points']:
            iPoints = iDrawingContainer.Points
            for i in range(iPoints.Count):
                iObj = iPoints.Point(i)
                Result.append(self.ConvertPoint(iObj, transform))
        # Text
        if self.params['DrawTexts']:
            iDrawingTexts = iDrawingContainer.DrawingTexts
            for i in range(iDrawingTexts.Count):
                iObj = iDrawingTexts.DrawingText(i)
                Result.append(self.ConvertText(iObj, transform))
        return Result

    def ConvertMacro(self, IMacroObject, transform):
        """ Generator of XML-tag '<group>' for Macro Object """
        _, x0, y0, ang, _ = IMacroObject.GetPlacement(0.0, 0.0, 0.0, False)
        if transform:
            _, x, y = Rotate(x0, y0, 0.0, 0.0, ang)
            transform['x'] += x
            transform['y'] += y
            transform['ang'] += ang
        else:
            transform = {'x': x0, 'y': y0, 'ang': ang}
        elements = self.ObjectsProcessing(IMacroObject, transform)
        if not elements: return None
        macro = CElement(X_GROUP, id=ID_MACRO %self.getID())
        for element in elements:
            macro.append(element)
        return macro

    def ConvertArc(self, IArc, transform):
        """ Generator of XML-tag '<arc>' for Arcs """
        tag = CElement(X_PATH, id=ID_ARCUS %self.getID())
        tag.addAttrib(('d', self.MakeArcus(IArc, transform)))
        tag.addAttrib(('class', self.MakeStyle(IArc.Style)))
        return tag

    def ConvertBezier(self, IBezier, transform):
        """ Generator of XML-tag '<path>' for Beziers """
        tag = CElement(X_PATH, id=ID_BEZIER %self.getID())
        tag.addAttrib(('d', self.MakeBezier(IBezier, transform)))
        tag.addAttrib(('class', self.MakeStyle(IBezier.Style)))
        return tag

    def ConvertCircle(self, ICircle, transform):
        """ Generator of XML-tag '<circle>' """
        Xc, Yc = self.GetXY(ICircle.Xc, ICircle.Yc, transform)
        tag = CElement(X_CIRCL, id=ID_CIRCLE %self.getID())
        tag.setAttribs((
            ('cx', '%.4f' %Xc), ('cy', '%.4f' %Yc),
            ('r', '%.4f' %ICircle.Radius),
            ('class', self.MakeStyle(ICircle.Style))
            ))
        return tag

    def ConvertDrawingContour(self, IDrawingContour, transform):
        """ Generator of XML-tag '<g>' for drawing contour """
        iContour = KAPI7.IContour(IDrawingContour)
        tag = CElement(X_PATH, id=ID_CONTOUR %self.getID())
        tag.addAttrib(('d', self.ReadContour(iContour, True, transform)))
        tag.addAttrib(('class', self.MakeStyle(IDrawingContour.Style)))
        return tag

    def ConvertEllipse(self, IEllipse, transform):
        """ Generator of XML-tag '<ellipse>' """
        Xc, Yc = self.GetXY(IEllipse.Xc, IEllipse.Yc, transform)
        Angle = self.GetAngle(IEllipse.Angle, transform)    # [deg]
        tag = CElement('ellipse', id=ID_ELLIPSE %self.getID())
        tag.setAttribs((
            ('cx', '%.4f' %Xc), ('cy', '%.4f' %Yc),
            ('rx', '%.4f' %IEllipse.SemiAxisA),
            ('ry', '%.4f' %IEllipse.SemiAxisB),
            ('class', self.MakeStyle(IEllipse.Style))
            ))
        if Angle != 0.0:
            tag.addAttrib(('transform', 'rotate(%.4f %.4f %.4f)' %(Angle, Xc, Yc)))
        return tag

    def ConvertEllipseArc(self, IEllipseArc, transform):
        """ Generator of XML-tag '<path>' for Ellipse Arc """
        tag = CElement(X_PATH, id=ID_ELLIPSEARC %self.getID())
        tag.addAttrib(('d', self.MakeEllipseArc(IEllipseArc, transform)))
        tag.addAttrib(('class', self.MakeStyle(IEllipseArc.Style)))
        return tag

    def ConvertHatch(self, IHatch, transform):
        """ The generator of a XML-tag '<path>' for Hatch """
        # Get a boundary of hatch
        iBounds = KAPI7.IBoundariesObject(IHatch).Boundaries
        Forest = MakeForest([CNode(KAPI7.IContour(item)) for item in iBounds])
        sPath = ''
        for Tree in Forest:
            direct = True
            for lvl in range(Tree.getMaxLevel() + 1):
                for contour in Tree.getNodesByLevel(lvl):
                    sPath += self.ReadContour(contour.data, direct, transform)
                direct = not direct
        # Get pattern
        url = self.makePattern(KAPI7.IHatchParam(IHatch), IHatch.Reference)
        # Finally, make XML-tag
        tag = CElement(X_PATH, id=ID_HATCH %self.getID())
        tag.addAttrib(('d', sPath))
        tag.addAttrib(('style', 'fill:url(#%s);stroke:none' %url))
        return tag

    def ConvertColouring(self, IColouring, transform):
        """ Generator of XML-tag '<path>' for Colourings boundary """
        # Prepare data: Create list of nodes from contours of colour. boundaries
        iBounds = KAPI7.IBoundariesObject(IColouring).Boundaries
        Forest = [CNode(KAPI7.IContour(item)) for item in iBounds]
        Forest = MakeForest(Forest) # make true Forest
        sPath = ''
        for Tree in Forest: # enum. Tree in Forest
            direct = True   # flag of path-tracing direction
            for lvl in range(Tree.getMaxLevel() + 1):  # enum. contours from each lvl
                for contour in Tree.getNodesByLevel(lvl):   # enum contours from current lvl
                    sPath += self.ReadContour(contour.data, direct, transform)
                direct = not direct # reverce path-tracing, if go to next level

        #cType = IColouring.ColouringType   # while not used
        Color1 = ColorLong2RGB(IColouring.Color1)  # color as long --> color as hex
        opac1 = 1.0 - 0.01*IColouring.Transparency1 # transparency (%) --> opacity (0...1 units)
        # make XML-tag
        tag = CElement('path', id=ID_COLOURING %self.getID())
        tag.addAttrib(('d', sPath))
        tag.addAttrib(('style', 'fill:#%s;fill-opacity:%.4f;stroke:none' %(Color1, opac1)))
        return tag

    def ConvertLineSeg(self, ILineSegment, transform):
        """ Generator of XML-tag '<path>' for line segment """
        X1, Y1 = self.GetXY(ILineSegment.X1, ILineSegment.Y1, transform)
        X2, Y2 = self.GetXY(ILineSegment.X2, ILineSegment.Y2, transform)
        tag = CElement('path', id=ID_LINESEG %self.getID())
        tag.addAttrib(('d', 'M %.4f,%.4f %.4f,%.4f' %(X1, Y1, X2, Y2)))
        tag.addAttrib(('class', self.MakeStyle(ILineSegment.Style)))
        return tag

    def ConvertNurbs(self, INurbs, transform):
        """ Generator of XML-tag '<path>' for NURBS """
        tag = CElement('path', id=ID_NURBS %self.getID())
        tag.addAttrib(('d', self.ApproxNurbs(INurbs, transform)))
        tag.addAttrib(('class', self.MakeStyle(INurbs.Style)))
        return tag

    def ConvertPoint(self, IPoint, transform, diameter=None):
        """ Generator of XML-tag '<circle>' for point """
        # TODO: use patterns
        X, Y = self.GetXY(IPoint.X, IPoint.Y, transform)
        if not diameter:
            diameter = self.params['D_POINT']
        tag = CElement('circle', id=ID_POINT %self.getID())
        tag.setAttribs((
            ('cx', '%.4f' %X), ('cy', '%.4f' %Y),
            ('r', '%.4f' %(diameter/2.0))#, ('style', self.params['ARROWSTYLE'])
            ))
        return tag

    def ConvertPolyline2D(self, IPolyLine2D, transform):
        """ Generate of XML-tag '<path>' for Polyline """
        sPath = 'M'
        for i in range(IPolyLine2D.PointsCount):
            _, Xp, Yp = IPolyLine2D.GetPoint(i) 
            sPath += ' %.4f,%.4f' %self.GetXY(Xp, Yp, transform)
        tag = CElement('path', id=ID_POLYLINE %self.getID())
        tag.addAttrib(('d', sPath + ' Z'*IPolyLine2D.Closed))
        tag.addAttrib(('class', self.MakeStyle(IPolyLine2D.Style)))
        return tag

    def ConvertPolygon(self, IPolygon, transform):
        """ Generator of XML-tag '<path>' for any polygons """
        # TODO: refact this VERY slow fucking method
        pTmpGroup = iDocument2D.ksDecomposeObj(IPolygon.Reference, 1, self.params['TOLERANCE'], 0)
        iTmpGroup = iKompasObj.TransferReference(pTmpGroup, 0)
        sPath = 'M'
        for iObj in iTmpGroup.Objects():    # type of iObj - only LineSegment
            sPath += ' %.4f, %.4f' %self.GetXY(iObj.X1, iObj.Y1, transform)
        iTmpGroup.Delete()
        tag = CElement('path', id=ID_POLYGON %self.getID())
        tag.addAttrib(('d', sPath + ' Z'))    # always closed
        tag.addAttrib(('class', self.MakeStyle(IPolygon.Style)))
        return tag

    def ConvertText(self, IDrawingText, transform):
        """ Generator of XML-tag '<text>' """
        X, Y = self.GetXY(IDrawingText.X, IDrawingText.Y, transform)
        Angle = self.GetAngle(IDrawingText.Angle, transform)
        alloc = IDrawingText.Allocation
        return self.CreateText(KAPI7.IText(IDrawingText).Str, X, Y, Angle, alloc)

    def ConvertAxisLine(self, IAxisLine, transform):
        """ Generator of XML-tag for axis line """
        # TODO: this method is very slow (average time = 0.1 s), recover old version?
        pTempGroup = iDocument2D.ksDecomposeObj(IAxisLine.Reference, 1, 0.5, 0)
        iTempGroup = iKompasObj.TransferReference(pTempGroup, 0)
        tag = self.ConvertLineSeg(iTempGroup.Objects(), transform) # group contain only 1 object 'ILineSegment'
        iTempGroup.Delete()
        return tag

    def ConvertCentreMarker(self, ICentreMarker, transform):
        """ Generator of XML-tag '<g>' for centre markers """
        # this method is VERY slow (aver. time = 0.23 s), recover old version?
        pTmpGroup = iDocument2D.ksDecomposeObj(ICentreMarker.Reference, 1, 0.5, 0)
        iTmpGroup = iKompasObj.TransferReference(pTmpGroup, 0)
        tag = CElement('g', id=ID_CMARK %self.getID())
        for obj in iTmpGroup.Objects(): # get all objects from group as SAFEARRAY | VT_DISPATCH
            if obj.Length > self.params['OVERHANG']:   # skip 'tails'
                tag.append(self.ConvertLineSeg(obj, transform))
        iTmpGroup.Delete()
        return tag

    def ConvertLineDimension(self, ILineDimension, transform):
        """ Generator of XML-tag for objects group """
        X1, Y1 = self.GetXY(ILineDimension.X1, ILineDimension.Y1, transform)
        X2, Y2 = self.GetXY(ILineDimension.X2, ILineDimension.Y2, transform)
        X3, Y3 = self.GetXY(ILineDimension.X3, ILineDimension.Y3, transform)
        Angle, Xtext, Ytext = self.GetTextPlacement(ILineDimension.Reference, transform) # Get text coords
        params = KAPI7.IDimensionParams(ILineDimension)

        ## Calculate coords of dims
        # Reset dimension to horizontal
        _, x, y = Rotate(X1, Y1, 0, 0, -Angle)
        P10 = {'X': x, 'Y': y}
        _, x, y = Rotate(X2, Y2, 0, 0, -Angle)
        P20 = {'X': x, 'Y': y}
        _, x, y = Rotate(X3, Y3, 0, 0, -Angle)
        P30 = {'X': x, 'Y': y}
        VLeader1 = {'X': 0.0, 'Y': P30['Y'] - P10['Y']}
        VLeader2 = {'X': 0.0, 'Y': P30['Y'] - P20['Y']}
        # Recovering rotate coordinates
        _, x, y = Rotate(VLeader1['X'], VLeader1['Y'], 0, 0, Angle)
        VLeader1 = {'X': x, 'Y': y}
        _, x, y = Rotate(VLeader2['X'], VLeader2['Y'], 0, 0, Angle)
        VLeader2 = {'X': x, 'Y': y}
        # Other coords
        x, y = ScaleCoordinates(VLeader1['X'], VLeader1['Y'], 0, 0, dm=self.params['OVERHANG'])
        VLeader1ex = {'X': x, 'Y': y}
        x, y = ScaleCoordinates(VLeader2['X'], VLeader2['Y'], 0, 0, dm=self.params['OVERHANG'])
        VLeader2ex = {'X': x, 'Y': y}

        ## Calculate coords  
        # coords of leaders lines
        PE1 = {'X': X1 + VLeader1ex['X'], 'Y': Y1 + VLeader1ex['Y']}
        PE2 = {'X': X2 + VLeader2ex['X'], 'Y': Y2 + VLeader2ex['Y']}
        # coords of dimension line
        PA1 = {'X': X1 + VLeader1['X'], 'Y': Y1 + VLeader1['Y']}
        PA2 = {'X': X2 + VLeader2['X'], 'Y': Y2 + VLeader2['Y']}

        # create tag
        tag = CElement('g', id=ID_LINEDIM %self.getID())
        # create 1st leader line
        if params.RemoteLine1:
            tag_l1 = CElement('path', id=ID_LINESEG %self.getID())
            tag_l1.setAttribs((
                ('d', 'M %.4f,%.4f %.4f,%.4f' %(X1, Y1, PE1['X'], PE1['Y'])),
                ('class', self.defstyle)
                ))
            tag.append(tag_l1)
        # create 2st leader line
        if params.RemoteLine2:
            tag_l2 = CElement('path', id=ID_LINESEG %self.getID())
            tag_l2.setAttribs((
                ('d', 'M %.4f,%.4f %.4f,%.4f' %(X2, Y2, PE2['X'], PE2['Y'])),
                ('class', self.defstyle)
                ))
            tag.append(tag_l2)
        # create dimensions line
        tag_ld = CElement('path', id=ID_LINESEG %self.getID())
        tag_ld.setAttribs((
            ('d', 'M %.4f,%.4f %.4f,%.4f' %(PA1['X'], PA1['Y'], PA2['X'], PA2['Y'])),
            ('class', self.defstyle)
            ))
        dir = self._getArrowPos(params.ArrowPos, KAPI7.IDimensionText(ILineDimension).NominalValue)
        mark_s = self.addMarker(params.ArrowType1, dir, pos=DL_START)
        mark_e = self.addMarker(params.ArrowType2, dir, pos=DL_END)
        if mark_s or mark_e:
            marks = [x for x in (mark_s, mark_e) if x]
            tag_ld.addAttrib(('style', ';'.join(marks)))
        tag.append(tag_ld)
        # create text
        tag.append(self.CreateText(GetDimText(ILineDimension), Xtext, Ytext, Angle))
        return tag

    def ConvertBreakLineDimension(self, IBreakLineDimension, transform):
        """ Generator of XML-tag for objects group 'break line dimension' """
        Xp1, Yp1 = self.GetXY(IBreakLineDimension.X1, IBreakLineDimension.Y1, transform)
        Xp2, Yp2 = self.GetXY(IBreakLineDimension.X2, IBreakLineDimension.Y2, transform)
        Xp3, Yp3 = self.GetXY(IBreakLineDimension.X3, IBreakLineDimension.Y3, transform)
        Angle, Xt, Yt = self.GetTextPlacement(IBreakLineDimension.Reference, transform)  # Get text coords
        params = KAPI7.IDimensionParams(IBreakLineDimension)

        # calc coordinates
        VLeader = {'X': Xp2 - Xp1, 'Y': Yp2 - Yp1}
        x, y = ScaleCoordinates(VLeader['X'], VLeader['Y'], 0, 0, dm=self.params['OVERHANG'])
        X2, Y2 = Xp1 + x, Yp1 + y

        # create tag
        tag = CElement('g', id=ID_BRLINEDIM %self.getID())
        # create leader
        if params.RemoteLine1:
            tag_l = CElement('path', id=ID_LINESEG %self.getID())
            tag_l.setAttribs((
                ('d', 'M %.4f,%.4f %.4f,%.4f' %(Xp1, Yp1, X2, Y2)),
                ('class', self.defstyle)
                ))
            tag.append(tag_l)
        # create dim line
        tag_ld = CElement('path', id=ID_LINESEG %self.getID())
        tag_ld.setAttribs((
            ('d', 'M %.4f,%.4f %.4f,%.4f' %(Xp2, Yp2, Xp3, Yp3)),
            ('class', self.defstyle)
            ))
        mark = self.addMarker(params.ArrowType1, params.ArrowPos, DL_START)
        if mstyle:
            tag_ld.addAttrib(('style', mark))
        tag.append(tag_ld)
        # create text
        tag.append(self.CreateText(GetDimText(IBreakLineDimension), Xt, Yt, Angle))
        return tag

    def ConvertDiametralDimension(self, IDiametralDimension, transform):
        """ Generator of XML-tag for objects group 'diametral dimension' """
        Xc, Yc = self.GetXY(IDiametralDimension.Xc, IDiametralDimension.Yc, transform)
        Angle = self.GetAngle(IDiametralDimension.Angle, transform)
        R = IDiametralDimension.Radius
        _, X1, Y1 = Rotate(Xc + R, Yc, Xc, Yc, Angle)
        Angle2 = Angle + 180.0  # angle 2st arrow
        Cutted = IDiametralDimension.DimensionType  # full (False) or cutted (True) dim line
        if not Cutted:
            _, X2, Y2 = Rotate(Xc + R, Yc, Xc, Yc, Angle2)
        else:
            _, X2, Y2 = Rotate(Xc + 0.1*R, Yc, Xc, Yc, Angle2)
        params = KAPI7.IDimensionParams(IDiametralDimension)
        
        # dim line
        tag = CElement('g', id=ID_DIAMETRALDIM %self.getID())
        tag_ld = CElement('path', id=ID_LINESEG %self.getID())
        tag_ld.setAttribs((
            ('d', 'M %.4f,%.4f %.4f,%.4f' %(X1, Y1, X2, Y2)),
            ('class', self.defstyle)
            ))
        dir = self._getArrowPos(params.ArrowPos, R*2)
        mark_s = self.addMarker(params.ArrowType1, dir, pos=DL_START)
        mark_e = self.addMarker(params.ArrowType2, dir, pos=DL_END)
        if mark_e or mark_s:
            marks = [x for x in (mark_s, mark_e) if x]
            tag_ld.addAttrib(('style', ';'.join(marks)))
        tag.append(tag_ld)
        # create text
        AngT, Xt, Yt = self.GetTextPlacement(IDiametralDimension.Reference, transform)
        tag.append(self.CreateText(GetDimText(IDiametralDimension), Xt, Yt, AngT))
        return tag

    def ConvertRadialDimension(self, IRadialDimension, transform):
        """ Generator of XML-tag for object group 'radial dimension' """
        Xc, Yc = self.GetXY(IRadialDimension.Xc, IRadialDimension.Yc, transform)
        Angle = self.GetAngle(IRadialDimension.Angle, transform)
        R = IRadialDimension.Radius
        params = KAPI7.IDimensionParams(IRadialDimension)
        # calculate other points
        _, XT1, YT1 = Rotate(Xc + R, Yc, Xc, Yc, Angle)
        if IRadialDimension.DimensionType:  # full (True) or cutted (False) dim line
            coords = (XT1, YT1, Xc, Yc)
        else:   # TODO!
            coords = (XT1, YT1, Xc, Yc)
            print('Cutted radius while not supported!')
        # create dim line
        tag = CElement('g', id=ID_RADIALDIM %self.getID())
        tag_ld = CElement('path', id=ID_LINESEG %self.getID())
        tag_ld.setAttribs((
            ('d', 'M %.4f,%.4f %.4f,%.4f' %coords),
            ('class', self.defstyle)
            ))
        # markers
        dir = self._getArrowPos(params.ArrowPos, R*2)
        marker = self.addMarker(params.ArrowType1, dir, pos=DL_START)
        if marker:
            tag_ld.addAttrib(('style', marker))
        tag.append(tag_ld)
        # create text
        AngT, Xt, Yt = self.GetTextPlacement(IRadialDimension.Reference, transform)
        tag.append(self.CreateText(GetDimText(IRadialDimension), Xt, Yt, AngT))
        return tag

    def ConvertAngleDimension(self, IAngleDimension, transform):
        """ Generator of XML-tag for object group 'angle dimension' """
        Xc, Yc = self.GetXY(IAngleDimension.Xc, IAngleDimension.Yc, transform)
        X1, Y1 = self.GetXY(IAngleDimension.X1, IAngleDimension.Y1, transform)
        X2, Y2 = self.GetXY(IAngleDimension.X2, IAngleDimension.Y2, transform)
        Angle1 = self.GetAngle(IAngleDimension.Angle1, transform)
        Angle2 = self.GetAngle(IAngleDimension.Angle2, transform)
        params = KAPI7.IDimensionParams(IAngleDimension)
        R = IAngleDimension.Radius
        # calculate coordinates of other points
        theta1, theta2 = radians(Angle1), radians(Angle2)
        P1a = {'X': Xc + R*cos(theta1), 'Y': Yc + R*sin(theta1)}
        P2a = {'X': Xc + R*cos(theta2), 'Y': Yc + R*sin(theta2)}
        V1 = {'X': P1a['X'] - X1, 'Y': P1a['Y'] - Y1}
        V2 = {'X': P2a['X'] - X2, 'Y': P2a['Y'] - Y2}
        x, y = ScaleCoordinates(V1['X'], V1['Y'], 0, 0, dm=self.params['OVERHANG'])
        P1e = {'X': X1 + x, 'Y': Y1 + y}
        x, y = ScaleCoordinates(V2['X'], V2['Y'], 0, 0, dm=self.params['OVERHANG'])
        P2e = {'X': X2 + x, 'Y': Y2 + y}
        # additional calc
        if Angle1 > Angle2: # if dir positive, swap
            Angle1, Angle2 = Angle2, Angle1 # values Angle1 and Angle2...
            P1a, P2a = P2a, P1a     # ...and arrows coordinates
        ## drawing
        tag = CElement('g', id=ID_ANGLEDIM %self.getID())
        # create 1st leader
        if params.RemoteLine1:
            tag_l1 = CElement(X_PATH, id=ID_LINESEG %self.getID())
            tag_l1.setAttribs((
                ('d', 'M %.4f,%.4f %.4f,%.4f' %(X1, Y1, P1e['X'], P1e['Y'])),
                ('class', self.defstyle)
                ))
            tag.append(tag_l1)
        # create 2st leader
        if params.RemoteLine2:
            tag_l2 = CElement(X_PATH, id=ID_LINESEG %self.getID())
            tag_l2.setAttribs((
                ('d', 'M %.4f,%.4f %.4f,%.4f' %(X2, Y2, P2e['X'], P2e['Y'])),
                ('class', self.defstyle)
                ))
            tag.append(tag_l2)
        # create arrows
        dir = self._getArrowPos(params.ArrowPos, abs(Angle2 - Angle1)*pi*R/180)
        mark_s = self.addMarker(params.ArrowType1, dir, pos=DL_START)
        mark_e = self.addMarker(params.ArrowType2, dir, pos=DL_END)
        # create arc 
        coords = (P1a['X'], P1a['Y'], R, R, P2a['X'], P2a['Y'])
        tag_arc = CElement(X_PATH, id=ID_ARCUS %self.getID())
        tag_arc.setAttribs((
            ('d', 'M %.4f,%.4f A %.4f,%.4f 0 0 1 %.4f,%.4f' %coords),
            ('class', self.defstyle)
            ))
        if mark_s or mark_e:
            marks = [x for x in (mark_s, mark_e) if x]
            tag_arc.addAttrib(('style', ';'.join(marks)))
        tag.append(tag_arc)
        AngT, Xt, Yt = self.GetTextPlacement(IAngleDimension.Reference, transform)
        tag.append(self.CreateText(GetDimText(IAngleDimension), Xt, Yt, AngT))
        return tag

    def ConvertLeader(self, IBaseLeader, transform):
        """ Generator of XML-tag for leader """
        iLeader = KAPI7.ILeader(IBaseLeader)    # leader
        iBranchs = KAPI7.IBranchs(IBaseLeader)  # branch
        X0, Y0 = self.GetXY(iBranchs.X0, iBranchs.Y0, transform)
        tag = CElement('g', id=ID_LEADER %self.getID())
        for i in range(iBranchs.BranchCount):
            tag_branch = CElement(X_PATH, id=ID_BRANCH %self.getID())
            # make path
            sPath = 'M %.4f,%.4f' %(X0, Y0)
            points = list(iBranchs.BranchPoints(i))
            if iBranchs.BranchPointsCount(i) > 1: # branch is polyline
                x1, y1 = self.GetXY(points[-4], points[-3], transform)
                x2, y2 = self.GetXY(points[-2], points[-1], transform)   # last 2 points
                while points:   # this block is very slow, rewrite!
                    sPath += ' %.4f,%.4f' %self.GetXY(points.pop(0), points.pop(0), transform)
            else:   # branch is line segment
                x1, y1 = X0, Y0
                x2, y2 = self.GetXY(points[0], points[1], transform)
                sPath += ' %.4f,%.4f' %(x2, y2)
            tag_branch.setAttribs((
                ('d', sPath),
                ('class', self.defstyle)
                ))
            marker = self.addMarker(IBaseLeader.ArrowType, pos=DL_END)
            if marker:
                tag_branch.addAttrib(('style', marker))
            tag.append(tag_branch)
            ang = degrees(asin((y2 - y1)/sqrt((x2 - x1)**2 + (y2 - y1)**2)))
            if x2 < x1:
                ang = 180.0 - ang
        tag.append(self.CreateText(iLeader.TextOnShelf.Str, X0, Y0 - 1.0, angle=0.0))
        return tag

    def ConvertRaster(self, IRaster, transform):
        """ Generator of XML-tag for raster """
        if IRaster.InsertionType:   # is it embedded?
            return None # embedded raster not supported
        fpath = IRaster.FileName
        try:
            mime = getMIME(fpath)
            with open(fpath, 'rb') as f:
                import base64
                encoded = base64.b64encode(f.read())
                # It's beautyful output for encoded string (it doesn't functions)
                #width = 72  # sic!
                #start, end = 0, width
                #segs = []
                #while True:
                #    seg = encoded[start:end]
                #    start = end + 1
                #    end = start + width
                #    if not seg: break
                #    segs.append(seg)
                #encoded = '\n'.join(segs)
                del base64
        except IOError: # TODO: как-то надо обработать ошибку
            return None
        if not mime: return None
        _, x, y, ang, _ = IRaster.GetPlacement(0.0, 0.0, 0.0, False)
        X, Y = self.GetXY(x, y, transform)
        Angle = self.GetAngle(ang, transform)

        tag = CElement(X_IMAGE, id=ID_IMAGE %self.getID())
        tag.setAttribs( (
            ('width', IRaster.SourceWidth),
            ('height', IRaster.SourceHeight),
            ('x', X), ('y', Y - IRaster.SourceHeight),  # FUCKING CG COORDINATES SYSTEM!!!111
            ('xlink:href', 'data:%s;base64,%s' %(mime, encoded))
            #('preserveAspectRatio', 'none')    # Maybe remove it?
            ) )
        if Angle != 0.0:
            tag.addAttrib(('transform', 'rotate(%.4f,%.4f,%,4f)' %(Angle, X, Y)))
        return tag
#  ======== end of class definition ===========

## ======== Document's interface ==============
class CDocument:
    path, name = '', ''
    valid = False
    def __init__(self, app):
        idoc = app.ActiveDocument
        if idoc:    # Does document exist?
            self.path = idoc.Path
            self.name = idoc.Name
            self.valid = (idoc.DocumentType == CONST2D.ksDocumentFragment)
            self._converter = CConverter(KAPI7.IKompasDocument2D(idoc))
        else: pass

    def Convert(self, params):
        if self.valid:
            return self._converter.Convert(params)
        return ''   # TODO: show errormsg or raise exception?
