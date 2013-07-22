##
# rsCleanShapes Command
# @author Juan Lara
# @date 2013-07-20
# @file rsCleanShapes.py

import win32com.client
from win32com.client import constants

Application = win32com.client.Dispatch('XSI.Application').Application


def XSILoadPlugin(in_reg):
    in_reg.Author = 'Juan Lara'
    in_reg.Name = 'rsCleanShapes'
    in_reg.Email = "info@rigstudio.com"
    in_reg.URL = "www.rigstudio.com"
    in_reg.Major = 1
    in_reg.Minor = 0

    in_reg.RegisterCommand('rsCleanShapes', 'rsCleanShapes')
    in_reg.RegisterMenu(constants.siMenuTbAnimateDeformShapeID, "rsCleanShapes_Menu", False, False)
    
    return True


def XSIUnloadPlugin(in_reg):
    return True


def rsCleanShapes_Menu_Init(ctxt):
    oMenu = ctxt.Source
    oMenu.AddCommandItem("rsCleanShapes", "rsCleanShapes")
    return True


def rsCleanShapes_Init(in_ctxt):
    oCmd = in_ctxt.Source
    oCmd.Description = ''
    oCmd.ReturnValue = True

    oArgs = oCmd.Arguments
    oArgs.AddWithHandler('in_c_geometry', 'Collection')
    
    return True


## Clean the shapes of the given geometries.
# @param in_c_geometry The geometries to clean.
# @return (boolean) If cleaned or not.
def rsCleanShapes_Execute(in_c_geometry):
    c_geometry = Application.SIFilter(in_c_geometry, constants.siGeometryFilter)
    if not c_geometry:
        Application.Logmessage('Select some geometries before executing.', constants.siError)
        return False
    
    s_precision = '%.3f'
        
    for o_geometry in c_geometry:
        Application.Logmessage('Cleaning shapes in "%s"' % o_geometry.Name, constants.siInfo)
        
        for o_cls in o_geometry.ActivePrimitive.Geometry.Clusters:
            c_shapes = win32com.client.Dispatch('XSI.Collection')
            if str(o_cls.Properties(0)).split('.')[-1] == 'ResultClusterKey':
                for o_shape in o_cls.LocalProperties:
                    if str(o_shape).split('.')[-1] == 'ResultClusterKey' or o_shape.type != 'clskey':
                        continue
                    
                    c_shapes.Add(o_shape)
                    
                    # Crete the weightMap.
                    
                    o_wmap = Application.CreateWeightMap('', o_geometry, 'wm_%s' % o_shape.name, '', False)[0]
                    o_clswmap = o_wmap.parent
                    l_wmap = list(o_wmap.Elements.Array[0])
                    l_elem = list(o_shape.Elements.Array)
                                            
                    for i_elem in range(len(l_elem[0])):
                        l_tmp = []
                        [l_tmp.append(abs(l_elem[i_tmp][i_elem])) for i_tmp in range(0, 3)]
                        l_tmp = s_precision % sum(l_tmp)
                        if l_tmp != s_precision % 0:
                            l_wmap[i_elem] = 1
                            
                    o_wmap.Elements.Array = l_wmap
                    
                    Application.ApplyOp('ClsKeyWeightMapOp', '%s;%s' % (o_shape.FullName.replace('currentprops.ClsProp.', ''), o_wmap.FullName), 3, 'siPersistentOperation', '', 0)
            
            if c_shapes.Count != 0:
                Application.FreezeObj(c_shapes)
                Application.DeleteObj(o_clswmap)
                        
    return True
