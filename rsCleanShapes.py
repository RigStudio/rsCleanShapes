##
# rsCleanShapes Command
# @author Juan Lara
# @date 2013-07-20
# @file rsCleanShapes.py

import win32com.client
from win32com.client import constants


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
            if str(o_cls.Properties(0)).split('.')[-1] == 'ResultClusterKey':
                for o_shape in o_cls.LocalProperties:
                    if str(o_shape).split('.')[-1] == 'ResultClusterKey' or o_shape.type != 'clskey':
                        continue
                    
                    c_wmap = win32com.client.Dispatch('XSI.Collection')
                    for o_wmap in o_shape.NestedObjects:
                        if o_wmap.type == 'ClusterKeyWeightMap':
                            Application.DeactivateAbove(o_wmap.FullName, True)
                            c_wmap.Add(o_wmap)
        
                    l_elem = list(o_shape.Elements.Array)
                    for i_tmp in range(len(l_elem)):
                        l_elem[i_tmp] = list(l_elem[i_tmp])

                    for i_elem in range(len(l_elem[0])):
                        l_tmp = []
                        [l_tmp.append(abs(l_elem[i_tmp][i_elem])) for i_tmp in range(0, 3)]
                        l_tmp = s_precision % sum(l_tmp)
                        if l_tmp == s_precision % 0:
                            for i_tmp in range(0, 3):
                                l_elem[i_tmp][i_elem] = 0.0
                        
                    o_shape.Elements.Array = l_elem
                        
                    for o_wmap in o_shape.NestedObjects:
                        if o_wmap.type == 'ClsSetValuesOp':
                            Application.FreezeObj(o_wmap)

                    for o_wmap in c_wmap:
                        Application.DeactivateAbove(o_wmap.FullName, False)

    return True
