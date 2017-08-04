# -*- coding: utf-8 -*-
"""
/***************************************************************************
 EnpaTools
                                 A QGIS plugin
 This plugin is a smorgasbord of tools to make a rangers life easier
                             -------------------
        begin                : 2017-06-22
        copyright            : (C) 2017 by ENPA
        email                : enquiries@exmoor-nationalpark.gov.uk
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load EnpaTools class from file EnpaTools.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
      
    from .enpa_tools import EnpaTools
    return EnpaTools(iface)


