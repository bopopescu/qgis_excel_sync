from collections import OrderedDict

from qgis.testing import unittest, start_app
from shpsync.core.project_handler import ProjectHandler

start_app()


class TestProjectHandler(unittest.TestCase):

    def setUp(self):
        pass

    def test_write_read_setting(self):
        ProjectHandler.writeSetting("test_tag", "attribute_string", "foo")
        ProjectHandler.writeSetting("test_tag", "attribute_int", 123)
        ProjectHandler.writeSetting("test_tag", "attribute_float", 12.35)
        ProjectHandler.writeSetting("test_tag", "attribute_bool", False)
        ProjectHandler.writeSetting("test_tag", "attribute_list",
                                    ['foo', 'bar', '71'])

        settings = OrderedDict()
        settings['attribute_string'] = (str, None)
        settings['attribute_int'] = (int, None)
        settings['attribute_float'] = (float, None)
        settings['attribute_bool'] = (bool, False)
        settings['attribute_list'] = (list, [])

        result = ProjectHandler.readSettings("test_tag", settings)

        self.assertEqual(result['attribute_string'], 'foo')
        self.assertEqual(result['attribute_int'], 123)
        self.assertEqual(result['attribute_float'], 12.35)
        self.assertEqual(result['attribute_bool'], False)
        self.assertEqual(result['attribute_list'],
                         ['foo', 'bar', '71'])
