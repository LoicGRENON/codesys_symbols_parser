import xml.etree.ElementTree as ET


def parse_comment(comment):
    if comment is None:
        return ''
    return comment.text.strip().replace('\\n', '\n')


def parse_attribute(attribute_elmt):
    if attribute_elmt is None:
        return None
    return attribute_elmt.text.split(':=')


class CodesysSymbolParser:
    # namespace to use to parse the XML file
    __namespace = {'ns': 'http://www.3s-software.com/schemas/Symbolconfiguration.xsd'}

    def __init__(self, symbols_file=''):
        self.symbols_file = symbols_file

        # self._simple_type_defs = {}
        self._usertype_defs = []

        self.root = None

    def parse(self, symbols_file=''):
        if symbols_file:
            self.symbols_file = symbols_file
        tree = ET.parse(self.symbols_file)
        self.root = tree.getroot()

        # self._simple_type_defs = self._extract_simpletype_defs()
        self._usertype_defs = self._extract_usertype_defs()

    def _extract_usertype_defs(self):
        """
        This method lists the user types definitions from the XML file.
        :return: A dictionary indexed on type names keys.
        """

        types = {}
        # Filter on Userdef typeclass attribute value as Enum are defined as TypeUserDef as well
        for type_def in self.root.findall(".//ns:TypeUserDef[@typeclass='Userdef']", namespaces=self.__namespace):
            type_name = type_def.get('name')
            elements = []
            for element in type_def.findall('ns:UserDefElement', namespaces=self.__namespace):
                # Comment is a sub-node of UserDefElement
                comment = element.find('ns:Comment', namespaces=self.__namespace)
                # Attribute is a sub-node of UserDefElement
                attribute_elmt = element.find('ns:Attribute', namespaces=self.__namespace)
                attribute = parse_attribute(attribute_elmt)

                # Ignore elements with attribute hmi_ignore
                if attribute is not None and attribute[0] == 'hmi_ignore':
                    continue

                element_info = {
                    'type': element.get('type'),  # required attribute
                    'iecname': element.get('iecname'),  # required attribute
                    'byteoffset': element.get('byteoffset'),  # optional attribute
                    'vartype': element.get('vartype'),  # optional attribute
                    'enumvalue': element.get('enumvalue'),  # optional attribute
                    'compileroffset': element.get('compileroffset'),  # optional attribute
                    'bitoffset': element.get('bitoffset'),  # optional attribute
                    'inherited_from': element.get('inherited_from'),  # optional attribute
                    'propertytype': element.get('propertytype'),  # optional attribute
                    'access': element.get('access'),  # optional attribute
                    'comment': parse_comment(comment)
                }
                elements.append(element_info)
            types[type_name] = elements
        return types

    # def extract_simpletype_defs(self):
    #     """
    #     This method lists the simple types definitions from the XML file.
    #     :return: A dictionary indexed on type names keys.
    #     """
    #     types = {}
    #     for type_def in self.root.findall(".//ns:TypeSimple", namespaces=self.__namespace):
    #         type_name = type_def.get('name')
    #         types[type_name] = {
    #             'size': type_def.get('size'),
    #             'swapsize': type_def.get('swapsize'),
    #             'typeclass': type_def.get('typeclass'),
    #             'iecname': type_def.get('iecname'),
    #             'basetype': type_def.get('basetype'),
    #             'aliasedtype': type_def.get('aliasedtype'),
    #             'aliasediecname': type_def.get('aliasediecname'),
    #             'underlyingtype': type_def.get('underlyingtype'),
    #             'lowerborder': type_def.get('lowerborder'),
    #             'upperborder': type_def.get('upperborder'),
    #         }
    #     return types

    def _get_type_element_paths(self, type_name, parent_path):
        """
        Recursive method to get each member of a specific type definition identified by its name.
        :param type_name: Name of the type definition.
        :param parent_path: parent path of the current node to concatenate with.
        :return: A list containing the symbol data for the specified type.
        """

        paths = []
        if type_name in self._usertype_defs:
            for element in self._usertype_defs[type_name]:
                current_path = f"{parent_path}.{element['iecname']}"

                # Add members of SimpleType. Testing on UserType definition make it to work as well for ArrayType
                if element['type'] not in self._usertype_defs:
                    paths.append({
                        'name': current_path,
                        'comment': element['comment']
                    })
                else:  # Recursive call to add the sub-members
                    paths.extend(self._get_type_element_paths(element['type'], current_path))
        return paths

    def _get_node_paths(self, node, current_path=""):
        """
        Recursive method to traverse all Node elements (so-called symbols) and to return them as a list.
        :param node:
        :param current_path: A string representing the current parent path of the symbol which is constructed recursively.
        :return: A list of symbols from Node elements.
        """

        paths = []
        node_name = node.get('name')
        node_type = node.get('type')
        current_path = f"{current_path}.{node_name}" if current_path else node_name
        child_nodes = node.findall('./ns:Node', namespaces=self.__namespace)

        # Add the elements depending on the type of the current node (for example, the structure members)
        paths.extend(self._get_type_element_paths(node_type, current_path))

        # Add the current node to the symbols list only if it is the last node (no children) and is of simple type
        if not child_nodes and node_type not in self._usertype_defs:
            comment = node.find('ns:Comment', namespaces=self.__namespace)
            # if comment is not None:
            #     print(comment.text)
            paths.append({
                'name': current_path,
                'comment': parse_comment(comment)
            })
        else:
            for child in child_nodes:
                paths.extend(self._get_node_paths(child, current_path))

        return paths

    def get_symbols(self):
        """
        Traverse all the Node nodes from NodeList element to get the symbols' data.
        :return: A list of dictionnaries containing the symbols' data.
        """
        return [symbol
                for node in self.root.findall('.//ns:NodeList', namespaces=self.__namespace)
                for symbol in self._get_node_paths(node)]


if __name__ == '__main__':
    import csv
    from pathlib import Path

    symbols_filepath = '../assets/PZ_PLC.MyController.Application.xml'
    parser = CodesysSymbolParser(symbols_filepath)
    parser.parse()
    symbols = parser.get_symbols()

    print(f'{len(symbols)} symbols found.')

    output_filepath = Path(symbols_filepath).with_suffix('.csv')
    fieldnames = ['name', 'comment']
    with open(output_filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writerows(symbols)
