import os
import xml.dom.minidom

class XMLPrettyPrint:
    def pretty_print_xml(self, xml_string):
        """Convert an XML string to a pretty-printed format."""
        parsed_xml = xml.dom.minidom.parseString(xml_string)
        return parsed_xml.toprettyxml(indent="  ")

    def convert_xml_files_in_folder(self, folder_path):
        """Convert all XML files in the specified folder and its subfolders to pretty print."""
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:
                if file_name.endswith('.xml'):
                    file_path = os.path.join(root, file_name)
                    self.convert_xml_file_in_folder(file_path)

    def convert_xml_file_in_folder(self, file_path):
        print(f'Processing {file_path}')
        
        # Read the XML file
        with open(file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        # Pretty print the XML
        pretty_xml = self.pretty_print_xml(xml_content)

        # Write the pretty-printed XML back to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(pretty_xml)
