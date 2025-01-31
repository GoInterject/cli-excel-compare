import subprocess
import os


class ExcelConvertor:
    def run_excel_convert_to_json(excel_file: str, output_folder: str):
        # Path to JS code of JSON convertor.
        js_path = '../js-excel-to-json-convertor/main.js'

        # Absolute path to your temp.js file
        js_file_path = os.path.abspath(js_path)

        # Ensure the file exists
        if not os.path.exists(js_file_path):
            print(f'File "{js_file_path}" not found!')

        # Run the JavaScript code using Node.js
        popen_args_array = ['node', js_file_path, excel_file]
        if output_folder:
            popen_args_array.append(output_folder)
        process = subprocess.Popen(popen_args_array, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        stdout, stderr = process.communicate()
        # Print the output
        print("\nstdout:\n", stdout)
        if stderr:
            print("\nstderr:\n", stderr)


