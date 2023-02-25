<h1> Memo Template Generator </h1>

<p> This script generates a memo template in either PDF or DOCX format based on user input. The memo template contains several placeholders that are replaced with user input values.
Installation </p>

- To use this script, you must have Python 3.x installed on your system. You can download Python from the official website here: https://www.python.org/downloads/

You also need to install the following dependencies:

    PySimpleGUI
    fpdf
    docx

To install these dependencies, run the following command in your terminal or command prompt:

pip install PySimpleGUI 
pip install fpdf 
pip install docx

<h2> Usage </h2>

To use the script: 
- Run the memo_template_generator.py file using Python: This will open a GUI form where you can enter the values for the memo template placeholders. 
- Once you have filled in all the values, click the "Submit" button to generate the memo template in the selected format (PDF or DOCX).

<p>You can choose the output format by selecting your preferred file output via text in the dialogue box.</p>

# The generated memo template will be saved to the "HAULAGE_GENERATED PDFs" directory on your desktop.


<h2>Limitations</h2>

The script only supports a single format (PDF or DOCX) per run. To generate the template in another format, you need to run the script again and select the desired format.

The template does not bolden and underline places where need be automatically, but you can edit the .docx file when it is generated. 


<h2>Contributing</h2>

If you find any bugs or have any suggestions for improving the script, feel free to open an issue or submit a pull request on GitHub.
License

This project is licensed under the MIT License. See the LICENSE file for more information.
