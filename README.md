# Programming-Laboratory--3-python-TA3

Our project, the Automatic Certificate Generator, 
facilitates the creation of certificates by automatically extracting information from an Excel file. 
Through this system, certificates are efficiently generated with minimal manual input, streamlining the process and enhancing overall productivity.

#Execution of the Program
1. The script starts by importing necessary modules, including `datetime`, `Path` from `pathlib`, `pandas`, and `DocxTemplate` from `docxtpl`.
2. It sets the base directory, template path, output directory, and Excel file path using the `Path` class and creates the `output_dir` if it doesn't exist.
3. Reads the data from the 'students_name.xlsx' Excel file into a Pandas DataFrame, converting the 'date' column to the desired date format.
4. Iterates through the DataFrame records, creating a DocxTemplate object for each record using the specified 'CertificateFormate.docx' template.
5. Renders the template with the record data, including the formatted date, and saves the generated certificate as a new Word document in the 'Generated Certificate' directory.
6. The script utilizes pathlib to handle file paths robustly and ensures the 'Generated Certificate' directory is created to store the output.
7. The result is individualized certificates for each student based on the provided template, containing the student's name and the formatted current date.
8. The script is designed to enhance efficiency in creating certificates, leveraging Pandas for data manipulation and Docxtpl for rendering dynamic templates.
