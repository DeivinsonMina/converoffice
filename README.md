# PDF to Office Document Converter

This project is a web application that allows users to upload a PDF file and convert it to a selected Office document format (such as DOCX or XLSX). The application is built using Flask and provides a simple interface for file uploads and conversions.

## Features

- Upload PDF files
- Select the desired Office document format for conversion
- Download the converted document

## Requirements

To run this application, you need to have Python installed on your machine. You also need to install the required dependencies listed in the `requirements.txt` file.

## Installation

1. Clone the repository:

   ```
   git clone <repository-url>
   cd pdf-to-office-app
   ```

2. Install the required packages:

   ```
   pip install -r requirements.txt
   ```

## Running the Application

To start the Flask web server, run the following command:

```
python app.py
```

The application will be accessible at `http://127.0.0.1:5000`.

## Usage

1. Open your web browser and navigate to the application URL.
2. Use the provided form to upload a PDF file and select the desired output format.
3. After the conversion is complete, you will be redirected to a results page where you can download the converted document.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.