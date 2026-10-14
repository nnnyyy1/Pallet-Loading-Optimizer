# Pallet Loading Optimizer
## Live Demo

**Try the Pallet Loading Optimizer online:**

https://pallet-loading-optimizer.onrender.com/

A Flask-based engineering application for evaluating pallet loading capacity, testing box orientations, and visualizing pallet arrangements in 2D and 3D.

**Key Features**
- Box orientation comparison
- Height and weight constraint evaluation
- Pallet space utilization analysis
- Interactive 2D and 3D visualizations
- Excel-based batch calculations

*Note: The application runs on a free hosting instance, so the first request may take longer after inactivity.*
**A Flask-based pallet loading calculator with 2D/3D visualizations and batch Excel processing.**

This project helps explore how uniformly sized boxes can be arranged on a pallet under dimensional, height, and optional weight constraints. It is designed as a practical decision-support prototype for warehousing, production planning, and logistics.

## Features

- **Single-pallet calculation:** Enter box dimensions, pallet dimensions, maximum total height, and pallet height.
- **Orientation comparison:** Evaluate the six axis-aligned orientations of a box and select the one that fits the most boxes in the implemented grid layout.
- **Optional weight constraint:** Limit the number of boxes using box weight and maximum allowable pallet load.
- **Utilization metrics:** View the number of boxes per layer, total layers, total boxes, pallet base-area utilization, and available-volume utilization.
- **Visual output:** Inspect an SVG isometric (3D-style) drawing and a top-down layout diagram.
- **Batch calculation:** Upload an Excel spreadsheet to evaluate multiple products at once and display the results in a table.

## Tech Stack

- Python
- Flask and Jinja templates
- pandas and openpyxl (Excel processing)
- HTML, CSS, and SVG (interface and drawings)

## Run Locally

**Requirements:** Python 3 and pip.

1. Clone the repository:

   ```bash
   git clone https://github.com/nnnyyy1/Pallet-Loading-Optimizer.git
   cd Pallet-Loading-Optimizer
   ```

2. Create a virtual environment:

   ```bash
   python -m venv .venv
   ```

3. Activate it:

   **Windows PowerShell**

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

   **macOS / Linux**

   ```bash
   source .venv/bin/activate
   ```

4. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

5. Start the app:

   ```bash
   python app.py
   ```

6. Open **http://127.0.0.1:8080/** in your browser.

## Excel Batch Input

The application expects an Excel file with the following column headers (in Turkish, matching the current interface):

| Column | Meaning | Required |
| --- | --- | --- |
| `Ürün Kodu` | Product identifier | No |
| `Koli Uzunluğu` | Box length | Yes |
| `Koli Genişliği` | Box width | Yes |
| `Koli Yüksekliği` | Box height | Yes |
| `Koli Ağırlığı` | Box weight | No |
| `Palet Uzunluğu` | Pallet length | Yes |
| `Palet Genişliği` | Pallet width | Yes |
| `Maksimum Yükseklik` | Maximum total loaded height | Yes |
| `Palet Yüksekliği` | Pallet height | No |
| `Maksimum Palet Yükü` | Maximum total box weight on pallet | No |

Use a consistent length unit for all dimensions (for example, centimeters). Weight fields are expressed in kilograms.

## Model Assumptions and Limitations

The current calculation uses a **uniform rectangular grid layout** per layer. It evaluates box orientations, computes how many boxes fit along each pallet dimension, and applies height and optional weight limits.

- It does **not** search for mixed-orientation patterns within a layer or different patterns across layers.
- It does **not** model box crush strength, weight distribution, physical stability, securing materials, or pallet overhang.
- It is a planning aid, **not** a certified pallet-loading or transport-safety tool.
- The optional weight limit constrains total **box** weight in the current implementation; check pallet tare weight separately when assessing actual load limits.

## Deployment Note

This repository contains a Flask application. For public hosting (for example, on Render), use a production WSGI server such as Gunicorn, disable Flask's debug mode, and validate/limit uploaded files before deployment. **Do not use the local development server as a public production server.**

## Privacy

Avoid uploading personal data, confidential company documents, or sensitive inventory information to public demo deployments. The application is intended for demonstration and engineering analysis.

## Project Status

Portfolio / educational prototype. Improvements and feedback are welcome.
