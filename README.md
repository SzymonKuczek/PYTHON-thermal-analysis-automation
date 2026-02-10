# Thermal Bridge Automation Tool 🏗️🌡️



An engineering automation tool written in Python designed to calculate the U-value (thermal transmittance) and visualize temperature distribution in multi-layer building partitions.



# 🚀 Key Features

- **Symbolic Calculation:** Uses `SymPy` to dynamically derive physical equations for any number of layers, eliminating manual formula errors.

- **Automated Reporting:** Reads geometry and material data from Excel files (`openpyxl`) and generates visual plots automatically.

- **Visualization:** Generates temperature gradient charts using `Matplotlib` to identify potential freezing points inside the structural wall.

- **Batch Processing:** Capable of processing hundreds of project files simultaneously in a single run.



# 🛠️ Tech Stack

- **Python 3.10+**

- **SymPy**(Symbolic mathematics \& physics formulas)

- **Matplotlib**(Data visualization)

- **OpenPyXL** (Excel data integration)



# 📂 Project Structure



├── data           [Input Excel files (materials db \& project definitions)]

├── docs           [Generated PDF reports and documentation]

├── images         [Output charts (Temperature gradients)]

├── src            [Source code (Python scripts)]

└── README.md      [Project documentation]



# 📊 Example Results

1. Passive House Standard (U = 0.12 W/m²K)

The structural layer remains in the positive temperature zone (approx. +19°C).



2. Uninsulated Brick Wall (U = 1.25 W/m²K)

Rapid temperature drop across the entire structure, high risk of freezing.



# 🔧 How to Run

## Clone the repository:

`git clone \[https://github.com/SzymonKuczek/PYTHON\_thermal\_analysis\_automation.git](https://github.com/SzymonKuczek/PYTHON\_thermal\_analysis\_automation.git)`

`cd PYTHON\_thermal\_analysis\_automation`



## Install dependencies:

`pip install -r requirements.txt`



## Generate example data (Optional): This script creates sample Excel files in the data/ directory.

`cd src`

`python data\_gen.py`



## Run the analysis: This script processes all Excel files in data/ and saves charts to images/.

`python analysis.py`



# 🚀 Future Improvements

[ ] Implementation of Glaser Method (Dew point calculation) to predict interstitial condensation risk.


[ ] Integration with .CSV weather data for dynamic simulations.


[ ] GUI (Graphical User Interface) using PyQt or Tkinter.



# 📄 Documentation

📄🇵🇱 [PL] [Download the full technical report in Polish (PDF)](docs/Raport\_PYTHON\_AGH\_PL.pdf)

📄🇺🇸 [EN] [Download the full technical report in English (PDF)](docs/Report\_CFD\_PYTHON\_EN.pdf)



# 👨‍💻 Author

Szymon Kuczek 
Energy & Civil Engineering Student at AGH University of Science and Technology.

Focus: CFD, Building Physics, and Engineering Automation.

