# ArchaeoDB

ArchaeoDB is software designed to facilitate the management of archaeological data and analysis of zooarchaeological data. It allows for the cataloging, visualization, and manipulation of information about bone remains, helping researchers organize and interpret their data.

## Documentation

- English README: [README.md](README.md)  
- Portuguese README: [LEIA-ME.md](LEIA-ME.md)

## Key Features

- **Data Management**: Comprehensive management of Sites, Collections, Assemblages, Excavation Units, Levels, Materials, and Specimens.  
- **Database**: Uses SQLite for robust and portable data storage. Create new databases or open existing ones easily.  
- **Import & Export**:  
  - Create and update databases directly from Excel files.  
  - Export the entire database to Excel.  
  - Generate PDF reports for Materials, Units, and Specimens.  
- **Statistics & Visualization**:  
  - Generate charts for Material Types, Descriptions, and Quantities.  
  - Visualize Unit density.  
  - Calculate and plot NISP (Number of Identified Specimens) by Taxon.  
  - Export charts to PNG.  
- **User Experience**:  
  - Interface with hierarchical filtering.  
  - Multi-language support: English, Portuguese, and Spanish.  
  - Light and Dark themes.  

## Installation

### 1. Install Python

Download and install Python 3.10 or higher from:

https://www.python.org/downloads/

During installation on Windows, make sure to check:

- "Add Python to PATH"

After installation, confirm in the terminal:

```bash
python --version
```

### 2. Clone the Repository

```bash
git clone https://github.com/your-username/ArchaeoDB.git
cd ArchaeoDB
```

### 3. Run Setup

On Windows, execute:

```bash
setup.bat
```

The `setup.bat` script will:

- Create a virtual environment (venv)
- Install all required dependencies from `requirements.txt`

Wait until installation completes successfully.

---

## Run the Program

After running `setup.bat`, start the application by executing:

```bash
run.bat
```

The `run.bat` script will:

- Activate the virtual environment
- Launch the main application

Important: Always run `setup.bat` at least once before executing `run.bat`.

If dependencies are updated in `requirements.txt`, run `setup.bat` again.


## License

This project is licensed under the **GNU AFFERO GENERAL PUBLIC LICENSE  
Version 3, 19 November 2007 (AGPL-3.0)**.

Under the AGPL-3.0:

- You may use, modify, and redistribute the software.  
- Any modified version that is distributed or made available over a network must also be licensed under AGPL-3.0.  
- The complete corresponding source code must be made available to users interacting with the software over a network.  

See the `LICENSE` file for the full legal text.

## Warranty Disclaimer

This software is provided **"AS IS"**, without warranty of any kind, express or implied, including but not limited to:

- Merchantability  
- Fitness for a particular purpose  
- Non-infringement  

The authors and contributors shall not be liable for any claim, damages, or other liability, whether in contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software. Use at your own risk.

## Contributions

Contributions are welcome under the terms of the AGPL-3.0 license. By submitting code, you agree that your contribution will be licensed under the same license.

## Intended Use

ArchaeoDB is intended for academic, research, and educational purposes in zooarchaeology and related disciplines. It does not replace formal curation standards, institutional data governance policies, or regulatory compliance requirements.

## Contact

For questions, suggestions, or support:  
geraldo.pmj@gmail.com




