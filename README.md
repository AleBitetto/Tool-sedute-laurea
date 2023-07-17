# Download lista studenti seduta laurea

### Installation

Clone the repository with
`git clone https://github.com/AleBitetto/Tool-sedute-laurea.git`

From console navigate the cloned folder and create a new environment with:
```
conda env create -f environment.yml
conda activate lauree
python -m ipykernel install --user --name lauree --display-name "Python (Lauree)"
```
This will create a `lauree` environment and a kernel for Jupyter Notebook called `Python (Lauree)`



### Chrome Web Driver

To download data, this notebook relies on [`selenium`](https://selenium-python.readthedocs.io/) and [`ChromeDriver`](https://chromedriver.chromium.org/).

This requires a `chromedriver` executable which can be downloaded [here](https://chromedriver.chromium.org/downloads). Make sure that your `Chrome` version is the same as your `chromedriver` version.

`lauree` assumes that the `chromedriver` executable is located at `/WebDriver` in the main folder. To supply a different path, change the variable `chromedriver_path` in the notebook.

### Credentials

Credentials must be inputed manually into a `config.py` file. Create a text file with the following lines:
```
Codice_Fiscale = "codice fiscale"
Password = "password login esse3"
```
where you need to write the esse3 credentials (Fiscal Code and password). Then simply rename the file as `.py`

### Usage

Simply run the notebook `Download lista studenti and upload voti.ipynb` [here](https://github.com/AleBitetto/Tool-sedute-laurea/blob/master/Download%20lista%20studenti%20and%20upload%20voti.ipynb).
