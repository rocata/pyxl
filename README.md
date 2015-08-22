# pyxl
Custom python script which transforms a set of data (JSON) into a spreadsheet (XLSX)

### Version
1.0.0

### Prerequisites
* python 2.7
* pip

### Installation
```sh
pip install -r requirements.txt
```
### Usage
```sh
$ python transform_it.py <input_file> <output_file>
$ python transform_it.py raw_data.json pretty_data.xlsx
```

### Notes
> The input file will be type checked before performing any other operations.
> So, if the path is wrong, or if the JSON file is corrupted, the script will stop without touching the output file

> The output file will ALWAYS be overwritten if the input file is being processed.
 
