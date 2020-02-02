# Sequence Comparator

### Description
Compare two xlsx files whose rows are pairs `(sequence, frequency)`, generating the following lists (along with the corresponding frequency):
- common sequences
- sequences appearing only in the first file
- sequences appearing only in the second file


### Usage
```text
usage: SequencesComparator.py [-h] -i1 IN_FILE1 -i2 IN_FILE2 [-m MIN_FREQ]
                              [-mc MIN_FREQ_COMMON] [-o OUTPUT_FILE]

Sequences comparator

optional arguments:
  -h, --help            show this help message and exit
  -i1 IN_FILE1, --file1 IN_FILE1
                        First XLSX input file
  -i2 IN_FILE2, --file2 IN_FILE2
                        Second XLSX input file
  -m MIN_FREQ, --min-freq MIN_FREQ
                        Minimum frequency for a sequence to be considered
                        valid
  -mc MIN_FREQ_COMMON, --min-freq-common MIN_FREQ_COMMON
                        Minimum frequency for a common sequence to be
                        considered valid
  -o OUTPUT_FILE, --output-file OUTPUT_FILE
                        Output XLSX file
```
