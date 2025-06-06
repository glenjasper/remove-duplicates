remove-duplicates
======================
[![License](https://poser.pugx.org/badges/poser/license.svg)](./LICENSE)

This script eliminates the duplicated records from formatted .xlsx files from [Scopus](https://www.scopus.com), [Web of Science](https://clarivate.com/webofsciencegroup/solutions/web-of-science), [PubMed](https://www.ncbi.nlm.nih.gov/pubmed), [PubMed Central](https://www.ncbi.nlm.nih.gov/pmc), [Dimensions](https://app.dimensions.ai), Cochrane, Embase, IEEE, BVS, CAB, SciELO, or Google Scholar exported from [Publish or Perish](https://harzing.com/resources/publish-or-perish). Is mandatory that there be at least 2 different files from 2 different databases.

## Table of content

- [Pre-requisites](#pre-requisites)
    - [Python libraries](#python-libraries)
- [Installation](#installation)
    - [Clone](#clone)
    - [Download](#download)
- [How To Use](#how-to-use)
- [Author](#author)
- [Organization](#organization)
- [License](#license)
- [Required Attribution](#required-attribution)

## Pre-requisites

### Python libraries

```sh
  $ sudo apt install -y python3-pip
  $ sudo pip3 install --upgrade pip
```

```sh
  $ sudo pip3 install argparse
  $ sudo pip3 install xlsxwriter
  $ sudo pip3 install numpy
  $ sudo pip3 install pandas
  $ sudo pip3 install crossrefapi
  $ sudo pip3 install openpyxl
  $ sudo pip3 install tqdm
  $ sudo pip3 install colorama
```

## Installation

### Clone

To clone and run this application, you'll need [Git](https://git-scm.com) installed on your computer. From your command line:

```bash
  # Clone this repository
  $ git clone https://github.com/glenjasper/remove-duplicates.git

  # Go into the repository
  $ cd remove-duplicates

  # Run the app
  $ python3 remove_duplicates.py --help
```

### Download

You can [download](https://github.com/glenjasper/remove-duplicates/archive/master.zip) the latest installable version of _remove-duplicates_.

## How To Use

```sh
$ python3 remove_duplicates.py --help
usage: remove_duplicates.py [-h] -f FILES [-o OUTPUT] [--version]

This script eliminates the duplicated records from formatted .xlsx files from
Scopus, Web of Science, PubMed, PubMed Central, Dimensions, Cochrane, Embase,
ScienceDirect, IEEE, BVS, CAB, SciELO, or Google Scholar (Publish or Perish).
Is mandatory that there be at least 2 different files from 2 different
databases.

optional arguments:
  -h, --help            show this help message and exit
  -f FILES, --files FILES
                        .xlsx files separated by comma
  -o OUTPUT, --output OUTPUT
                        Output folder
  --version             show program's version number and exit

Thank you!
```

## Author

* [Glen Jasper](https://github.com/glenjasper)

## Organization
* [Molecular and Computational Biology of Fungi Laboratory](https://e3sys.com.br/grupo) (LBMCF, ICB - **UFMG**, Belo Horizonte, Brazil).

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

## Required Attribution

**Any use, distribution, or modification of this software must include one of the following attributions:**

> This software was developed by Glen Jasper (https://github.com/glenjasper), originally available at https://github.com/glenjasper/remove-duplicates

**or** (example citation):

> Jasper G. (2024). *format_input.py* script. Available at: https://github.com/glenjasper/remove-duplicates (Accessed 21 July 2024)
