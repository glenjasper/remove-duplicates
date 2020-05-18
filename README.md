remove-redundancy
======================
[![License](https://poser.pugx.org/badges/poser/license.svg)](./LICENSE)

Script that eliminates the redundancy of at least two .xlsx formatted files from Scopus, Web of Science, PubMed or Dimensions

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
- [Acknowledgments](#acknowledgments)

## Pre-requisites

### Python libraries

```sh
  $ sudo apt install -y python3-pip
  $ sudo pip3 install --upgrade pip
```

```sh
  $ sudo pip3 install argparse
  $ sudo pip3 install openpyxl
  $ sudo pip3 install xlsxwriter
```

## Installation

### Clone

To clone and run this application, you'll need [Git](https://git-scm.com) installed on your computer. From your command line:

```bash
  # Clone this repository
  $ git clone https://github.com/glenjasper/remove-redundancy.git

  # Go into the repository
  $ cd remove-redundancy

  # Run the app
  $ python3 remove-redundancy.py --help
```

### Download

You can [download](https://github.com/glenjasper/remove-redundancy/archive/master.zip) the latest installable version of _remove-redundancy_.

## How To Use

```sh  
  $ python3 remove_redundancy.py --help
  usage: remove_redundancy.py [-h] -f FILES [-o OUTPUT] [--version]

  Script que elimina a redundância de pelo menos dois arquivos formatados .xlsx
  do Scopus, Web of Science, PubMed ou Dimensions

  optional arguments:
    -h, --help            show this help message and exit
    -f FILES, --files FILES
                          Arquivos formatados .xlsx, separados por vírgula
    -o OUTPUT, --output OUTPUT
                          Pasta de saida
    --version             show program's version number and exit

  Thank you!
```

## Author

* [Glen Jasper](https://github.com/glenjasper)

## Organization
* [Molecular and Computational Biology of Fungi Laboratory](http://lbmcf.pythonanywhere.com) (LBMCF, ICB - **UFMG**, Belo Horizonte, Brazil)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## Acknowledgments

* Aristóteles Góes-Neto
* Rosimeire Floripes
* Joyce da Cruz Ferraz
