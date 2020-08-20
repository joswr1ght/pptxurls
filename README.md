pptxurls
======

Generate a Markdown file of all links in one or more PowerPoint documents.

## Linux or OS X Installation

On Linux or Mac OS X systems, install the `python-pptx` package using `pip`:

```
$ pip install python-pptx
```

If you don't have the `pip` utility, install it with the following command,
then run the `pip` command, above:

```
$ sudo easy_install pip
```

## Windows Binary Installation

Download the
[`pptxurls.exe`](https://github.com/joswr1ght/pptxurls/blob/master/bin/pptxurls.exe)
program from the GitHub project `bin/` folder. Copy it to a location in your
system PATH.


## Windows Python Installation

Instead of using the `pptxurls.exe` binary, Windows can install Python to use
pptxurls.  Download and install the latest [Python
3](https://www.python.org/downloads/windowsi).  During the installation
process, add the Python directory to your PATH.

After installing Python run the `pip` utility to install the dependencies:

```
C:\temp> pip install python-docx python-pptx
Collecting python-docx
...
Successfully installed Pillow-6.1.0 XlsxWriter-1.1.8 lxml-4.3.4 python-docx-0.8.10 python-pptx-0.6.18
```

## Usage

```
$ python pptxurls.py -h
usage: pptxurls.py [-h] [-m MDFILE] [-t TITLE] [pptxfiles [pptxfiles ...]]

positional arguments:
  pptxfiles

optional arguments:
  -h, --help            show this help message and exit
  -m MDFILE, --mdfile MDFILE
                        Markdown output report filename
  -t TITLE, --title TITLE
                        Label each URL with its page title (Off by default)
```

If you don't specify the `-m` argument, Pptxurls will create a file
called `pptxurls.md` in the current directory.

## Sample Usage

```
$ python pptxurls.py *.pptx -m Links.md
pptxurls (master) $ head -10 Links.md
# List of Links

Use this resource to easily open links included in your printed materials.

## Book 1

| Page | URL |
|------|-----|
| 9 | [http://www.sans.org](http://www.sans.org) |
| 13 | [http://www.sans.org/score/incidentforms](http://www.sans.org/score/incidentforms) |

$ python pptxurls.py *.pptx -m Links.md -t true
pptxurls (master) $ head -10 Links.md
# List of Links

Use this resource to easily open links included in your printed materials.

## Book 1

| Page | URL |
|------|-----|
| 9 | [Information Security Training \| SANS Cyber Security Certifications & Research](http://www.sans.org) |
| 13 | [Sample Incident Handling Forms \| SCORE \| SANS Institute](http://www.sans.org/score/incidentforms) |
```

## On Page Numbering

Pptxurls accepts one or more PowerPoint files to use for building the URL
list. When multiple PowerPoint files are specified, it
creates the Markdown with headings indicating Book 1, Book 2, etc.

The book number is taken from the PowerPoint filename order, alphanumerically
sorted.  Consider the following PowerPoint filenames:

```
$ ls *.pptx
Sec575_1_A09.pptx
Sec575_2_A09.pptx
Sec575_3_A09.pptx
```

Here, the `Sec575_1_A09.pptx` file will be marked as book number 1. The file
`Sec575_2_A09.pptx` would be page 2, etc.  This assumes you don't name your
files like this:

```
$ ls *.pptx
Sec575_MobArch_1_A09.pptx
Sec575_MobPentest_2_A09.pptx
Sec575_MobDevRecommend_3_A09.pptx
```

In this naming convention, the file `Sec575_MobDevRecommend_3_A09.pptx` would
be alphanumerically sorted as book 2, and `Sec575_MobPentest_2_A09.pptx` would
be book 3.  *Don't name your files this way.*


## Questions, Comments, Concerns?

Open a ticket, or drop me a note: jwright@hasborg.com.

-Josh
