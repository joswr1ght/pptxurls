#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# With code by Eric Jang ericjang2004@gmail.com
from pptx import Presentation
import sys
import re
import os
import shutil
import glob
import tempfile
import argparse

import signal
from zipfile import ZipFile
from xml.dom.minidom import parse


# Remove trailing unwanted characters from the end of URL's
# This is a recursive function. Did I do it well? I don't know.
def striptrailingchar(s):
    # The valid URL charset is A-Za-z0-9-._~:/?#[]@!$&'()*+,;= and & followed by hex character
    # I don't have a better way to parse URL's from the cruft that I get from XML content, so I
    # also remove .),;'? too.  Note that this is only the end of the URL (making ? OK to remove)
    # if s[-1] not in
    # "ABCDEFGHIJKLMNOPQRSTUVWXYZZabcdefghijklmnopqrstuvwxyzz0123456789-_~:/#[]@!$&(*+=":
    if s[-1] not in "ABCDEFGHIJKLMNOPQRSTUVWXYZZabcdefghijklmnopqrstuvwxyzz0123456789-_~:#[]@!$&(*+=/":
        s = striptrailingchar(s[0:-1])
    elif s[-5:] == "&quot":
        s = striptrailingchar(s[0:-5])
    else:
        pass
    return s


# Parse the given root recursively (root is intended to be the paragraph element <a:p>
# If we encounter a link-break element a:br, add a new line to global paragraphtext
# If we encounter an element with type TEXT_NODE, append value to paragraphtext
paragraphtext = ""


def parse_node(root):
    global paragraphtext
    if root.childNodes:
        for node in root.childNodes:
            if node.nodeType == node.TEXT_NODE:
                paragraphtext += node.nodeValue.encode(
                    'ascii', 'ignore').decode('utf-8')
            if node.nodeType == node.ELEMENT_NODE:
                if node.tagName == 'a:br':
                    paragraphtext += "\n"
                parse_node(node)

# Returns a dictionary of pagenum:url


def parseslidenotes(pptxfile, urls):
    global paragraphtext
    tmpd = tempfile.mkdtemp()

    ZipFile(pptxfile).extractall(path=tmpd, pwd=None)
    path = tmpd + os.sep + 'ppt' + os.sep + 'notesSlides' + os.sep

    for infile in glob.glob(os.path.join(path, '*.xml')):
        # parse each XML notes file from the notes folder.

        # Get the slide number
        slideNumber = re.match(r".*notesSlide(\d+).xml", infile).group(1)

        # Parse slide notes, adding a space after each paragraph marker, and
        # removing XML markup
        dom = parse(infile)
        paragraphs = dom.getElementsByTagName('a:p')
        for paragraph in paragraphs:
            paragraphtext = ""
            parse_node(paragraph)

            # Parse URL content from notes text for the current paragraph
            urlmatches = re.findall(urlmatchre, paragraphtext)
            if len(urlmatches) > 0:
                for match in urlmatches:  # Now it's a tuple
                    for urlmatch in match:
                        if urlmatch != '':
                            # urls.append([striptrailingchar(urlmatch), slideNumber])
                            urls.update(
                                {int(slideNumber): striptrailingchar(urlmatch)})

    # Remove all the files created with unzip
    shutil.rmtree(tmpd)
    return urls

# Parse the text on slides using the python-pptx module, return URLs


def parseslidetext(prs):
    urls = []
    singletextrun = ""
    slidenum = 0
    for slide in prs.slides:
        slidenum += 1
        text_runs = []
        for shape in slide.shapes:
            try:
                if not shape.has_text_frame:
                    continue
            except AttributeError:
                sys.stderr.write(
                    "Error: Please upgrade your version of python-pptx: pip "
                    "uninstall python-pptx ; pip install python-pptx\n")
                sys.exit(-1)
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    singletextrun += run.text
                text_runs.append(singletextrun)

            for text in text_runs:
                if text is None:
                    continue
                try:
                    m = re.match(urlmatchre, text)
                except IndexError:
                    continue
                if m is not None:
                    url = striptrailingchar(m.groups()[0])
                    if url not in urls:
                        urls.append([url, slidenum])
    return urls


def is_valid_file(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        return open(arg, 'r')  # return an open file handle


def signal_exit(signal, frame):
    sys.exit(0)


if __name__ == "__main__":
    signal.signal(signal.SIGINT, signal_exit)

    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-m',
        '--mdfile',
        dest='mdfile',
        default="pptxurls.md",
        help='Markdown output report filename')
    parser.add_argument(
        'pptxfiles',
        type=lambda x: is_valid_file(
            parser,
            x),
        action='store',
        nargs='*')
    args = parser.parse_args()

    # ensure that pptxfiles are provided
    if len(args.pptxfiles) == 0:
        print("No pptx files provided")
        sys.exit(0)

    try:
        mdfile = open(args.mdfile, 'w')
    except Exception:
        sys.stderr.write(
            f"Unable to open the output file {args.mdfile} for writing.\n")
        sys.exit(-1)

    # The PPTX files are treated as "books" in the report, with the first PPTX file on the command line
    # identified as book 1, the next as book 2, etc. Sory the PPTX files by
    # name to place them in order.
    args.pptxfiles.sort(key=lambda f: f.name)

    booknum = 0
    for pptxfile in args.pptxfiles:
        booknum += 1
        try:
            prs = Presentation(pptxfile.name)
        except Exception:
            sys.stderr.write("Invalid PPTX file: " + sys.argv[1] + "\n")
            sys.exit(-1)

        if (args.mdfile):
            mdfile.write("# List of Links\n\n")
            mdfile.write("Use this resource to easily open links included in your printed materials.\n\n")

            # If there is more than one book, write a header and open table
            if (len(args.pptxfiles) > 1):
                mdfile.write(f"\n## Book {booknum}\n\n| Page | URL |\n|-----|-----|\n")
            else:
                # Only a single book, just open the table for URLs
                mdfile.write(f"| Page | URL |\n|-----|-----|\n")

        # This may be the most insane regex I've ever seen.  It's very comprehensive, but it's too aggressive for
        # what I want.  It matches arp:remote in ettercap -TqM arp:remote // //, so I'm using something simpler
        # urlmatchre = re.compile(r"""((?:[a-z][\w-]+:(?:/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.‌​][a-z]{2,4}/)(?:[^\s()<>]+|(([^\s()<>]+|(([^\s()<>]+)))*))+(?:(([^\s()<>]+|(‌​([^\s()<>]+)))*)|[^\s`!()[]{};:'".,<>?«»“”‘’]))""", re.DOTALL)
        urlmatchre = re.compile(
            r'((https?://[^\s<>"]+|www\.[^\s<>"]+))', re.DOTALL)
        privateaddr = re.compile(
            r'(\S+127\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')

        urls = {}
        # TOFIX urls += parseslidetext(prs)
        parseslidenotes(pptxfile.name, urls)

        for pagenum in sorted(list(urls.keys())):
            url = urls[pagenum]
            url = url.encode('ascii', 'ignore').decode('utf-8')

            # Add default URI for www.anything
            if url[0:3] == "www":
                url = "http://" + url

            # Some authors include URLs in the form http://www.josh.net.[1], http://www.josh.net[1].
            # or http://www.josh.net[1]
            # Remove the footnote and/or leading or trailing dot.
            footnote = re.compile(r"(\.\[\d+\]|\[\d+\]\.|\[\d+\])")
            if re.search(footnote, url):
                url = re.sub(footnote, "", url)

            # Remove a trailing period
            if url[-1] == ".":
                url = url[:-1]

            # Skip private IP addresses
            if re.match(privateaddr, url):
                continue

            mdfile.write(f"| {pagenum} | [{url}]({url}) |\n")

    if os.name == 'nt':
        x = input("Press Enter to exit.")
