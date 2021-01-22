#! python3
# pdf_editor.py - Terminal app for some more commonly needed PDF functions


import json
import logging
import os
import os.path
from pathlib import Path
import subprocess
import time

from natsort import natsorted
import pyinputplus as pyip
import PyPDF2
from tqdm import tqdm


def enter_to_quit():
    input("Press ENTER to quit...")
    quit()


def enter_to_continue():
    input("\nPress ENTER to return to the main menu...")
    time.sleep(1.5)
    return


def move_to_parent_dir(filepath):
    p = Path(filepath)
    os.chdir(p.parent)
    return


class Pdf_Cache:
    def __init__(self):

        os.chdir(Path(__file__).parent)

        try:
            with open("pdf_cache.json") as infile:
                self.pdfs = json.load(infile)
        except Exception as e:
            print(e)
            print("Failed to load PDF cache from file.")
            self.update_cache()

        return

    def update_cache(self):

        os.chdir(Path(__file__).parent)
        print("This could take a few minutes. Please don't close this window.")
        drives = [
            f"{drive}:\\"
            for drive in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            if os.path.exists(f"{drive}:")
        ]

        pdf_path_lst = []
        for drive in tqdm(drives, desc="Searching connected drives for PDFs"):
            for path, dirs, files in os.walk(drive):
                pdf_path_lst += [
                    os.path.abspath(file) for file in files if file.endswith(".pdf")
                ]

        self.pdfs = {}
        for pdf_path in pdf_path_lst:
            basename = os.path.basename(pdf_path)
            self.pdfs[basename] = pdf_path

        try:
            with open("pdf_cache.json", "w") as outfile:
                json.dump(self.pdfs, outfile)
            print("PDFs cached successfully!")
        except Exception as e:
            print(e)
            print("PDF caching failed.")
            enter_to_quit()

        return


def get_user_choice():
    title = "\n" + " PDF Editor ".center(30, "-") + "\n"
    print(title)

    choices = ["Combine PDFs", "Delete pages of a PDF", "Rotate pages of a PDF", "Quit"]

    print("What would you like to do?")
    response = pyip.inputMenu(choices, numbered=True, blank=True)

    return response


def get_pdf(Pdf_Cache):
    pdf_name = None
    while pdf_name != "":
        pdf_name = pyip.inputStr(
            prompt="\nEnter the name of a PDF (or ENTER to continue): ", blank=True
        )
        if pdf_name == "":
            return None
        elif pdf_name.endswith(".pdf") is False:
            pdf_name += ".pdf"

        print(f"Searching for {pdf_name}...")

        if pdf_name in Pdf_Cache.pdfs.keys():
            pdf_path = Pdf_Cache.pdfs[pdf_name]
            print(f"Found PDF {pdf_path}!")
            return pdf_path
        else:
            print(f"\nCould not find {pdf_name}.")
            update_cache_yn = pyip.inputYesNo(
                "Do you want to update the PDF cache again (Y/N)? "
            )
            if update_cache_yn == "yes":
                Pdf_Cache.update_cache()
            else:
                print(f"Check the file name and try again.")


def write_and_open_new_pdf(pypdf2_class):
    while True:
        outfile_name = pyip.inputStr(
            prompt="\nEnter a name for your new PDF: ", blank=True
        )
        if outfile_name == "":
            return
        elif outfile_name.endswith(".pdf") is False:
            outfile_name += ".pdf"

        if Path(outfile_name).exists():
            print(f"\n{outfile_name} already exists. Please try again.")
        else:
            break

    try:
        with open(outfile_name, "wb") as outfile:
            pypdf2_class.write(outfile)
    except:
        print(f"Failed to write {outfile_name} to file. Please try again.\n")
        return

    print(f"\nSuccessfully written {outfile_name} to file.")
    print(f"Opening {outfile_name}...")
    subprocess.Popen(f"explorer \select, {outfile_name}")
    os.startfile(outfile_name)
    return


def combine_pdfs():
    print(f"\n {' COMBINING PDFs '.center(30, '+')}")

    pdf = ""
    pdfs = []
    while pdf is not None:
        pdf = get_pdf(Pdf_Cache)

        if pdf is None:
            break

        pdfs.append(pdf)
        if len(pdfs) > 1:
            print("\nCombining:")
            for pdf in pdfs:
                if pdf is not None:
                    print(pdf)

    if len(pdfs) < 2:
        enter_to_continue()
        return

    pdfs = [pdf for pdf in pdfs if pdf is not None]
    move_to_parent_dir(pdfs[0])

    pdf_merger = PyPDF2.PdfFileMerger()
    for pdf in pdfs:
        file = PyPDF2.PdfFileReader(pdf)
        pdf_merger.append(file)

    write_and_open_new_pdf(pdf_merger)
    enter_to_continue()
    return


def delete_pages():
    print(f"\n{' DELETING PAGES '.center(30, 'x')}")

    pdf = get_pdf(Pdf_Cache)
    if pdf is None:
        enter_to_continue()
        return

    move_to_parent_dir(pdf)

    with open(pdf, "rb") as infile:
        pdf_reader = PyPDF2.PdfFileReader(infile)
        total_pages = pdf_reader.numPages
        print(f"\nThere are {total_pages} pages in {pdf}")

        user_input = None
        user_input_pages = []
        while user_input != "":
            user_input = pyip.inputStr(
                prompt=f"\nEnter the page number(s) you want to delete, separated by commas. Use hyphens for ranges: ",
                blank=True,
            )

            if user_input != "":
                try:
                    user_input_pages += [s.strip() for s in user_input.split(",")]
                    user_input_pages = natsorted(list(set(user_input_pages)))

                    idx_to_del = []
                    for x in user_input_pages:
                        if "-" in x:
                            x = x.split("-")
                            x = [int(n) for n in x]
                            range_of_idx = [i - 1 for i in range(min(x), max(x))]
                            idx_to_del += range_of_idx
                        elif int(x) < 0 or int(x) > total_pages:
                            raise ValueError
                        else:
                            idx = int(x) - 1
                            idx_to_del.append(idx)

                    idx_to_del = sorted(list(set(idx_to_del)))
                    print(f"\nYou've entered: page(s) {', '.join(user_input_pages)}")

                except ValueError:
                    print(f"Please enter valid page number(s) or range(s).")

            elif len(user_input_pages) < 1:
                enter_to_continue()
                return

        confirm_del = pyip.inputYesNo(
            prompt=f"Confirm deleting page(s) {', '.join(user_input)} (Y/N): ",
            blank=True,
        )

        if confirm_del in ["", "no"]:
            enter_to_continue()
            return

        pages = [
            pdf_reader.getPage(i) for i in range(total_pages) if i not in idx_to_del
        ]
        pdf_writer = PyPDF2.PdfFileWriter()
        for page in pages:
            pdf_writer.addPage(page)

        write_and_open_new_pdf(pdf_writer)
        enter_to_continue()
        return


def rotate_pages():  # FIXME write
    print(f"\n {' ROTATING PAGES '.center(30, '~')}")

    pdf = get_pdf(Pdf_Cache)
    if pdf is None:
        enter_to_continue()
        return

    move_to_parent_dir(pdf)

    with open(pdf, "rb") as infile:
        pdf_reader = PyPDF2.PdfFileReader(infile)
        total_pages = pdf_reader.numPages

        pdf_writer = PyPDF2.PdfFileWriter()

        page_num = ""
        pages_turned = (
            []
        )  # FIXME return to main menu if no changes, write block if changes?
        while page_num is not None:
            print(f"There are {total_pages} page(s) in {pdf}")
            page_num = pyip.inputInt(
                prompt=f"Enter the page number of the page you want to rotate: ",
                min=1,
                max=total_pages,
                blank=True,
            )

            if page_num == "":
                enter_to_continue()
                return

            idx = pdf_reader.getPage(page_num - 1)
            num_turns = pyip.inputInt(
                prompt=f"Enter the number of CLOCKWISE turns to turn page {page_num}: ",
                min=0,
                max=3,
                blank=True,
            )

            if num_turns == "":
                continue

            page = pdf_reader.getPage(idx)
            page.rotateClockwise(num_turns * 90)
            print(f"Rotated page {page_num} by {num_turns*90} degrees.")

        write_and_open_new_pdf(pdf_writer)


def main():

    pdf_cache = Pdf_Cache()

    while True:
        try:
            user_choice = get_user_choice()
            if user_choice == "Combine PDFs":
                combine_pdfs()
            elif user_choice == "Delete pages of a PDF":
                delete_pages()
            elif user_choice == "Rotate pages of a PDF":
                rotate_pages()
            elif user_choice == "Quit" or user_choice == "":
                quit()
        except KeyboardInterrupt:
            quit()


logging.basicConfig(level=logging.DEBUG, format="\t%(levelname)s - %(message)s")
logging.disable(logging.CRITICAL)


main()
