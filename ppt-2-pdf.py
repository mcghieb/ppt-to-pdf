#!/usr/bin/env python3

import os
import subprocess
import sys

def convert_to_pdf(input_filepath, output_dir):
    try:
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                input_filepath
            ],
            check=True
        )
        print(f"Converted {input_filepath} to PDF successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error during conversion: {e}")
    except FileNotFoundError:
        print("LibreOffice is not installed or not added to the system's PATH.")
    except Exception as e:
        print(f"Exception caught: {e}")


def get_filename_list(input_dir):
    all_items = os.listdir(input_dir)

    return [item for item in all_items if item.endswith(".pptx")]

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Invalid number of arguments.\nUsage: {script} input_dir output_dir")
        exit(1)
    elif not os.path.isdir(sys.argv[1] or not os.path.isdir(sys.argv[2])):
        print("Usage: {script} input_dir output_dir")
        exit(2)

    input_dir = os.path.abspath(sys.argv[1])
    output_dir = os.path.abspath(sys.argv[2])

    for filename in get_filename_list(input_dir):
        full_filepath = os.path.join(input_dir, filename)
        convert_to_pdf(full_filepath, output_dir)