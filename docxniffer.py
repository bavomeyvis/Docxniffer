import mammoth
import argparse
import os
from pathlib import Path
from zipfile import ZipFile
from zipfile import BadZipfile

pathroot = os.getcwd()

#def docx_sniffer():



def conv_docx(input_path, output_path, filetype):
    # Set the starting directory for the input
    print('input_path = %s'%input_path)
    print('output_path = %s'%output_path)
    print('filetype = %s'%filetype)
    
    result = ''
    for dirName, subdirList, fileList in os.walk(input_path):
        print("Entered directory %s" % dirName)
        for fname in fileList:
            if fname.endswith('.docx'):
                

                # output_path = ../../AFolder/AnotherFolder


                # Get the path of input file
                source_file_path = dirName+'/'+fname
                destination_file_path = ''
                try:
                    # Open input docx and save output file
                    with ZipFile(source_file_path) as zf:
                        print('Converting: %s\n' % source_file_path)
                        with open(source_file_path, "rb") as docx_file:
                            result = mammoth.convert_to_markdown(docx_file)
                            # Output to correct path with new extension
                        if input_path == output_path:
                            destination_file_path = Path(source_file_path).with_suffix(filetype)
                        else:
                            new_path = output_path+'/'+fname
                            destination_file_path = Path(new_path).with_suffix(filetype)
                        print('Saving %s'%destination_file_path)
                        with open(destination_file_path, "w") as markdown_file:
                            markdown_file.write(result.value)
                except BadZipfile:
                    print("[EXCEPT]: BadZipFile. Skipping %s...\n"%fname)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Docx converter")

    parser.add_argument("--input_path",
                        help="The folder that needs to be searched. Defaults to where script is run.",
                        nargs='?',
                        default=".")
    parser.add_argument("--output_path",
                        help="Save location",
                        nargs='?',
                        default=".")
    parser.add_argument("--filetype",
                    help="Filetype to convert to. E.g. '.md'",
                    nargs='?',
                    default=".md")
    args = parser.parse_args()
    conv_docx(args.input_path, args.output_path, args.filetype)