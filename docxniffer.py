import argparse
import mammoth
import os
from pathlib import Path
from zipfile import ZipFile
from zipfile import BadZipfile


def docxniffer(input_path):
    sources = []
    for dirName, subdirList, fileList in os.walk(input_path):
        for fname in fileList:
            if fname.endswith('.docx'):
                full_path = dirName+'/'+fname
                sources.append(full_path)
    for source in sources:
        print(source)
    return sources


def dogwhisperer(input_path, output_path, method):
    print('Running script with arguments: \n\t{0} (input)\n\t{1}(output)\n\t{2}Not implemented (method)\n'.format(input_path, output_path, method))

    sources = docxniffer(input_path)
    user_input = input("\nThe listed files are to be converted to markdown-files relative to %s\nPress any key to continue..."%output_path)
    
    result = ''
    for source in sources:
        print('Converting: %s...' % source)
        try:
            with ZipFile(source) as zf:
                # Temporarily save the content
                with open(source, "rb") as docx_file:
                    result = mammoth.convert_to_markdown(docx_file)
                # Determine whether to use relative or absolute path for output file
                destination_file_path = ''
                if input_path == output_path:
                    destination_file_path = Path(source).with_suffix('.md')
                else:
                    fname = os.path.basename(source) # TODO: Replace hacky solution
                    new_path = output_path+'/'+fname
                    destination_file_path = Path(new_path).with_suffix('.md')
                # Output content with new filetype to determined path 
                with open(destination_file_path, "w") as markdown_file:
                    markdown_file.write(result.value)
        except BadZipfile:
            print("[EXCEPT]: BadZipFile. Skipping %s...\n" % source)
            # File might: be corrupt, empty or not on path


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
    parser.add_argument("--method",
                    help="[Not yet implemented]",
                    nargs='?',
                    default="")
    args = parser.parse_args()
    dogwhisperer(args.input_path, args.output_path, args.method)