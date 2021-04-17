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


def convert(old_path, new_filetype):
    content = ''
    try:
        with ZipFile(old_path) as zf:
            print('Converting: %s...' % old_path)
            with open(old_path, "rb") as docx_file:
                if new_filetype == ".md": content = mammoth.convert_to_markdown(docx_file)
                else: content = mammoth.convert_to_html(docx_file)            
    except BadZipfile:
            print("[EXCEPT]: BadZipFile. Skipping %s...\n" % old_path)
            # File might: be corrupt, empty or not on path    
    return content.value
    

def save_new_file(new_file_path, content):
    print('Saving: %s...' % new_file_path)
    with open(new_file_path, "w") as new_file:
        new_file.write(content)



def dogwhisperer(input_path, output_path, method):
    print('Running script with arguments: \n\t{0} (input)\n\t{1}(output)\n\t{2}Not implemented (method)\n'.format(input_path, output_path, method))

    sources = docxniffer(input_path)
    user_input = input("\nThe listed files are to be converted to markdown-files relative to %s\nPress any key to continue..."%output_path)

    destination_file_path = ''
    for source in sources:
        # Get the content of docx in text-format
        md_content = convert(source, 'md')
        if md_content != '':
            # Use relative path for new file
            if input_path == output_path: 
                destination_file_path = Path(source).with_suffix('.md')
            # Use absolute path for new file
            else:
                fn = os.path.basename(source)
                absolute_path = output_path + '/' + fn
                destination_file_path = Path(absolute_path).with_suffix('.md')
            save_new_file(destination_file_path, md_content)

            # Check if file is large enough to attach an html version
            if(os.path.getsize(destination_file_path) > 10000):
                    print("\t%s is bigger than 10000 bytes. Attaching a html-file\n"%destination_file_path)
                    html_content = convert(source, '.html')
                    html_file_path = Path(destination_file_path).with_suffix('.html')
                    save_new_file(html_file_path, html_content)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Docx converter")
    parser.add_argument("--input_path",
                        help="The folder that needs to be searched. Defaults to where script is run.",
                        nargs='?',
                        default=".")
    parser.add_argument("--output_path",
                        help="Saves either where file is retrieved (default)"+
                            "or at an absolute path specified by user input (recommended).",
                        nargs='?',
                        default=".")
    parser.add_argument("--method",
                    help="[Not implemented]",
                    nargs='?',
                    default="")
    args = parser.parse_args()
    dogwhisperer(args.input_path, args.output_path, args.method)