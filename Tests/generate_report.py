import tempfile
from comtypes.client import CreateObject
from wordinserter import insert, parse
from selenium import webdriver
import pathlib
import time
import subprocess
import os
import glob
import CommonMark
import sys

if not os.path.exists("images"):
    os.mkdir("images")
else:
    import shutil
    shutil.rmtree("images")
    os.mkdir("images")

image_directory = pathlib.Path("images")
temp_directory = pathlib.Path(tempfile.mkdtemp())

word = CreateObject("Word.Application")
word.Visible = False

from comtypes.gen import Word as constants

imagemagick = "convert.exe"

# [(file name, word image, html image)]
results = []


if __name__ == "__main__":
    if len(sys.argv) != 1:
        file_names = [pathlib.Path(p) for p in sys.argv[1:]]
    else:
        file_names = list(pathlib.Path("docs").iterdir())

    browser = webdriver.Firefox()

    for file in file_names:
        if file.name.endswith(".html"):

            document = word.Documents.Add()

            try:
                content = file.open().read()

                if not content.strip():
                    continue

                print("Rendering {0}".format(file.name))

                parse_start = time.time()
                operations = parse(content, parser="html" if file.name.endswith("html") else "markdown")
                parse_end = time.time()

                render_start = time.time()
                insert(operations, document=document, constants=constants)
                render_end = time.time()

                # Now we export the document as a PDF:
                pdf_path = temp_directory / (file.name + ".pdf")
                document.SaveAs2(str(pdf_path), 17)

                # Now convert that PDF to a PNG
                png_path = temp_directory / (file.name + ".png")

                code = subprocess.call(
                    '"{0}" -interlace none -density 300 -trim -quality 80 {1} {2} '.format(
                    imagemagick,
                    str(pdf_path),
                    str(png_path)
                ), shell=True)

                if code != 0:
                    print("Error converting PDF to PNG, exit code {0}".format(code))
                    continue

                if not png_path.exists():
                    # Multiple images, we need to combine them together
                    combined_images = glob.glob(str(temp_directory) + file.name + "-*.png")
                    code = subprocess.call(
                        '"{0}" {1} -append {2}'.format(
                            imagemagick,
                            " ".join(combined_images),
                            str(png_path)
                        )
                    )

                    if code != 0:
                        print("Error combining multiple PNG files together, exit code {0}".format(code))
                        continue

                if file.name.endswith(".html"):
                    # Just load it
                    browser.get("file://{0}".format(file.absolute()))
                else:
                    # Render it to HTML, save it to a temp file then open it
                    p = CommonMark.DocParser()
                    html = CommonMark.HTMLRenderer().render(p.parse(content))
                    tf = tempfile.mktemp(".html")
                    with open(tf,"w") as f:
                        f.write(html)
                    browser.get(tf)

                html_path = temp_directory / (file.name + ".html.png")

                browser.save_screenshot(str(html_path))

                code = subprocess.call(
                    '"{0}" -trim {1} {2} '.format(
                    imagemagick,
                    str(html_path),
                    str(html_path)
                ), shell=True)

                if code != 0:
                    print("Error trimming browser screenshot, exit code {0}".format(code))
                    continue

                results.append((file.name, str(png_path), str(html_path)))

                print("Parse: {0:10.7f} | Render: {1:10.7f}".format(
                    parse_end - parse_start,
                    render_end - render_start
                ))

            finally:
                document.Close(SaveChanges=constants.wdDoNotSaveChanges)


    with open("report.html", "w") as fd:

        fd.write("""
        <html>
        <head>
            <style>
                img {max-width: 400px;}
            </style>
        </head>
        <body>
        <table border=1>
            <thead>
                <th>File Name</th>
                <th>WordInserter Output</th>
                <th>Firefox Output</th>
            </thead>
            <tbody>
        """)
        for name, png_path, html_path in results:
            new_png_path = str(image_directory / (name + ".png"))
            new_html_path = str(image_directory / (name + ".html.png"))

            shutil.copy(png_path, new_png_path)
            shutil.copy(html_path, new_html_path)

            fd.write("<tr>")
            fd.write("<td>{0}</td>".format(name))
            fd.write("<td><img src='{0}'></img></td>".format(new_png_path.replace("\\", "/")))
            fd.write("<td><img src='{0}'></img></td>".format(new_html_path.replace("\\", "/")))
            fd.write("</tr>")

        fd.write("</tbody></table>")

        fd.write("</html></body>")

    browser.quit()
    import webbrowser
    webbrowser.open("report.html")
