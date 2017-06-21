import pathlib
import shutil
import subprocess
import sys

from selenium import webdriver
from wand.color import Color
from wand.image import Image

image_directory = pathlib.Path("images")

if image_directory.exists():
    shutil.rmtree(str(image_directory))

# If the directory is open in explorer it won't be removed.
if not image_directory.exists():
    image_directory.mkdir()


# [(file name, word image, html image)]
results = []

if __name__ == "__main__":
    if len(sys.argv) != 1:
        file_names = [pathlib.Path(p) for p in sys.argv[1:]]
    else:
        file_names = list(pathlib.Path("../docs").iterdir())

    browser = webdriver.Chrome()

    print("Handing {0} files".format(len(file_names)))

    for file in file_names:
        if file.name.endswith(".html"):
            print('- Handling {0}'.format(file.absolute()))

            word_save = image_directory / (file.name + '.png')
            browser_save = image_directory / (file.name + '.html.png')

            subprocess.call(['wordinserter', str(file), '--close', '--save={save}'.format(save=word_save)])

            browser.get('file://{0}'.format(file.absolute()))
            browser.execute_script("document.body.style.margin = '0px'")
            browser.execute_script("document.body.style.padding = '10px'")
            browser.execute_script("document.body.style.display = 'inline-block'")
            body_size = browser.find_element_by_tag_name('body').size
            browser.save_screenshot(str(browser_save))

            with Image(filename=str(browser_save)) as browser_img:
                browser_img.crop(**body_size)
                browser_img.save(filename=str(browser_save))

            results.append((file.name, word_save, browser_save))

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
                <th>Chrome Output</th>
            </thead>
            <tbody>
        """)
        for name, png_path, html_path in results:

            png_path = png_path.relative_to('.').as_posix()
            html_path = html_path.relative_to('.').as_posix()

            fd.write("<tr>")
            fd.write("<td>{0}</td>".format(name))
            fd.write("<td><img src='{0}'></img></td>".format(png_path))
            fd.write("<td><img src='{0}'></img></td>".format(html_path))
            fd.write("</tr>")

        fd.write("</tbody></table>")

        fd.write("</html></body>")

    browser.quit()
    import webbrowser

    webbrowser.open("report.html")
