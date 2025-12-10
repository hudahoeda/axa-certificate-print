# import codecs
import json
from io import BytesIO

import cairosvg
import requests
from bs4 import BeautifulSoup as BS
from PIL import Image


def run():
    cert_name = "data-analytics"
    title = "RevoU Data Analytics Certificate"

    # save in the following file the copied html obtained from the canva embed using Chrome inspect view
    html_filename = "source.html"

    with open(html_filename, "r") as fp:
        html = fp.read()

    soup = BS(html, "html.parser")

    # Find all css and save them in a single file
    all_css = ""
    for tag in soup.find_all("link", rel="stylesheet"):
        cssurl = tag["href"]
        try:
            response = requests.get(cssurl)
        except Exception as e:
            print(f"failed url:{cssurl} - {e}")
            continue
        if response.ok:
            all_css += response.text

    with open(f"style-{cert_name}.css", "w") as f:
        f.write(all_css)

        # remove useless nodes and add fonts to CSS
        soup.find("div", {"role": "navigation"}).decompose()
        for div in soup.find_all("button"):
            div.decompose()
        for div in soup.find_all("a"):
            div.decompose()
        for div in soup.find_all("script"):
            if "window['bootstrap']" in div.text:
                json_content = div.text.split("'")[7]
                for font in json.loads(json_content)["page"]["Bj"]["A"]["H"]:
                    # for font in json.loads(
                    #     codecs.getdecoder("unicode_escape")(json_content)[0]
                    # )["page"]["Bj"]["A"]["H"]:
                    name = f'{font["A"]} {font["B"]}'
                    for style in font["D"]:
                        style_name = style["style"]
                        try:
                            file1_url = style["files"][0]["url"]
                            file1_format = (
                                style["files"][0]["format"]
                                .lower()
                                .replace("otf", "opentype")
                                .replace("ttf", "truetype")
                            )
                            file2_url = style["files"][1]["url"]
                            file2_format = (
                                style["files"][1]["format"]
                                .lower()
                                .replace("otf", "opentype")
                                .replace("ttf", "truetype")
                            )
                            file3_url = style["files"][2]["url"]
                            file3_format = (
                                style["files"][2]["format"]
                                .lower()
                                .replace("otf", "opentype")
                                .replace("ttf", "truetype")
                            )
                        except Exception as e:
                            print(f"error for style: {style}:", e)
                            file1_url = ""
                            file2_url = ""
                            file3_url = ""
                            file1_format = ""
                            file2_format = ""
                            file3_format = ""

                        if style_name == "BOLD":
                            font_style = "font-weight: 700"
                        elif style_name == "ITALICS":
                            font_style = "font-style: italic"
                        elif style_name == "REGULAR":
                            font_style = "font-weight: 400; font-style: normal"
                        elif style_name == "BOLD_ITALICS":
                            font_style = "font-weight: 700; font-style: italic"
                        row = f'\n@font-face {{font-family: "{name}"; {font_style}; '
                        if file1_url and file1_format:
                            row += f'src: url({file1_url}) format("{file1_format}")'
                        if file2_url and file2_format:
                            row += f', url({file2_url}) format("{file2_format}")'
                        if file3_url and file3_format:
                            row += f', url({file3_url}) format("{file3_format}")'
                        row = f"{row}}}"
                        f.write(row)

            div.decompose()
    # for div in soup.find_all("div", {"aria-hidden": "true"}):
    #     div.decompose()

    # download all the images and rename them
    for num, img in enumerate(soup.find_all("img")):
        image_url = img["src"]
        ext = image_url.split(".")[-1].split("?")[0]

        img_new_name = f"image-{num}.{ext}"
        img["src"] = img_new_name
        if ext == "svg":
            img_new_name = f"image-{num}.png"
            img["src"] = img_new_name
            print(num, image_url, img_new_name)

            out = BytesIO()
            try:
                cairosvg.svg2png(url=image_url, write_to=out)
                imgfile = Image.open(out)
            except Exception as e:
                print(f"error reading SVG image:{e}")
            else:
                imgfile.save(img_new_name, quality=90)
        else:
            print(num, image_url, img_new_name)
            try:
                imgfile = Image.open(requests.get(image_url, stream=True).raw)
            except Exception as e:
                print(f"error reading image:{e}")
            else:
                imgfile.save(img_new_name, quality=90)

    # print out the final template
    head = f"""
      <meta charset="utf-8">
      <base href="https://revou-assets.s3-ap-southeast-1.amazonaws.com/certificates-static/{cert_name}/">
      <title>{title}</title>
      <meta name="app-name" content="design_viewer">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <meta name="author" content="RevoU"/>
      <meta name="subject" content="The RevoU {{{{program}}}} certificate for {{{{first_name}}}} {{{{last_name}}}}"/>
      <meta name="date" content="{{{{issue_date|date:'Y-m-d'}}}}"/>
      <link href="style-{cert_name}.css" rel="stylesheet">
      <style>@page {{ size: A4 landscape; margin: 0; }}</style>"""

    new_html = f"""<html dir="ltr" lang="en" class="EHoceA">
   <head>    
   {head}
   </head>
   {soup.body.prettify()}
</html>"""
    with open(f"canva-{cert_name}.html", "w") as f:
        f.write(new_html)


if __name__ == "__main__":
    run()