# Convert Word docx documents into markdown

import argparse
import logging
import os
import re
import sys
import traceback
from shutil import copy2, rmtree
from subprocess import check_call

import bs4
import docsrv
import util
import xlsxwriter
from tabulate import tabulate
from word2mmd import word2mmd

thisdir = os.path.dirname(os.path.realpath(__file__))

pandoc = os.path.join(thisdir, r"..\Pandoc\pandoc.exe")
magick = os.path.join(thisdir, r"..\imagemagick\magick.exe")
mmd2doc = os.path.join(thisdir, r"mmd2doc.py")

# To make imagemagick happy we must set up the MAGICK_CONFIGURE_PATH
my_env = os.environ
magick_path = os.path.abspath(os.path.join(thisdir, r"..\imagemagick"))
my_env["MAGICK_CONFIGURE_PATH"] = "%s;" % magick_path + my_env["PATH"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-5s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger()


class Globals:
    pass


g = Globals()


def setup_parser():
    """
    Set up the argument parser
    """
    parser = argparse.ArgumentParser(
        description="Convert markdown to misc formats build collateral"
    )
    parser.add_argument("source", metavar="*.docx", help="Source file")
    parser.add_argument(
        "-d", "--debug", action="store_true", help="Enable debug output"
    )
    parser.add_argument("--mmd", action="store_true", help="Generate MMD only")
    return parser


def docx2mmd(fnsrc, dirname):
    logger.info("Parsing docx document")
    base_name = os.path.split(fnsrc)[1]
    fnmd = os.path.splitext(base_name)[0] + ".md"
    try:
        word2mmd.convert(fnsrc, fnmd)
    except word2mmd.zipfile.BadZipfile:
        print("Error: couldn't read contents of '%s'. Is the file encrypted?" % fnsrc)
        exit(1)
    with open(fnmd, encoding="utf-8") as markdown_file:
        md_text = markdown_file.read()
    if not g.opts.debug:
        os.unlink(fnmd)
    return md_text


def move_embeddings(md_text, dirname):
    """
    Find all references to the "embeddings" folder, and move them to "assets" instead
    """
    if not os.path.exists("embeddings"):
        # Nothing to do
        return md_text

    logger.info("Moving embeddings")

    moved_files = {}

    def move2assets(m):
        # m - match object from the re.sub; return updated path
        quote, path, filename = m.groups()
        if filename not in moved_files:
            fro = "%s/%s" % (path, filename)
            to = "assets/%s" % (filename)
            if not os.path.exists(fro):
                raise Exception("Referenced file '%s' not found" % fro)
            try:
                if os.path.exists(to):
                    os.unlink(to)
                os.rename(fro, to)
            except OSError:
                logger.warning("Failed to move %s to %s" % (fro, to))
                return m.group(0)
            moved_files[filename] = "%sassets/%s" % (quote, filename)
        return moved_files[filename]

    md_text = re.sub(r'([\'"])(embeddings)/([^\'"]+)', move2assets, md_text)
    md_text = re.sub(r"(\]\()(embeddings)/([^\)]+)", move2assets, md_text)

    try:
        os.rmdir("embeddings")
    except OSError:
        logger.warning("'embeddings' folder still not empty")

    return md_text


def convert_images(md_text, dirname):
    """
    Parse through image references and convert all to png
    .png -> .png
    .emf -> .png
    .tmp -> .png

    It would be much cooler if we could convert EMF to SVG (since EMF is vector format).
    Inkscape doesn't do that (today, it just embeds rendered bitmap in SVG).
    Some references: http://www.imagemagick.org/discourse-server/viewtopic.php?t=28243
    1) MetafileToEPSConverter http://wiki.lyx.org/Windows/MetafileToEPSConverter
       then, ps2pdf (come with Ghostscript), then inkscape.
       See https://github.com/subugoe/edfu-daten/b ... nverter.sh
       (or use pdftocairo comes with poppler, instead of inkscape.)
    2) OpenOffice / LibreOffice + unoconv (http://dag.wiee.rs/home-made/unoconv/)
       Example usage: http://ciarang.com/posts/converting-emf-to-svg
    3) Write a program with libUEMF. http://libuemf.sourceforge.net/
    """

    a = re.split(
        r"(?s)(!\[(?:[^\]]+)?\]\((?:\./)?media/[\w\.]+\.\w+\)(?:\{.*?\})?)", md_text
    )
    docsrv_started = False
    try:
        for i, img_str in enumerate(a):
            # text <img1> text <img2> text ...
            if i % 2 == 0:
                continue
            if i == 1:
                logger.info("Converting images")
            m = re.search(
                r'(?s)media/(.*)\).*width="(\S+?)in.*height="(\S+?)in', img_str
            )
            if m:
                fnimg, w_in, h_in = m.groups(1)
            else:
                fnimg, w_in, h_in = (
                    re.search(r"(?s)media/(.*)\)", img_str).group(1),
                    10,
                    10,
                )
            ppi = 300  # Typical printer
            # Maximum size we need our images to be
            w_px, h_px = int(float(w_in) * ppi), int(float(h_in) * ppi)
            logger.info("Processing %s (%dx%d)" % (fnimg, w_px, h_px))

            fnpng = os.path.splitext(fnimg)[0] + ".png"
            try:
                if re.search(r"(?i)\.[ew]mf$", fnimg):
                    # Use Visio to convert emf/wmf images
                    if not docsrv_started:
                        docsrv.start_session()
                        docsrv_started = True
                    docsrv.submit_job(
                        "VISIO_EXPORT_PAGES_BY_NAME|media/%s|Page-1|assets/%s"
                        % (fnimg, fnpng)
                    )
                    a[i] = re.sub(
                        r"\(media/[^\)]+\)(?:\{.+?\})?", r"(assets/%s)" % fnpng, a[i]
                    )
                    continue
                else:
                    check_call(
                        [
                            magick,
                            "media/%s" % fnimg,
                            "-resize",
                            "%dx%d>" % (w_px, h_px),
                            "assets/%s" % fnpng,
                        ],
                        env=my_env,
                        cwd=dirname,
                    )
            except Exception as e:
                print("-E- Unable to convert %s to png: %s" % (fnimg, e))
                a[i] = "[FIXME - failed to convert %s]()" % fnimg
                continue
            orig_filesz = os.stat(os.path.join(dirname, "media", fnimg)).st_size
            filesz = os.stat(os.path.join(dirname, "assets", fnpng)).st_size
            if (
                fnimg.endswith("png") or fnimg.endswith("tmp")
            ) and filesz > orig_filesz:
                # Just copy over the PNG if size increased
                copy2(
                    os.path.join(dirname, "media", fnimg),
                    os.path.join(dirname, "assets", fnpng),
                )

            # Check if PNG format makes sense for large files. Some are better converted to JPG
            fnout = fnpng
            filesz = os.stat(os.path.join(dirname, "assets", fnpng)).st_size
            ii = (
                util.check_output_text(
                    [magick, "identify", "assets/%s" % fnpng], env=my_env, cwd=dirname
                )
                .strip()
                .split()
            )
            # image151.png PNG 3000x2250 3000x2250+0+0 8-bit sRGB 2.145MB 0.000u 0:00.000
            w, h = list(map(int, ii[2].split("x")))
            n_bytes_raw = w * h * 3
            comp_ratio = 1.0 * n_bytes_raw / filesz
            logger.debug(
                "%s: %dkB, compression ratio: %.1f" % (fnpng, filesz / 1024, comp_ratio)
            )
            if filesz > 100000 and comp_ratio < 50:
                # Big poorly compressed file - see if that's justified
                logger.debug("Candidate image for compression: %s" % fnpng)
                fnjpg = os.path.splitext(fnimg)[0] + ".jpg"
                check_call(
                    [
                        magick,
                        "assets/%s" % fnpng,
                        "-resize",
                        "1200>",
                        "assets/%s" % fnjpg,
                    ],
                    cwd=dirname,
                )
                jpg_filesz = os.stat(os.path.join(dirname, "assets", fnjpg)).st_size
                jpg_png_ratio = 1.0 * filesz / jpg_filesz
                logger.debug(
                    "JPEG to PNG size ratio (more is better): %.1f" % jpg_png_ratio
                )
                if jpg_png_ratio < 1.5:
                    # require at least 1.5 reduction, otherwise pointless
                    os.unlink(os.path.join(dirname, "assets", fnjpg))
                else:
                    # reduced size significantly -> use JPG
                    os.unlink(os.path.join(dirname, "assets", fnpng))
                    fnout = fnjpg
            a[i] = re.sub(r"\(media/[^\)]+\)(?:\{.+?\})?", r"(assets/%s)" % fnout, a[i])
    finally:
        if docsrv_started:
            docsrv.end_session()

    result = "".join(a)
    m = re.search(r"(?m)^(.*\(media/.*\.png\)).*$", result)
    if m:
        print("ERROR: not converted image\n%s" % m.group(1))
    elif not g.opts.debug:
        # Remove the media directory
        if os.path.exists(os.path.join(dirname, "media")):
            try:
                rmtree(os.path.join(dirname, "media"))
            except BaseException:
                logger.warning("Attempted to remove 'media' but failed.  Skipping")
                pass

    return result


def clean_utf8(s):
    s = s.replace("&#160;", " ")  # Non-breaking space
    s = s.replace("&#169;", "&copy;")  # COPYRIGHT SIGN
    s = s.replace("&#174;", "&reg;")  # REGISTERED SIGN
    s = s.replace("&#8216;", "'")  # LEFT SINGLE QUOTATION MARK
    s = s.replace("&#8217;", "'")  # RIGHT SINGLE QUOTATION MARK
    s = s.replace("&#8209;", "-")  # NON-BREAKING HYPHEN
    s = s.replace("&#8211;", "-")  # EN DASH
    s = s.replace("&#8220;", '"')  # LEFT DOUBLE QUOTATION MARK
    s = s.replace("&#8221;", '"')  # RIGHT DOUBLE QUOTATION MARK
    s = s.replace("&#8230;", "...")  # HORIZONTAL ELLIPSIS
    s = s.replace("\r\n", "\n")  # Windows -> Unix
    s = s.replace("\r\n", "\n\n")  # Second pass for \r\r\n cases
    return s


def clean_tags(s):
    cleanup_needed = 1
    while cleanup_needed:
        # Remove empty spans
        s, cleanup_needed = re.subn(r"(?ms)<span[^>]*>\s*</span>", "", s)
    # clean up empty comments (they serve some purpose but all
    # usages I've seen were bogus)
    s = re.sub(r"(?sm)<!--\s*-->", "", s)
    return s


def clean_backslashes(s):
    # Pandoc is smart enough to ignore underscores, bracketed indices, etc.
    # So for readability, we'll remove the escaping
    a = re.split(r"(\$\$.*?\$\$)", s)  # (except for undescores in formulas)
    a = [re.sub(r"(\w)\\_", r"\1_", x) if i % 2 == 0 else x for i, x in enumerate(a)]
    s = "".join(a)
    s = re.sub(r"\\\[(\d+\:\d+)\\\]", r"[\1]", s)
    s = re.sub(r"\\~", r"~", s)
    return s


def clean_markdown(s):
    # Remove empty list items
    s = re.sub(r"(?m)^\d+\.\s*$", "", s)
    s = re.sub(r"(?m)^-\s*$", "", s)
    # Remove empty/bogus tables
    s = re.sub(r"(?ms)^\s*\n(?:\s+-+)+\n\s*\n(?:\s+-+)+\n\s*$", "", s)
    # clean up spaces at EOL
    s = re.sub(r"(?m)\s+\n", "\n", s)
    # Collapse adjacent empty lines
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def tablefix(md_text, basename):
    """
    Parse a markdown file, extract HTML tables and convert to either
    markdown or xls as appropriate
    """

    token_specification = []
    # TODO: regex to tolerate nested tables
    token_specification.append(("HTML_TABLE", "<table>.*?</table>"))
    token_specification.append(("TEXT", r"."))  # anything else
    T = util.Tokenizer(token_specification)

    output_text = ""

    xlpath = "assets/%s.xlsx" % basename

    def init_xls_wb():
        logger.info("Converting tables to %s.xlsx" % basename)
        wb = xlsxwriter.Workbook(xlpath, {"strings_to_formulas": False})
        formats = {
            "merge": wb.add_format(
                {"align": "center", "border": True, "text_wrap": True}
            ),
            "merge_bold": wb.add_format(
                {"align": "center", "bold": True, "border": True, "text_wrap": True}
            ),
            "cell": wb.add_format({"align": "left", "border": True, "text_wrap": True}),
            "cell_bold": wb.add_format(
                {"align": "left", "bold": True, "border": True, "text_wrap": True}
            ),
            "header": wb.add_format(
                {"align": "center", "bold": True, "border": True, "text_wrap": True}
            ),
        }
        return wb, formats

    wb = None
    for token in T.tokenize(md_text):
        if token.typ == "HTML_TABLE":
            if not wb:
                wb, formats = init_xls_wb()
            table = html2table(token.value, wb, formats)
            if isinstance(table, str):
                output_text += "\n"
                output_text += table
                output_text += "\n"
            else:
                # Must have been xlsx
                sheetname = table.get_name()
                caption = table.full_name.replace('"', "'")
                output_text += '\n```xls("%s", "%s", "%s")\n```\n\n' % (
                    xlpath,
                    sheetname,
                    caption,
                )
        else:
            output_text += token.value

    if wb:
        wb.close()
    return output_text


class TableBetterForExcel(Exception):
    pass


def html2table(html, wb, formats):
    """
    Parse HTML and write to xlsx worksheet
    """
    s = bs4.BeautifulSoup(html, "html.parser")

    # First look to see if this is a simple table.  if it is, return markdown
    # - no newlines
    # - less than 100 chars
    try:
        table = markdown_table(s)
        return table
    except TableBetterForExcel:
        pass

    # Use table caption as page name, when possible
    if s.caption:
        if not hasattr(wb, "used_sheet_names"):
            wb.used_sheet_names = set()
        # Using 'text' instead of 'string' because string returns None if it is
        # wrapped by <br>...</br>
        caption = re.sub(r"[^\w ]", "", s.caption.text).strip()[:31].strip()
        n, orig_caption = 1, caption
        while caption.lower() in wb.used_sheet_names:
            caption = orig_caption[:28] + "_%d" % n
            n += 1
        wb.used_sheet_names.add(caption.lower())
        sheet = wb.add_worksheet(caption)
        sheet.full_name = s.caption.text.strip()
    else:
        sheet = wb.add_worksheet()
        sheet.full_name = sheet.get_name()

    # Gather column info first
    colwidth = write_cells(sheet, s, formats)
    colwidth = fit_columns(sheet, colwidth)

    return sheet


def fit_columns(sheet, colwidth):
    """
    Simple algorithm:
    Target 200 characters at most
    Take the largest columns and reduce by 30%
    Repeat until we meet 200 char limit
    """
    colsum = 9999
    while colsum > 190:
        maxcol = 0
        maxcol_i = -1
        for i in range(len(colwidth)):
            if colwidth[i] > maxcol:
                maxcol = colwidth[i]
                maxcol_i = i
        threshold = maxcol * 0.9
        for i in range(len(colwidth)):
            if colwidth[i] >= threshold:
                colwidth[i] = colwidth[i] * 0.7

        colsum = 0
        for c in colwidth:
            colsum += c

    for i in range(len(colwidth)):
        # Add char width to the columns to cover for some variation in character
        # widths.
        sheet.set_column(i, i, colwidth[i] + 1)


def bold2style(cell_text):
    # Remove markdown bold formatting (**), do it in XL style instead
    bold = ""
    ans = re.sub(r"(\*{2,})\n\1", "\n", cell_text)
    m = re.search(r"(?s)^\s*(\*{2,})\s*(.*)\s*\1\s*$", ans)
    if m and "**" not in m.group(2):
        ans = m.group(2)
        bold = "_bold"
    return bold, ans


def write_cells(sheet, s, formats):
    """
    Walk html table and write out cells to excel sheet
    """
    # Track column widths if colwidth is undefined
    colwidth = {}
    y = 0

    skip_merged = set()

    for row in s.findAll("tr"):
        header = False
        cols = row.findAll("td")
        if len(cols) == 0:
            header = True
            cols = row.findAll("th")
            if len(cols) == 0:
                continue
        x = 0
        for data in cols:

            # skip the cells that we have merged, they won't have a corresponding <td>
            while (x, y) in skip_merged:
                x += 1

            # Remove javascript from cells (sometimes used for hyperlinks like the following):
            # <script type="text/javascript">
            # <!--
            # h='&#48;&#46;&#x39;&#x41;';a='&#64;';n='&#56;&#48;&#x25;&#50;&#x35;';e=n+a+h;
            # document.write('<a h'+'ref'+'="ma'+'ilto'+':'+e+'" clas'+'s="em' + 'ail">');
            # // -->
            # </script>
            # <noscript>80%@ 0.9A (80%25 at 0 dot 9A)</noscript>
            for script in data.find_all("script"):
                script.decompose()

            cell_text = data.text.strip()

            # Right thing to do would be to parse markdown here
            # (since the input is html-wrapped markdown)
            # For now, just do some minimal stuff
            cell_text = re.sub(r"\n{2,}", "\n", cell_text)
            bold, cell_text = bold2style(cell_text)

            # Write our cell
            colspan = int(data.attrs.get("colspan", 1))
            rowspan = int(data.attrs.get("rowspan", 1))
            if colspan + rowspan > 2:
                sheet.merge_range(
                    y,
                    x,
                    y + rowspan - 1,
                    x + colspan - 1,
                    cell_text,
                    formats["merge%s" % bold],
                )
                # Mark merged cells to skip them when walking <td> elements
                for rs in range(rowspan):
                    for cs in range(colspan):
                        if rs + cs == 0:
                            continue
                        skip_merged.add((x + cs, y + rs))

                # Find the longest line and divide by the number of columns
                width = len(max(data.text.split("\n"), key=len))
                for cs in range(colspan):
                    colwidth[x + cs] = max(colwidth.get(x + cs, 0), width / colspan)
            else:
                # Trick -- the max function will find the longest full line in the list and then the
                # len() wrapper finds the character count.
                colwidth[x] = max(
                    colwidth.get(x, 0), len(max(cell_text.split("\n"), key=len))
                )

                if header:
                    sheet.write(y, x, cell_text, formats["header"])
                else:
                    sheet.write(y, x, cell_text, formats["cell%s" % bold])
            x += 1
        y += 1

    num_cols = max(colwidth.keys()) + 1
    return [colwidth.get(i, 0) for i in range(num_cols)]


def markdown_table(s):
    """
    Try to convert to a simple table
    """
    header = []
    table = []
    rowlen = 0
    for row in s.findAll("tr"):
        isheader = False
        cols = row.findAll("td")
        if len(cols) == 0:
            isheader = True
            cols = row.findAll("th")
            if len(cols) == 0:
                continue
        x = 0
        row = []
        for data in cols:
            rowspan = 1
            colspan = 1
            cell = data.text.strip()
            # Escape html when embedded in the table
            cell = cell.replace("<", "&lt;")
            if "colspan" in data.attrs:
                colspan = int(data.attrs["colspan"])
            if "rowspan" in data.attrs:
                rowspan = int(data.attrs["rowspan"])
            if rowspan + colspan > 2:
                raise TableBetterForExcel

            if cell.find("\n") >= 0:
                # No newlines allowed
                raise TableBetterForExcel

            if cell == "":
                cell = "&nbsp;"
            row.append(cell)
        if isheader:
            header = row
        else:
            table.append(row)

        # check max row size
        if len("".join(row)) > 100:
            raise TableBetterForExcel

    ans = tabulate(table, headers=header)
    if s.caption:
        ans = ans.rstrip() + "\n\nTable: %s\n" % s.caption.text.strip()
    return ans


if __name__ == "__main__":
    try:
        # log usage details
        util.log_app_details_async(command=" ".join(sys.argv[0:]))

        parser = setup_parser()
        g.opts = parser.parse_args(sys.argv[1:])
        if not g.opts.source.lower().endswith(".docx"):
            raise Exception("Input file name should end with .docx")
        if g.opts.debug:
            logger.setLevel(logging.DEBUG)

        source_path = os.path.abspath(g.opts.source)
        abspath_src, fn_src = os.path.split(source_path)

        # Create new output directory (clean up existing one first, if needed)
        base_name = re.sub("\..*", "", fn_src)
        assert base_name
        output_path = os.path.join(abspath_src, base_name)
        logger.info("Creating output directory '%s'" % output_path)
        if os.path.exists(output_path):
            rmtree(output_path)
        os.makedirs(output_path)
        os.chdir(output_path)
        fn_src = "../" + fn_src

        # Make the assets directory if needed
        os.makedirs("assets")

        md_text = docx2mmd(source_path, output_path)
        md_text = convert_images(md_text, output_path)
        md_text = move_embeddings(md_text, output_path)
        md_text = clean_utf8(md_text)
        md_text = clean_tags(md_text)
        md_text = clean_backslashes(md_text)

        # Post-process the markdown text and replace tables
        md_text = tablefix(md_text, base_name)

        # Save output markdown file
        logger.info("Saving markdown")
        fnout = base_name + ".mmd"
        with open(fnout, "w", encoding="utf-8") as markdown_file:
            markdown_file.write(md_text)

        # Try compiling to HTML
        if not g.opts.mmd:
            logger.info("Compiling to HTML")
            check_call([sys.executable, mmd2doc, fnout, "--chrome"])
    except Exception as e:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        util.log_app_details_async(
            command=" ".join(sys.argv[0:]), exc_traceback=exc_traceback, message=repr(e)
        )
        traceback.print_exc()
        exit(1)

    # Remove the assets directory if empty
    try:
        os.rmdir("assets")
    except OSError:
        pass
